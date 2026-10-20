/* Pure functions shared by Apps Script and the offline test harness. */
var YTCore = (function () {
  'use strict';
  var VIDEO = /^[A-Za-z0-9_-]{11}$/;
  function fail(message) { throw new Error(message); }
  function inputUrl(value) {
    var raw = String(value || '').trim();
    if (!/^https?:\/\//i.test(raw)) raw = 'https://' + raw;
    // Apps Script does not provide the browser URL constructor.
    var m = raw.match(/^https?:\/\/([^/?#]+)([^?#]*)(?:\?([^#]*))?(?:#.*)?$/i);
    if (!m || !/^(?:www\.|m\.|music\.)?(?:youtube\.com|youtu\.be)$/i.test(m[1])) return null;
    var query = Object.create(null);
    (m[3] || '').split('&').forEach(function (part) {
      var x = part.split('=');
      try { query[decodeURIComponent(x[0])] = decodeURIComponent((x[1] || '').replace(/\+/g, ' ')); } catch (_) {}
    });
    return { host: m[1].toLowerCase(), path: m[2], query: query };
  }
  function videoId(value) {
    var raw = String(value || '').trim();
    if (VIDEO.test(raw)) return raw;
    var u = inputUrl(raw); if (!u) return null;
    var parts = u.path.split('/').filter(Boolean), id;
    if (u.host === 'youtu.be' || u.host === 'www.youtu.be') id = parts[0];
    else if (u.path === '/watch') id = u.query.v;
    else if (['shorts', 'embed', 'live'].indexOf(parts[0]) >= 0) id = parts[1];
    return VIDEO.test(id || '') ? id : null;
  }
  function playlistId(value) {
    var raw = String(value || '').trim(), u = inputUrl(raw);
    var id = u ? u.query.list : raw;
    return /^[A-Za-z0-9_-]{12,150}$/.test(id || '') && !/^UC[A-Za-z0-9_-]{22}$/.test(id) ? id : null;
  }
  function channelRequest(value) {
    var raw = String(value || '').trim(), u;
    if (/^UC[A-Za-z0-9_-]{22}$/.test(raw)) return { id: raw };
    if (/^@[^\s/?#]{1,100}$/u.test(raw)) return { forHandle: raw };
    u = inputUrl(raw); if (!u) return null;
    var p = u.path.split('/').filter(Boolean);
    if (p[0] === 'channel' && /^UC[A-Za-z0-9_-]{22}$/.test(p[1] || '')) return { id: p[1] };
    if (/^@[^\s/?#]{1,100}$/u.test(p[0] || '')) return { forHandle: p[0] };
    if (p[0] === 'user' && /^[A-Za-z0-9_-]+$/.test(p[1] || '')) return { forUsername: p[1] };
    return null;
  }
  function dateBoundary(value, end) {
    if (!value) return null;
    var m = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) fail('Use dates in YYYY-MM-DD format.');
    var d = new Date(value + 'T00:00:00.000Z');
    if (!isFinite(d.getTime()) || d.toISOString().slice(0, 10) !== value) fail('Choose a valid calendar date.');
    return d.getTime() + (end ? 86400000 : 0);
  }
  function dateRange(start, end) {
    var a = dateBoundary(start, false), b = dateBoundary(end, true);
    if (a !== null && b !== null && a >= b) fail('Start date must be on or before end date.');
    return { start: a, end: b };
  }
  function inRange(value, range) {
    if (range.start === null && range.end === null) return true;
    var t = Date.parse(value || '');
    return isFinite(t) && (range.start === null || t >= range.start) && (range.end === null || t < range.end);
  }
  function durationSeconds(value) {
    var m = String(value || '').match(/^P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?)?$/);
    if (!m || !m.slice(1).some(Boolean)) return null;
    return Number(m[1] || 0)*86400 + Number(m[2] || 0)*3600 + Number(m[3] || 0)*60 + Number(m[4] || 0);
  }
  function longVideoDecision(item) {
    var snippet = item && item.snippet || {};
    if (snippet.liveBroadcastContent === 'upcoming' || snippet.liveBroadcastContent === 'live') {
      return { include: false, reason: 'Duration is not final for a live or upcoming broadcast.' };
    }
    var seconds = durationSeconds(item && item.contentDetails && item.contentDetails.duration);
    if (seconds === null) return { include: false, reason: 'Duration unavailable. Cannot verify the video exceeds three minutes.' };
    return { include: seconds > 180, reason: seconds > 180 ? '' : 'Excluded by duration safeguard: video is three minutes or shorter. This does not establish YouTube Shorts classification.' };
  }
  function count(value) {
    if (value === undefined || value === null || value === '') return '';
    var s = String(value);
    if (!/^\d+$/.test(s)) return '';
    var n = Number(s); return Number.isSafeInteger(n) ? n : s;
  }
  function safeCell(value) {
    if (typeof value !== 'string') return value;
    var s = value.slice(0, 45000);
    return /^[\s\u0000-\u001f]*[=+@-]/.test(s) ? "'" + s : s;
  }
  function unique(values) { return Array.from(new Set(values)); }
  function chunks(values, size) {
    var out=[];for(var i=0;i<values.length;i+=size) out.push(values.slice(i,i+size));return out;
  }
  function classify(error) {
    var text = String(error && error.message || error);
    var detail = error && error.details ? JSON.stringify(error.details) : '';
    var all = text + ' ' + detail;
    if (/quotaExceeded|dailyLimitExceeded|exceeded your quota/i.test(all)) return 'QUOTA';
    if (/rateLimitExceeded|userRateLimitExceeded|backendError|internalError|temporarily|\b(429|500|502|503|504)\b/i.test(all)) return 'TRANSIENT';
    if (/forbidden|unauthorized|insufficientPermissions|accessNotConfigured|\b(401|403)\b/i.test(all)) return 'AUTH';
    return 'PERMANENT';
  }
  function parseInputs(text) {
    var tokens = String(text || '').split(/[\s,;]+/).filter(Boolean);
    if (!tokens.length) fail('Paste at least one video URL or ID.');
    if (tokens.length > 5000) fail('Use at most 5,000 video inputs per job.');
    return tokens.map(function (value) { return { input: value, id: videoId(value) }; });
  }
  function apiBatches(ids) { return chunks(unique(ids.filter(Boolean)), 50); }
  return { videoId:videoId, playlistId:playlistId, channelRequest:channelRequest, dateRange:dateRange,
    inRange:inRange, durationSeconds:durationSeconds, longVideoDecision:longVideoDecision, count:count, safeCell:safeCell, unique:unique,
    chunks:chunks, classify:classify, parseInputs:parseInputs, apiBatches:apiBatches };
})();
