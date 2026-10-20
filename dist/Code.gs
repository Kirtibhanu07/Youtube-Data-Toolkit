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

/* Bound Google Sheets adapter. All external writes target a new job-owned tab. */
var YT_HEADERS = ['Input','Video ID','Status','Title','Channel','Channel ID','Published UTC','Published local','Duration seconds','Duration','Views','Likes','Comments','Broadcast status','Scheduled start UTC','Actual start UTC','Country code','Description','Retrieved UTC','Error'];
var YT_TERMINAL = ['COMPLETE','COMPLETE_WITH_ISSUES','CANCELLED'];
function onOpen() {
  SpreadsheetApp.getUi().createMenu('YouTube Data Toolkit')
    .addItem('Get metadata from video URLs', 'fetchYoutubeDetails')
    .addItem('Retry unavailable videos', 'retryFailedRows')
    .addSeparator()
    .addItem('Import playlist videos', 'importFromPlaylist')
    .addItem('Import channel uploads', 'importFromChannel')
    .addItem('Index channel playlists', 'getAllPlaylistsFromChannel')
    .addSeparator()
    .addItem('Open workspace / resume collection', 'openToolkit').addToUi();
}
function openToolkit() { openToolkitMode_(''); }
function fetchYoutubeDetails() { openToolkitMode_('videos'); }
function importFromPlaylist() { openToolkitMode_('playlist'); }
function importFromChannel() { openToolkitMode_('channel'); }
function getAllPlaylistsFromChannel() { openToolkitMode_('playlists'); }
function retryFailedRows() {
  try { var j=ytLoad_();if(!j)throw new Error('Start a collection first.');toolkitRetry(j.id,false); }
  catch(e){SpreadsheetApp.getUi().alert(String(e.message||e));}
  openToolkit();
}
function openToolkitMode_(mode) {
  var html=HtmlService.createHtmlOutputFromFile('Sidebar').getContent();
  html=html.replace('<!--INITIAL_MODE-->','<script>window.YT_INITIAL_MODE='+JSON.stringify(mode)+';</script>');
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle('YouTube Data Toolkit'));
}
function ytKey_() { return 'youtube.toolkit.v2.' + SpreadsheetApp.getActiveSpreadsheet().getId(); }
function ytLoad_() { var s=PropertiesService.getUserProperties().getProperty(ytKey_());return s?JSON.parse(s):null; }
function ytSave_(j) { j.updatedAt=new Date().toISOString();PropertiesService.getUserProperties().setProperty(ytKey_(),JSON.stringify(j)); }
function ytLocked_(work) {
  var lock=LockService.getUserLock();
  if (!lock.tryLock(1000)) throw new Error('Another action is running. Wait a moment and try again.');
  try { return work(); } finally { lock.releaseLock(); }
}
function ytSheet_(j) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(ss.getId()!==j.spreadsheetId) throw new Error('Open the spreadsheet where this job started.');
  var sheet=ss.getSheetById(j.sheetId);
  if(!sheet) throw new Error('The output tab was removed. Start a new job.');
  var h=sheet.getRange(1,1,1,YT_HEADERS.length).getValues()[0];
  if(h.join('\t')!==YT_HEADERS.join('\t')) throw new Error('Output headers changed. Restore the headers before continuing.');
  return sheet;
}
function ytRows_(sheet) { return sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,YT_HEADERS.length).getValues():[]; }
function ytEmptyRow_(input,id) { var row=new Array(YT_HEADERS.length).fill('');row[0]=YTCore.safeCell(input);row[1]=id||'';row[2]=id?'PENDING':'INVALID';if(!id)row[19]='Not a supported YouTube video URL or ID';return row; }
function ytWrite_(sheet,start,rows) {
  if(!rows.length)return;
  var needed=start+rows.length-1-sheet.getMaxRows();if(needed>0)sheet.insertRowsAfter(sheet.getMaxRows(),needed);
  sheet.getRange(start,1,rows.length,YT_HEADERS.length).setValues(rows.map(function(r){return r.map(YTCore.safeCell);}));
}
function ytNewSheet_(j) {
  var sheet=SpreadsheetApp.getActiveSpreadsheet().insertSheet('YT '+Utilities.formatDate(new Date(),'UTC','MMdd HHmm')+' '+j.id.slice(0,6));
  if(sheet.getMaxColumns()<YT_HEADERS.length) sheet.insertColumnsAfter(sheet.getMaxColumns(),YT_HEADERS.length-sheet.getMaxColumns());
  sheet.getRange(1,1,1,YT_HEADERS.length).setValues([YT_HEADERS]).setBackground('#163a42').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);sheet.setColumnWidths(1,YT_HEADERS.length,150);sheet.setColumnWidth(4,300);sheet.setColumnWidth(18,320);
  sheet.getRange(1,1).setNote('YouTube Toolkit 2 output. Keep headers unchanged. Metadata is a snapshot, not a live feed.');
  return sheet;
}
function ytApi_(j, resource, part, options) {
  for(var attempt=0;attempt<3;attempt++) {
    // Every attempted list request is counted, including failed requests and retries.
    j.apiAttempts++;ytSave_(j);
    try { return YouTube[resource].list(part,options); }
    catch(e) {
      var type=YTCore.classify(e);
      if(type!=='TRANSIENT'||attempt===2) throw e;
      Utilities.sleep(Math.min(4000,500*Math.pow(2,attempt))+Math.floor(Math.random()*200));
    }
  }
}
function toolkitSelection() {
  var range=SpreadsheetApp.getActiveRange();
  if(!range)throw new Error('Select a column of video URLs first.');
  if(range.getNumColumns()!==1)throw new Error('Select one column containing video URLs.');
  if(range.getNumRows()>5000)throw new Error('Select at most 5,000 rows.');
  return range.getDisplayValues().map(function(r){return r[0];}).filter(Boolean).join('\n');
}
function toolkitStart(payload) {
  return ytLocked_(function(){
    payload=payload||{};var previous=ytLoad_();
    if(previous&&YT_TERMINAL.indexOf(previous.state)<0)throw new Error('Finish or cancel your current job before starting another.');
    var mode=payload.mode;
    if(['videos','playlist','channel','playlists'].indexOf(mode)<0)throw new Error('Choose a supported import mode.');
    var range=YTCore.dateRange(payload.startDate,payload.endDate), parsed,source,channel;
    if(mode==='videos')parsed=YTCore.parseInputs(payload.source);
    else if(mode==='playlist') {source=YTCore.playlistId(payload.source);if(!source)throw new Error('Enter a valid YouTube playlist URL or ID.');}
    else {channel=YTCore.channelRequest(payload.source);if(!channel)throw new Error('Enter a channel handle, channel ID or supported channel URL.');}
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var j={id:Utilities.getUuid(),spreadsheetId:ss.getId(),mode:mode,source:source||'',channelRequest:channel||null,
      dateRange:range,excludeShorts:payload.excludeShorts!==false,excludedShort:0,unverifiedDuration:0,phase:mode==='videos'?'METADATA':(mode==='playlist'?'COLLECT':'RESOLVE'),pageToken:'',state:'READY',
      apiAttempts:0,pages:0,scanned:0,filtered:0,unknownDates:0,duplicates:0,createdAt:new Date().toISOString(),message:'Ready to collect.'};
    var sheet=ytNewSheet_(j);j.sheetId=sheet.getSheetId();
    if(parsed){ytWrite_(sheet,2,parsed.map(function(x){return ytEmptyRow_(x.input,x.id);}));}
    ytSave_(j);return ytState_(j);
  });
}
function ytState_(j) {
  if(!j)return {state:'EMPTY',message:'Start a new collection.'};
  var rows=ytRows_(ytSheet_(j)),counts={};rows.forEach(function(r){counts[r[2]]=(counts[r[2]]||0)+1;});
  return {id:j.id,state:j.state,phase:j.phase,mode:j.mode,counts:counts,total:rows.length,
    apiAttempts:j.apiAttempts,pages:j.pages,scanned:j.scanned,filtered:j.filtered,unknownDates:j.unknownDates,duplicates:j.duplicates,excludedShort:j.excludedShort||0,unverifiedDuration:j.unverifiedDuration||0,excludeShorts:j.excludeShorts,
    message:j.message,updatedAt:j.updatedAt,sheetUrl:SpreadsheetApp.getActiveSpreadsheet().getUrl()+'#gid='+j.sheetId};
}
function toolkitStatus() { return ytLocked_(function(){return ytState_(ytLoad_());}); }
function ytRequireJob_(id) {var j=ytLoad_();if(!j||j.id!==id)throw new Error('This job changed. Reload its status.');return j;}
function toolkitCancel(id) {return ytLocked_(function(){var j=ytRequireJob_(id);j.state='CANCELLED';j.message='Cancelled. Collected rows remain in the output tab.';ytSave_(j);return ytState_(j);});}
function toolkitRun(id) {
  return ytLocked_(function(){
    var j=ytRequireJob_(id);if(YT_TERMINAL.indexOf(j.state)>=0)return ytState_(j);
    var sheet=ytSheet_(j);j.state='READY';j.message='Processing';
    try {
      if(j.phase==='RESOLVE'){
        var response=ytApi_(j,'Channels','snippet,contentDetails',j.channelRequest),channel=(response.items||[])[0];
        if(!channel)throw new Error('Channel not found or unavailable to this account.');
        j.channelId=channel.id;j.source=j.mode==='playlists'?channel.id:channel.contentDetails.relatedPlaylists.uploads;j.phase='COLLECT';
      } else if(j.phase==='COLLECT') ytCollect_(j,sheet);
      else ytMetadata_(j,sheet);
      if(j.state==='READY')j.message=j.phase==='METADATA'?'Collecting video details.':'Collecting the next page.';
    }catch(e){j.state='PAUSED';j.message=YTCore.classify(e)+': '+String(e.message||e).slice(0,300);}
    ytSave_(j);return ytState_(j);
  });
}
function ytCollect_(j,sheet) {
  var index=j.mode==='playlists',options={maxResults:50};
  options[index?'channelId':'playlistId']=j.source;if(j.pageToken)options.pageToken=j.pageToken;
  var response=ytApi_(j,index?'Playlists':'PlaylistItems','snippet,contentDetails',options);
  if(response.nextPageToken&&response.nextPageToken===j.pageToken)throw new Error('Repeated page token. Resume later or restart the import.');
  var rows=ytRows_(sheet),seen=new Set(rows.map(function(r){return r[1];})),added=[];
  (response.items||[]).forEach(function(item){
    j.scanned++;
    var cd=item.contentDetails||{},sn=item.snippet||{},id=index?item.id:(cd.videoId||(sn.resourceId||{}).videoId);
    if(!id){j.filtered++;return;}
    if(seen.has(id)){j.duplicates++;return;}
    // Video publication time differs from the time a video was added to a playlist.
    var published=index?sn.publishedAt:cd.videoPublishedAt;
    if(!YTCore.inRange(published,j.dateRange)){j.filtered++;if(!published)j.unknownDates++;return;}
    var row=ytEmptyRow_(index?'https://www.youtube.com/playlist?list='+id:'https://www.youtube.com/watch?v='+id,id);
    row[3]=sn.title||'';row[6]=published||'';
    if(index){row[2]='PLAYLIST';row[4]=sn.channelTitle||'';row[5]=sn.channelId||'';row[17]=sn.description||'';row[18]=new Date().toISOString();}
    added.push(row);seen.add(id);
  });
  if(rows.length+added.length>5000){j.state='PAUSED';j.message='The 5,000-row limit was reached. Start a narrower import.';return;}
  if (!index && j.excludeShorts && added.length) {
    // PlaylistItems has no duration. Check candidate video durations in one list request.
    var details=ytApi_(j,'Videos','snippet,contentDetails',{id:added.map(function(r){return r[1];}).join(',')}),byId=Object.create(null);
    (details.items||[]).forEach(function(item){byId[item.id]=item;});
    added=added.filter(function(row){
      var item=byId[row[1]],decision=YTCore.longVideoDecision(item);
      if(!decision.include){j.filtered++;if(item&&YTCore.durationSeconds((item.contentDetails||{}).duration)!==null)j.excludedShort++;else j.unverifiedDuration++;}
      return decision.include;
    });
  }
  ytWrite_(sheet,rows.length+2,added);SpreadsheetApp.flush();
  // A crash after the write replays the page. Existing IDs make replay idempotent.
  j.pages++;j.pageToken=response.nextPageToken||'';
  if(!j.pageToken){if(index){j.state='COMPLETE';j.message='Playlist index complete.';}else j.phase='METADATA';}
}
function ytFormatRow_(input,id,item,channels,j) {
  var row=ytEmptyRow_(input,id),now=new Date().toISOString();row[18]=now;
  if(!item){row[2]='UNAVAILABLE';row[19]='Not returned by YouTube. May be private, deleted or inaccessible.';return row;}
  if(j.excludeShorts){
    var decision=YTCore.longVideoDecision(item);
    if(!decision.include){row[2]='FILTERED';row[19]=decision.reason;return row;}
  }
  var s=item.snippet||{},stats=item.statistics||{},details=item.contentDetails||{},live=item.liveStreamingDetails||{};
  row[2]='OK';row[3]=s.title||'';row[4]=s.channelTitle||'';row[5]=s.channelId||'';row[6]=s.publishedAt||'';
  if(s.publishedAt&&isFinite(Date.parse(s.publishedAt)))row[7]=Utilities.formatDate(new Date(s.publishedAt),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');
  var seconds=YTCore.durationSeconds(details.duration);row[8]=seconds===null?'':seconds;
  if(seconds!==null){var t=Math.floor(seconds);row[9]=Math.floor(t/3600)+':'+String(Math.floor(t/60)%60).padStart(2,'0')+':'+String(t%60).padStart(2,'0');}
  row[10]=YTCore.count(stats.viewCount);row[11]=YTCore.count(stats.likeCount);row[12]=YTCore.count(stats.commentCount);
  row[13]=s.liveBroadcastContent||'none';row[14]=live.scheduledStartTime||'';row[15]=live.actualStartTime||'';
  row[16]=channels[s.channelId]||'';row[17]=s.description||'';
  if(!YTCore.inRange(s.publishedAt,j.dateRange)){row[2]='FILTERED';row[19]='Outside the selected UTC publication date range.';}
  return row;
}
function ytMetadata_(j,sheet) {
  var rows=ytRows_(sheet),pending=[];
  rows.forEach(function(r,i){if(r[2]==='PENDING'&&pending.length<50)pending.push(i);});
  if(!pending.length){j.state=rows.some(function(r){return ['INVALID','UNAVAILABLE'].indexOf(r[2])>=0;})?'COMPLETE_WITH_ISSUES':'COMPLETE';j.message='Collection finished. Review row status and empty fields.';return;}
  var ids=YTCore.unique(pending.map(function(i){return rows[i][1];})),items=Object.create(null),channels=Object.create(null);
  var response=ytApi_(j,'Videos','snippet,contentDetails,statistics,liveStreamingDetails',{id:ids.join(',')});
  (response.items||[]).forEach(function(item){items[item.id]=item;});
  var channelIds=YTCore.unique((response.items||[]).map(function(i){return (i.snippet||{}).channelId;}).filter(Boolean));
  if(channelIds.length){var cr=ytApi_(j,'Channels','snippet',{id:channelIds.join(',')});(cr.items||[]).forEach(function(c){channels[c.id]=(c.snippet||{}).country||'';});}
  // Group contiguous output rows into bulk writes. Successful rows are never rewritten on resume.
  var group=[],groupStart=0,last=-2;
  function flush(){if(group.length)ytWrite_(sheet,groupStart+2,group);group=[];}
  pending.forEach(function(i){if(i!==last+1){flush();groupStart=i;}group.push(ytFormatRow_(rows[i][0],rows[i][1],items[rows[i][1]],channels,j));last=i;});flush();SpreadsheetApp.flush();
}
function toolkitRetry(id,allRows) {
  return ytLocked_(function(){
    var j=ytRequireJob_(id);if(YT_TERMINAL.indexOf(j.state)<0)throw new Error('Finish or cancel the current job before refreshing.');
    if(j.mode==='playlists')throw new Error('Start a new playlist index to refresh its metadata.');
    var sheet=ytSheet_(j),rows=ytRows_(sheet),changed=0;
    rows.forEach(function(r){if(r[1]&&(allRows||r[2]==='UNAVAILABLE')){var fresh=ytEmptyRow_(r[0],r[1]);for(var k=0;k<r.length;k++)r[k]=fresh[k];changed++;}});
    if(!changed)throw new Error('No eligible rows to retry.');
    ytWrite_(sheet,2,rows);j.phase='METADATA';j.state='READY';j.message='Ready to refresh '+changed+' rows.';ytSave_(j);return ytState_(j);
  });
}
