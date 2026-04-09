// ============================================================
// GLOBAL CONSTANTS
// ============================================================


var METADATA_HEADERS = ["Title", "Duration", "Date(GMT)", "Day(GMT)", "Time(GMT)", "Date(Local)", "Day(Local)", "Time(Local)", "Country", "Views", "Likes", "Comments", "Description"];

var PLAYLIST_HEADERS = ["Playlist Name", "Video Name", "URL", "Published Date (UTC)", "Status"];
//                                                                                      ^^^^^^^^ new column we added

var PLAYLIST_INDEX_HEADERS = ["NAME", "PLAYLIST URL"];
// ============================================================
// API DOCUMENTATION MAP
// Each entry: [Header Name, API Object, API Property Path, API Method, Notes]
// ============================================================

var API_DOC_METADATA = [
  ["Title",       "Video",   "snippet.title",                          "YouTube.Videos.list('snippet')",                  "Video title as set by the uploader"],
  ["Duration",    "Video",   "contentDetails.duration",                "YouTube.Videos.list('contentDetails')",            "ISO 8601 duration, converted to H:MM:SS. Shows 'Upcoming' or 'LIVE NOW' for live broadcasts"],
  ["Date(GMT)",   "Video",   "snippet.publishedAt",                    "YouTube.Videos.list('snippet')",                  "Publish date formatted in GMT/UTC timezone"],
  ["Day(GMT)",    "Video",   "snippet.publishedAt",                    "YouTube.Videos.list('snippet')",                  "Day of week (e.g. Monday) derived from publishedAt in GMT"],
  ["Time(GMT)",   "Video",   "snippet.publishedAt",                    "YouTube.Videos.list('snippet')",                  "Time of day (HH:mm:ss) derived from publishedAt in GMT"],
  ["Date(Local)", "Video",   "snippet.publishedAt",                    "YouTube.Videos.list('snippet')",                  "Publish date formatted in script's local timezone (Project Settings > Time zone)"],
  ["Day(Local)",  "Video",   "snippet.publishedAt",                    "YouTube.Videos.list('snippet')",                  "Day of week derived from publishedAt in local timezone"],
  ["Time(Local)", "Video",   "snippet.publishedAt",                    "YouTube.Videos.list('snippet')",                  "Time of day derived from publishedAt in local timezone"],
  ["Country",     "Channel", "snippet.country",                        "YouTube.Channels.list('snippet')",                "ISO 3166-1 alpha-2 country code of the channel, mapped to full name"],
  ["Views",       "Video",   "statistics.viewCount",                   "YouTube.Videos.list('statistics')",               "Lifetime total views at time of fetch. Cast to Number for formulas. Shows 'Scheduled'/'Live' for broadcasts"],
  ["Likes",       "Video",   "statistics.likeCount",                   "YouTube.Videos.list('statistics')",               "Total like count. Cast to Number. May be hidden by uploader"],
  ["Comments",    "Video",   "statistics.commentCount",                "YouTube.Videos.list('statistics')",               "Total comment count. Cast to Number. May be disabled by uploader"],
  ["Description", "Video",   "snippet.description",                    "YouTube.Videos.list('snippet')",                  "First 2 lines, truncated to 150 chars"]
];

var API_DOC_PLAYLIST = [
  ["Playlist Name",        "Playlist/Channel", "snippet.title (channel) or snippet.title (playlist)", "YouTube.Channels.list / YouTube.Playlists.list",  "Source name: channel title or playlist title"],
  ["Video Name",           "PlaylistItem",     "snippet.title",                                       "YouTube.PlaylistItems.list('snippet')",            "Video title from the playlist item entry"],
  ["URL",                  "PlaylistItem",     "snippet.resourceId.videoId",                          "YouTube.PlaylistItems.list('snippet')",            "Constructed: https://youtube.com/watch?v={videoId}"],
  ["Published Date (UTC)", "PlaylistItem",     "snippet.publishedAt",                                 "YouTube.PlaylistItems.list('snippet')",            "Date the video was published, formatted in GMT"],
  ["Status",               "PlaylistItem",     "status.privacyStatus + snippet.title",                "YouTube.PlaylistItems.list('snippet,status')",     "Public | Private | Deleted — derived from privacyStatus and title sentinel"]
];

var API_DOC_PLAYLIST_INDEX = [
  ["NAME",         "Playlist", "snippet.title", "YouTube.Playlists.list('snippet')",  "Playlist title"],
  ["PLAYLIST URL", "Playlist", "id",            "YouTube.Playlists.list('snippet')",  "Constructed: https://youtube.com/playlist?list={id}"]
];

var API_DOC_LIVE_OVERRIDES = [
  ["Duration (Upcoming)",  "Video", "snippet.liveBroadcastContent = 'upcoming'",  "YouTube.Videos.list('snippet')",                  "Shows literal string 'Upcoming' instead of duration"],
  ["Duration (Live)",      "Video", "snippet.liveBroadcastContent = 'live'",      "YouTube.Videos.list('snippet')",                  "Shows literal string 'LIVE NOW' instead of duration"],
  ["Views (Upcoming)",     "Video", "snippet.liveBroadcastContent = 'upcoming'",  "YouTube.Videos.list('snippet')",                  "Shows literal string 'Scheduled' instead of view count"],
  ["Views (Live)",         "Video", "snippet.liveBroadcastContent = 'live'",      "YouTube.Videos.list('snippet')",                  "Shows literal string 'Live' instead of view count"],
  ["Date (Upcoming)",      "Video", "liveStreamingDetails.scheduledStartTime",    "YouTube.Videos.list('liveStreamingDetails')",      "Uses scheduled start time instead of publishedAt"],
  ["Date (Live)",          "Video", "liveStreamingDetails.actualStartTime",       "YouTube.Videos.list('liveStreamingDetails')",      "Uses actual start time instead of publishedAt"]
];

// ============================================================
// SHARED HELPERS
// ============================================================

/**
 * Writes bold headers one row above the data start row.
 * If startRow is 1 (no room above), inserts a new row and shifts data down.
 * Returns the (possibly adjusted) startRow for data.
 *
 * @param {Sheet}    sheet      - The active sheet
 * @param {number}   startRow   - Where data starts
 * @param {number}   startCol   - Column where headers begin
 * @param {string[]} headers    - 1D array of header strings
 * @param {boolean}  skipIfExists - If true, won't overwrite existing headers
 * @return {number}  adjusted startRow (shifted by 1 if a row was inserted)
 */
function writeHeaders(sheet, startRow, startCol, headers, skipIfExists) {
  var headerRow = startRow - 1;
  var numCols = headers.length;

  if (headerRow > 0) {
    if (skipIfExists) {
      var existing = sheet.getRange(headerRow, startCol).getValue().toString().trim();
      if (existing) return startRow; // headers already there
    }
    sheet.getRange(headerRow, startCol, 1, numCols)
         .setValues([headers])
         .setFontWeight("bold");
    return startRow;
  }

  // startRow === 1: no room above — insert a row for headers
  sheet.insertRowBefore(1);
  sheet.getRange(1, startCol, 1, numCols)
       .setValues([headers])
       .setFontWeight("bold");
  return startRow + 1; // data shifted down by 1
}

/**
 * Builds a metadata row array from a single YouTube video API item.
 * Shared by fetchVideoMetadata (single) and fetchMetadataBatch (bulk).
 *
 * @param {Object} item          - A single item from YouTube.Videos.list response
 * @param {Object} channelCache  - Shared cache: { channelId: channelObj|null }
 * @return {string[]} Row array matching METADATA_HEADERS
 */
function buildMetadataRow(item, channelCache) {
  var snippet = item.snippet;
  var stats = item.statistics;
  var liveDetails = item.liveStreamingDetails;

  // --- CHANNEL LOOKUP (cached) ---
  var chId = snippet.channelId;
  var channel = null;

  if (channelCache.hasOwnProperty(chId)) {
    channel = channelCache[chId];
  } else {
    try {
      var channelRes = YouTube.Channels.list('snippet', { id: chId });
      if (channelRes && channelRes.items && channelRes.items.length > 0) {
        channel = channelRes.items[0];
      }
    } catch (chErr) {
      channel = null;
    }
    channelCache[chId] = channel;
  }

  // --- LIVE / UPCOMING / STANDARD ---
  var liveStatus = snippet.liveBroadcastContent;
  var finalDuration, finalViews, relevantDateObj;

  if (liveStatus === 'upcoming') {
    finalDuration = "Upcoming";
    finalViews = "Scheduled";
    var dateStr = (liveDetails && liveDetails.scheduledStartTime) ? liveDetails.scheduledStartTime : snippet.publishedAt;
    relevantDateObj = new Date(dateStr);
  } else if (liveStatus === 'live') {
    finalDuration = "LIVE NOW";
    finalViews = "Live";
    var dateStr = (liveDetails && liveDetails.actualStartTime) ? liveDetails.actualStartTime : snippet.publishedAt;
    relevantDateObj = new Date(dateStr);
  } else {
    finalDuration = convertISO8601ToTime(item.contentDetails.duration);
    finalViews = Number(stats.viewCount) || 0;
    relevantDateObj = new Date(snippet.publishedAt);
  }

  // --- DATE & TIME ---
  var gmtDate = Utilities.formatDate(relevantDateObj, "GMT", "yyyy-MM-dd");
  var gmtDay  = Utilities.formatDate(relevantDateObj, "GMT", "EEEE");
  var gmtTime = Utilities.formatDate(relevantDateObj, "GMT", "HH:mm:ss");

  var localTz   = Session.getScriptTimeZone();
  var localDate  = Utilities.formatDate(relevantDateObj, localTz, "yyyy-MM-dd");
  var localDay   = Utilities.formatDate(relevantDateObj, localTz, "EEEE");
  var localTime  = Utilities.formatDate(relevantDateObj, localTz, "HH:mm:ss");

  // --- COUNTRY ---
  var countryCode = (channel && channel.snippet.country) ? channel.snippet.country : "N/A";
  var countryName = getFullCountryName(countryCode);

  // --- DESCRIPTION ---
  var cleanDescription = snippet.description.split('\n').slice(0, 2).join(' ').substring(0, 150) + "...";

  return [
    snippet.title, finalDuration,
    gmtDate, gmtDay, gmtTime,
    localDate, localDay, localTime,
    countryName, finalViews,
    Number(stats.likeCount) || 0,
    Number(stats.commentCount) || 0,
    cleanDescription
  ];
}

/**
 * Fetches metadata for a SINGLE video ID. Used by fetchYoutubeDetails and retryFailedRows.
 */
function fetchVideoMetadata(videoId, channelCache) {
  try {
    var videoRes = YouTube.Videos.list('snippet,contentDetails,statistics,liveStreamingDetails', { id: videoId });

    if (!videoRes.items || videoRes.items.length === 0) {
      return {
        success: false,
        data: ["Video Not Found / Geoblocked"].concat(new Array(METADATA_HEADERS.length - 1).fill("")),
        message: "Not found"
      };
    }

    return { success: true, data: buildMetadataRow(videoRes.items[0], channelCache) };

  } catch (e) {
    return {
      success: false,
      data: ["Error", e.message].concat(new Array(METADATA_HEADERS.length - 2).fill("")),
      message: e.message
    };
  }
}

/**
 * Fetches metadata for MANY video IDs in batches of 50 (max allowed by YouTube API).
 * Returns an array of row arrays in the same order as the input videoIds.
 *
 * @param {string[]} videoIds     - Array of YouTube video IDs
 * @param {Object}   channelCache - Shared cache
 * @return {string[][]} Array of row arrays matching METADATA_HEADERS
 */
function fetchMetadataBatch(videoIds, channelCache) {
  var resultMap = {};
  var notFoundRow    = ["Video Not Found / Geoblocked"].concat(new Array(METADATA_HEADERS.length - 1).fill(""));
  var unavailableRow = ["[Deleted / Private — no metadata]"].concat(new Array(METADATA_HEADERS.length - 1).fill(""));

  for (var i = 0; i < videoIds.length; i += 50) {
    var batch = videoIds.slice(i, Math.min(i + 50, videoIds.length));
    // Filter out empty/null IDs from the batch request
    var validBatch = [];
    for (var b = 0; b < batch.length; b++) {
      if (batch[b]) validBatch.push(batch[b]);
    }
    if (validBatch.length === 0) continue;

    try {
      var videoRes = YouTube.Videos.list('snippet,contentDetails,statistics,liveStreamingDetails', { id: validBatch.join(',') });
      if (videoRes.items) {
        for (var j = 0; j < videoRes.items.length; j++) {
          resultMap[videoRes.items[j].id] = buildMetadataRow(videoRes.items[j], channelCache);
        }
      }
    } catch (e) {
      var errorRow = ["Error", e.message].concat(new Array(METADATA_HEADERS.length - 2).fill(""));
      for (var k = 0; k < validBatch.length; k++) {
        if (!resultMap[validBatch[k]]) {
          resultMap[validBatch[k]] = errorRow;
        }
      }
    }
  }

  // Return in original order; null IDs (deleted/private) get a placeholder row
  var output = [];
  for (var v = 0; v < videoIds.length; v++) {
    if (!videoIds[v]) {
      output.push(unavailableRow);
    } else {
      output.push(resultMap[videoIds[v]] || notFoundRow);
    }
  }
  return output;
}

// ============================================================
// MENU
// ============================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 YouTube Data Toolkit')
      .addItem('1. Get Metadata from Video URLs', 'fetchYoutubeDetails')
      .addItem('1b. Retry Failed / Error Rows', 'retryFailedRows')
      .addSeparator()
      .addItem('2. List All Videos in a Playlist', 'importFromPlaylist')
      .addItem('3. List All Videos from a Channel', 'importFromChannel')
      .addSeparator()
      .addItem('4. Index All Playlists from a Channel', 'getAllPlaylistsFromChannel')
      .addSeparator()
      .addItem('📖 Print API Documentation', 'printApiDocumentation')
      .addToUi();
}

// ============================================================
// PRINT API DOCUMENTATION
// ============================================================

function printApiDocumentation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "API Documentation";

  // Delete existing doc sheet if it exists, then create fresh
  var existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);
  var doc = ss.insertSheet(sheetName);

  var row = 1;
  var docHeaders = ["Column Header", "API Object", "API Property Path", "API Method", "Notes"];
  var numCols = docHeaders.length;

  // --- HELPER: Write a section ---
  function writeSection(title, data) {
    // Section title
    doc.getRange(row, 1, 1, numCols).merge();
    doc.getRange(row, 1)
       .setValue(title)
       .setFontWeight("bold")
       .setFontSize(12)
       .setBackground("#4285f4")
       .setFontColor("#ffffff");
    row++;

    // Column headers
    doc.getRange(row, 1, 1, numCols)
       .setValues([docHeaders])
       .setFontWeight("bold")
       .setBackground("#e8f0fe");
    row++;

    // Data rows
    if (data.length > 0) {
      doc.getRange(row, 1, data.length, numCols).setValues(data);
      row += data.length;
    }

    // Blank spacer row
    row++;
  }

  // --- TITLE ---
  doc.getRange(1, 1, 1, numCols).merge();
  doc.getRange(1, 1)
     .setValue("YouTube Data Toolkit — API Reference")
     .setFontWeight("bold")
     .setFontSize(14)
     .setFontColor("#333333");
  row = 2;

  doc.getRange(row, 1, 1, numCols).merge();
  doc.getRange(row, 1)
     .setValue("Generated: " + Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd HH:mm:ss") + " UTC  |  Script Timezone: " + Session.getScriptTimeZone())
     .setFontColor("#888888")
     .setFontSize(10);
  row = 4;

  // --- SECTIONS ---
  writeSection("1. Video Metadata Headers  (Menu: Get Metadata / Retry Failed / Auto-fetch after Channel/Playlist import)", API_DOC_METADATA);
  writeSection("2. Live / Upcoming Overrides  (When video is a livestream or scheduled premiere)", API_DOC_LIVE_OVERRIDES);
  writeSection("3. Playlist / Channel Import Headers  (Menu: List Videos in Playlist / Channel)", API_DOC_PLAYLIST);
  writeSection("4. Playlist Index Headers  (Menu: Index All Playlists from Channel)", API_DOC_PLAYLIST_INDEX);

  // --- API REFERENCE LINKS ---
  doc.getRange(row, 1, 1, numCols).merge();
  doc.getRange(row, 1)
     .setValue("API Reference Links")
     .setFontWeight("bold")
     .setFontSize(12)
     .setBackground("#4285f4")
     .setFontColor("#ffffff");
  row++;

  var links = [
    ["YouTube Data API v3 — Videos",        "https://developers.google.com/youtube/v3/docs/videos"],
    ["YouTube Data API v3 — Channels",       "https://developers.google.com/youtube/v3/docs/channels"],
    ["YouTube Data API v3 — PlaylistItems",  "https://developers.google.com/youtube/v3/docs/playlistItems"],
    ["YouTube Data API v3 — Playlists",      "https://developers.google.com/youtube/v3/docs/playlists"],
    ["YouTube Data API v3 — Search",         "https://developers.google.com/youtube/v3/docs/search"],
    ["Apps Script — YouTube Service",        "https://developers.google.com/apps-script/advanced/youtube"]
  ];

  for (var l = 0; l < links.length; l++) {
    doc.getRange(row, 1).setValue(links[l][0]).setFontWeight("bold");
    doc.getRange(row, 2, 1, numCols - 1).merge();
    doc.getRange(row, 2).setValue(links[l][1]).setFontColor("#1a73e8");
    row++;
  }

  row += 2;

  // --- QUOTA INFO ---
  doc.getRange(row, 1, 1, numCols).merge();
  doc.getRange(row, 1)
     .setValue("API Quota Usage Per Function")
     .setFontWeight("bold")
     .setFontSize(12)
     .setBackground("#4285f4")
     .setFontColor("#ffffff");
  row++;

  var quotaHeaders = ["Function", "API Calls", "Quota Cost Per Call", "Example: 200 Videos", "Notes"];
  doc.getRange(row, 1, 1, numCols)
     .setValues([quotaHeaders])
     .setFontWeight("bold")
     .setBackground("#e8f0fe");
  row++;

  var quotaData = [
    ["1. Get Metadata from URLs",           "1 Videos.list + 1 Channels.list per video",         "1 unit each",   "400 calls (200+200)",    "Channel calls cached per unique channel"],
    ["1b. Retry Failed Rows",               "Same as above, only for failed rows",                "1 unit each",   "Varies",                 "Only retries Error / Not Found rows"],
    ["2. List Videos in Playlist",           "1 PlaylistItems.list per 50 videos + batch metadata","1 + 1 units",  "4+4 = 8 calls",          "Batched: 50 videos per Videos.list call"],
    ["3. List Videos from Channel",          "Same as Playlist + date filter",                     "1 + 1 units",  "4+4 = 8 calls",          "Filters client-side, same API usage"],
    ["4. Index Playlists from Channel",      "1 Playlists.list per 50 playlists",                  "1 unit each",  "1 call for <50",         "Lightweight, no video data fetched"]
  ];
  doc.getRange(row, 1, quotaData.length, numCols).setValues(quotaData);
  row += quotaData.length + 1;

  // --- DATA TYPE NOTES ---
  doc.getRange(row, 1, 1, numCols).merge();
  doc.getRange(row, 1)
     .setValue("Data Type Notes")
     .setFontWeight("bold")
     .setFontSize(12)
     .setBackground("#4285f4")
     .setFontColor("#ffffff");
  row++;

  var typeNotes = [
    ["Views / Likes / Comments", "YouTube API returns these as STRINGS. This toolkit converts them to Numbers via Number() so SUM/SORT/formulas work correctly."],
    ["Dates (GMT vs Local)",     "GMT columns use hardcoded 'GMT' timezone. Local columns use Session.getScriptTimeZone() from Project Settings > Time zone."],
    ["Country",                  "API returns ISO 3166-1 alpha-2 code (e.g. 'US'). Mapped to full name (e.g. 'United States') via getFullCountryName()."],
    ["Duration",                 "API returns ISO 8601 (e.g. 'PT1H23M45S'). Converted to H:MM:SS string via convertISO8601ToTime()."],
    ["Description",              "Truncated to first 2 lines, max 150 characters, appended with '...'"]
  ];

  for (var t = 0; t < typeNotes.length; t++) {
    doc.getRange(row, 1).setValue(typeNotes[t][0]).setFontWeight("bold");
    doc.getRange(row, 2, 1, numCols - 1).merge();
    doc.getRange(row, 2).setValue(typeNotes[t][1]);
    row++;
  }

  // --- FORMATTING ---
  doc.setColumnWidth(1, 200);
  doc.setColumnWidth(2, 180);
  doc.setColumnWidth(3, 320);
  doc.setColumnWidth(4, 340);
  doc.setColumnWidth(5, 400);
  doc.setFrozenRows(0);

  // Activate the new sheet
  ss.setActiveSheet(doc);
  ss.toast('API Documentation sheet created!', 'Success');
}

// ============================================================
// 1. GET METADATA FROM VIDEO URLs
// ============================================================

function fetchYoutubeDetails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();
  var startRow = range.getRow();
  var startCol = range.getColumn();
  var values = range.getValues();
  var output = [];

  // Headers one row above, in the column after URLs
  writeHeaders(sheet, startRow, startCol + 1, METADATA_HEADERS, false);

  var channelCache = {};

  for (var i = 0; i < values.length; i++) {
    var url = values[i][0];
    var videoId = extractVideoId(url);

    // Check second column if first didn't have a video ID
    if (!videoId && values[i].length > 1) {
      url = values[i][1];
      videoId = extractVideoId(url);
    }

    if (!videoId) {
      output.push(new Array(METADATA_HEADERS.length).fill(""));
      continue;
    }

    var result = fetchVideoMetadata(videoId, channelCache);
    output.push(result.data);
  }

  if (output.length > 0) {
    sheet.getRange(startRow, startCol + 1, output.length, METADATA_HEADERS.length).setValues(output);
  }
}

// ============================================================
// 1b. RETRY FAILED / ERROR ROWS
// ============================================================

function retryFailedRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();
  var startRow = range.getRow();
  var startCol = range.getColumn();
  var numRows = range.getNumRows();
  var values = range.getValues();

  var numDataCols = METADATA_HEADERS.length;
  var titleCol = startCol + 1;

  // Write headers (skip if they already exist)
  var adjustedStartRow = writeHeaders(sheet, startRow, startCol + 1, METADATA_HEADERS, true);

  // If a row was inserted, re-read values from the shifted position
  if (adjustedStartRow !== startRow) {
    startRow = adjustedStartRow;
    range = sheet.getRange(startRow, startCol, numRows, range.getNumColumns());
    values = range.getValues();
  }

  var retried = 0;
  var fixed = 0;
  var stillFailed = 0;
  var channelCache = {};

  sheet.getParent().toast('Scanning for failed rows...', 'Retry Tool');

  for (var i = 0; i < numRows; i++) {
    var row = startRow + i;
    var url = values[i][0];
    var videoId = extractVideoId(url);

    if (!videoId && values[i].length > 1) {
      url = values[i][1];
      videoId = extractVideoId(url);
    }

    if (!videoId) continue;

    // Check if this row needs retrying
    var existingTitle = sheet.getRange(row, titleCol).getValue().toString().trim();
    var needsRetry = false;

    if (!existingTitle || existingTitle === "") {
      needsRetry = true;
    } else if (existingTitle.indexOf("Error") === 0) {
      needsRetry = true;
    } else if (existingTitle.indexOf("Video Not Found") === 0) {
      needsRetry = true;
    }

    if (!needsRetry) continue;

    retried++;
    var result = fetchVideoMetadata(videoId, channelCache);

    if (result.success) {
      sheet.getRange(row, startCol + 1, 1, numDataCols).setValues([result.data]);
      fixed++;
    } else {
      // Update with latest error info (mark as retry attempt)
      var retryRow = ["Error (Retry)", result.message || "Unknown"].concat(new Array(numDataCols - 2).fill(""));
      sheet.getRange(row, startCol + 1, 1, numDataCols).setValues([retryRow]);
      stillFailed++;
    }
  }

  var msg = '';
  if (retried === 0) {
    msg = 'No failed rows found in selection. All rows have data.';
  } else {
    msg = 'Retried ' + retried + ' rows: ' + fixed + ' fixed, ' + stillFailed + ' still failing.';
  }
  SpreadsheetApp.getUi().alert('Retry Complete', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================
// 2. LIST ALL VIDEOS IN A PLAYLIST
// ============================================================

function importFromPlaylist() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.prompt(
    'Import Playlist (No Shorts)',
    'Paste the YouTube Playlist URL (or ID):',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() == ui.Button.CANCEL) return;

  var input = result.getResponseText();
  var playlistId = extractPlaylistId(input);

  if (!playlistId) {
    ui.alert('Error: Could not find a valid Playlist ID in that URL.');
    return;
  }

  var playlistName = "Playlist " + playlistId;
  try {
    var response = YouTube.Playlists.list('snippet', { id: playlistId });
    if (response.items && response.items.length > 0) {
      playlistName = response.items[0].snippet.title;
    }
  } catch (e) {
    console.log("Could not fetch name: " + e.message);
  }

  processPlaylistToSheet(playlistId, playlistName, null, null);
}

// ============================================================
// 3. LIST ALL VIDEOS FROM A CHANNEL (with Date Picker Dialog)
// ============================================================

function importFromChannel() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.prompt(
    'Import Channel (No Shorts)',
    'Paste the YouTube Channel URL, Handle (@name), or ID:',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() == ui.Button.CANCEL) return;

  var input = result.getResponseText();
  var channelData = getChannelDetails(input);

  if (!channelData) {
    ui.alert('Error: Could not find channel. Try using the Handle (@name).');
    return;
  }

  // Store channel info + active cell position so the async callback can access them
  var activeCell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();
  var props = PropertiesService.getScriptProperties();
  props.setProperty('pendingUploadsId', channelData.uploadsId);
  props.setProperty('pendingChannelTitle', channelData.title);
  props.setProperty('pendingStartRow', activeCell.getRow().toString());
  props.setProperty('pendingStartCol', activeCell.getColumn().toString());
  props.setProperty('pendingSheetName', SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName());

  // Show HTML date picker dialog
  var html = HtmlService.createHtmlOutput(getDatePickerHtml())
      .setWidth(380)
      .setHeight(320);
  ui.showModalDialog(html, '📅 Date Range Filter — ' + channelData.title);
}

/**
 * Returns the HTML string for the date picker dialog.
 */
function getDatePickerHtml() {
  return '\
<!DOCTYPE html>\
<html>\
<head>\
<style>\
  body { font-family: Google Sans, Roboto, Arial, sans-serif; padding: 16px; color: #333; }\
  h3 { margin: 0 0 12px 0; font-size: 15px; color: #555; }\
  label { display: block; font-weight: 500; margin: 10px 0 4px 0; font-size: 13px; }\
  input[type="date"] { width: 100%; padding: 8px 10px; font-size: 14px; border: 1px solid #ccc; border-radius: 6px; box-sizing: border-box; }\
  input[type="date"]:focus { outline: none; border-color: #4285f4; box-shadow: 0 0 0 2px rgba(66,133,244,0.2); }\
  .hint { font-size: 11px; color: #888; margin-top: 2px; }\
  .btn-row { margin-top: 20px; display: flex; gap: 10px; justify-content: flex-end; }\
  button { padding: 8px 20px; font-size: 13px; border-radius: 6px; cursor: pointer; border: none; }\
  .btn-primary { background: #4285f4; color: #fff; } .btn-primary:hover { background: #3367d6; }\
  .btn-secondary { background: #f1f3f4; color: #333; } .btn-secondary:hover { background: #e0e0e0; }\
  .btn-skip { background: #fff; color: #4285f4; border: 1px solid #4285f4; } .btn-skip:hover { background: #e8f0fe; }\
  #status { margin-top: 12px; font-size: 12px; color: #d93025; display: none; }\
</style>\
</head>\
<body>\
  <h3>Filter videos by publish date (UTC)</h3>\
\
  <label for="startDate">Start Date</label>\
  <input type="date" id="startDate" />\
  <div class="hint">Earliest video date to include</div>\
\
  <label for="endDate">End Date</label>\
  <input type="date" id="endDate" />\
  <div class="hint">Leave blank to use today</div>\
\
  <div id="status"></div>\
\
  <div class="btn-row">\
    <button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>\
    <button class="btn-skip" onclick="submitDates(true)">Skip — Get All Videos</button>\
    <button class="btn-primary" onclick="submitDates(false)">Apply Filter</button>\
  </div>\
\
<script>\
  function submitDates(skipFilter) {\
    var statusEl = document.getElementById("status");\
    statusEl.style.display = "none";\
\
    var startVal = document.getElementById("startDate").value;\
    var endVal   = document.getElementById("endDate").value;\
\
    if (!skipFilter && !startVal) {\
      statusEl.textContent = "Please enter a Start Date, or click Skip to get all videos.";\
      statusEl.style.display = "block";\
      return;\
    }\
\
    if (!skipFilter && startVal && endVal && startVal > endVal) {\
      statusEl.textContent = "Start date cannot be after end date.";\
      statusEl.style.display = "block";\
      return;\
    }\
\
    var payload = {\
      startDate: skipFilter ? "" : startVal,\
      endDate:   skipFilter ? "" : endVal\
    };\
\
    document.querySelectorAll("button").forEach(function(b) { b.disabled = true; b.style.opacity = 0.5; });\
    statusEl.textContent = "Fetching videos...";\
    statusEl.style.color = "#4285f4";\
    statusEl.style.display = "block";\
\
    google.script.run\
      .withSuccessHandler(function() { google.script.host.close(); })\
      .withFailureHandler(function(e) {\
        statusEl.textContent = "Error: " + e.message;\
        statusEl.style.color = "#d93025";\
        statusEl.style.display = "block";\
        document.querySelectorAll("button").forEach(function(b) { b.disabled = false; b.style.opacity = 1; });\
      })\
      .processChannelImportFromDialog(payload);\
  }\
</script>\
</body>\
</html>';
}

/**
 * Server-side callback from the date picker dialog.
 * Reads the pending channel info from script properties and runs the import.
 */
function processChannelImportFromDialog(payload) {
  var props = PropertiesService.getScriptProperties();
  var uploadsId    = props.getProperty('pendingUploadsId');
  var channelTitle = props.getProperty('pendingChannelTitle');
  var savedRow     = parseInt(props.getProperty('pendingStartRow'), 10);
  var savedCol     = parseInt(props.getProperty('pendingStartCol'), 10);
  var sheetName    = props.getProperty('pendingSheetName');

  // Clean up all stored properties
  props.deleteProperty('pendingUploadsId');
  props.deleteProperty('pendingChannelTitle');
  props.deleteProperty('pendingStartRow');
  props.deleteProperty('pendingStartCol');
  props.deleteProperty('pendingSheetName');

  if (!uploadsId || !channelTitle || !savedRow || !savedCol || !sheetName) {
    throw new Error('Channel info not found. Please try again from the menu.');
  }

  var startDate = null;
  var endDate = null;

  if (payload.startDate) {
    startDate = parseInputDate(payload.startDate);
    if (!startDate) throw new Error('Invalid start date: ' + payload.startDate);

    if (payload.endDate) {
      endDate = parseInputDate(payload.endDate);
      if (!endDate) throw new Error('Invalid end date: ' + payload.endDate);
    } else {
      endDate = new Date();
    }

    endDate.setUTCHours(23, 59, 59, 999);

    if (startDate > endDate) {
      throw new Error('Start date cannot be after end date.');
    }
  }

  // Pass the saved cell position so processPlaylistToSheet writes to the correct location
  processPlaylistToSheet(uploadsId, channelTitle, startDate, endDate, sheetName, savedRow, savedCol);
}

// ============================================================
// SHARED: Process Playlist/Channel Videos to Sheet
// ============================================================

function processPlaylistToSheet(playlistId, sourceName, startDate, endDate, sheetName, forcedRow, forcedCol) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  // Use forced position (from dialog callback) or fall back to active cell
  var startRow = forcedRow || sheet.getActiveCell().getRow();
  var startCol = forcedCol || sheet.getActiveCell().getColumn();
  var hasDateFilter = (startDate && endDate);
  var numCols = PLAYLIST_HEADERS.length;

  // Headers
  writeHeaders(sheet, startRow, startCol, PLAYLIST_HEADERS, false);

  var filterLabel = hasDateFilter
    ? ' (Filtering: ' + Utilities.formatDate(startDate, "GMT", "yyyy-MM-dd") + ' to ' + Utilities.formatDate(endDate, "GMT", "yyyy-MM-dd") + ')'
    : '';
  sheet.getParent().toast('Fetching names and URLs...' + filterLabel, 'YouTube Tools');

  var rows = [];
  var rowVideoIds = [];   // Parallel array: video ID for each in-range row
  var skippedRows = [];
  var nextPageToken = '';

  do {
    try {
      var response = YouTube.PlaylistItems.list('snippet,status', {
        playlistId: playlistId,
        maxResults: 50,
        pageToken: nextPageToken
      });

      var items = response.items;
      if (!items || items.length === 0) break;

      for (var i = 0; i < items.length; i++) {
        var privacyStatus = (items[i].status && items[i].status.privacyStatus) ? items[i].status.privacyStatus : '';
        var snippet = items[i].snippet;
        var videoTitle = snippet.title;
        var vidId = snippet.resourceId.videoId;
        var videoUrl = "https://www.youtube.com/watch?v=" + vidId;
        var publishedAt = new Date(snippet.publishedAt);
        var publishedDateStr = Utilities.formatDate(publishedAt, "GMT", "yyyy-MM-dd HH:mm:ss");

        // Detect deleted or private videos — YouTube API returns privacyStatus 'private'
        // for both, but uses sentinel titles "Deleted video" vs "Private video".
        var videoStatus = "Public";
        var isUnavailable = false;
        if (privacyStatus === 'private') {
          isUnavailable = true;
          videoStatus = (videoTitle === "Deleted video") ? "Deleted" : "Private";
        }

        var row = [sourceName, videoTitle, videoUrl, publishedDateStr, videoStatus];

        if (hasDateFilter) {
          if (publishedAt >= startDate && publishedAt <= endDate) {
            rows.push(row);
            rowVideoIds.push(isUnavailable ? null : vidId); // null skips metadata fetch
          } else {
            skippedRows.push(row);
          }
        } else {
          rows.push(row);
          rowVideoIds.push(isUnavailable ? null : vidId);
        }
      }

      nextPageToken = response.nextPageToken;
    } catch (e) {
      ui.alert("Error: " + e.message);
      break;
    }
  } while (nextPageToken);

  if (rows.length > 0 || skippedRows.length > 0) {
    var allOutput = [];

    for (var r = 0; r < rows.length; r++) {
      allOutput.push(rows[r]);
    }

    if (hasDateFilter && skippedRows.length > 0) {
      allOutput.push(["--- SKIPPED (Outside " + Utilities.formatDate(startDate, "GMT", "yyyy-MM-dd") + " to " + Utilities.formatDate(endDate, "GMT", "yyyy-MM-dd") + ") ---", "", "", "", ""]);
      for (var s = 0; s < skippedRows.length; s++) {
        allOutput.push(skippedRows[s]);
      }
    }

    // Write playlist data
    sheet.getRange(startRow, startCol, allOutput.length, numCols).setValues(allOutput);

    var msg = 'Imported ' + rows.length + ' videos!';
    if (hasDateFilter) {
      msg += ' (' + skippedRows.length + ' skipped — outside date range)';
    }
    sheet.getParent().toast(msg, 'Success');

    // ====================================================
    // AUTO-FETCH METADATA (batch: 50 videos per API call)
    // ====================================================
    if (rowVideoIds.length > 0) {
      var metaStartCol = startCol + numCols; // Column right after playlist data
      var totalBatches = Math.ceil(rowVideoIds.length / 50);

      // Write metadata headers
      writeHeaders(sheet, startRow, metaStartCol, METADATA_HEADERS, false);

      sheet.getParent().toast(
        'Fetching metadata for ' + rowVideoIds.length + ' videos (' + totalBatches + ' batch' + (totalBatches > 1 ? 'es' : '') + ')...',
        'YouTube Tools'
      );

      var channelCache = {};
      var metaRows = fetchMetadataBatch(rowVideoIds, channelCache);

      // Write metadata for in-range rows only
      if (metaRows.length > 0) {
        sheet.getRange(startRow, metaStartCol, metaRows.length, METADATA_HEADERS.length).setValues(metaRows);
      }

      sheet.getParent().toast(
        'Done! ' + rows.length + ' videos with full metadata.',
        'Success'
      );
    }

  } else {
    ui.alert('No videos found in this playlist.' + (hasDateFilter ? ' Try a wider date range.' : ''));
  }
}

// ============================================================
// 4. INDEX ALL PLAYLISTS FROM A CHANNEL
// ============================================================

function getAllPlaylistsFromChannel() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var result = ui.prompt(
    'Get All Playlists',
    'Paste the Channel URL, Handle (@name), or ID:',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() == ui.Button.CANCEL) return;

  var input = result.getResponseText();
  var channelInfo = getChannelDetails(input);

  if (!channelInfo) {
    ui.alert('Error: Could not find channel. Try using the Handle (@name).');
    return;
  }

  var channelId = channelInfo.channelId;
  var activeCell = sheet.getActiveCell();
  var startRow = activeCell.getRow();
  var startCol = activeCell.getColumn();

  // Headers
  writeHeaders(sheet, startRow, startCol, PLAYLIST_INDEX_HEADERS, false);

  sheet.getParent().toast('Fetching Playlists for: ' + channelInfo.title, 'YouTube Tools');

  var rows = [];
  var nextPageToken = '';

  do {
    try {
      var response = YouTube.Playlists.list('snippet', {
        channelId: channelId,
        maxResults: 50,
        pageToken: nextPageToken
      });

      var items = response.items;
      if (!items || items.length === 0) break;

      for (var i = 0; i < items.length; i++) {
        var title = items[i].snippet.title;
        var pId = items[i].id;
        var url = "https://www.youtube.com/playlist?list=" + pId;
        rows.push([title, url]);
      }

      nextPageToken = response.nextPageToken;
    } catch (e) {
      ui.alert("Error fetching playlists: " + e.message);
      return;
    }
  } while (nextPageToken);

  if (rows.length > 0) {
    sheet.getRange(startRow, startCol, rows.length, 2).setValues(rows);
    sheet.getParent().toast('Found ' + rows.length + ' playlists!', 'Success');
  } else {
    ui.alert('No public playlists found for this channel.');
  }
}

// ============================================================
// HELPERS: YouTube Parsing
// ============================================================

/**
 * Extracts Channel Details (ID and Uploads Playlist)
 */
function getChannelDetails(input) {
  var request = {};

  if (input.includes("@")) {
    var handleMatch = input.match(/(@[\w\.-]+)/);
    if (handleMatch) request = { forHandle: handleMatch[1] };
  } else if (input.includes("channel/")) {
    var idMatch = input.match(/channel\/([\w-]+)/);
    if (idMatch) request = { id: idMatch[1] };
  } else if (input.startsWith("UC")) {
    request = { id: input };
  } else if (input.startsWith("@")) {
    request = { forHandle: input };
  }

  if (Object.keys(request).length === 0) return null;

  try {
    var response = YouTube.Channels.list('snippet,contentDetails', request);
    if (response.items && response.items.length > 0) {
      return {
        title: response.items[0].snippet.title,
        channelId: response.items[0].id,
        uploadsId: response.items[0].contentDetails.relatedPlaylists.uploads
      };
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert("Error: " + e.message);
  }

  return null;
}

/** Extract Playlist ID from URL or raw string */
function extractPlaylistId(url) {
  var regex = /[?&]list=([^&]+)/;
  var match = url.match(regex);
  if (match) return match[1];
  if (!url.includes("http") && url.length > 10 && !url.startsWith("UC")) return url;
  return null;
}

/** Extract Video ID from URL */
function extractVideoId(url) {
  if (!url || typeof url !== 'string') return null;
  var regex = /(?:v=|\/)([0-9A-Za-z_-]{11}).*/;
  var match = url.match(regex);
  return match ? match[1] : null;
}

// ============================================================
// HELPERS: Duration & Date Parsing
// ============================================================

/** Parse ISO 8601 duration to total seconds */
function parseDuration(duration) {
  var hours = 0, minutes = 0, seconds = 0;

  var hoursMatch = duration.match(/(\d+)H/);
  var minsMatch  = duration.match(/(\d+)M/);
  var secsMatch  = duration.match(/(\d+)S/);

  if (hoursMatch) hours   = parseInt(hoursMatch[1]);
  if (minsMatch)  minutes = parseInt(minsMatch[1]);
  if (secsMatch)  seconds = parseInt(secsMatch[1]);

  return (hours * 3600) + (minutes * 60) + seconds;
}

/** Convert ISO 8601 duration to H:MM:SS string */
function convertISO8601ToTime(duration) {
  var hours = 0, minutes = 0, seconds = 0;

  if (duration.indexOf('H') > -1) hours   = parseInt(duration.match(/(\d+)H/)[1]);
  if (duration.indexOf('M') > -1) minutes = parseInt(duration.match(/(\d+)M/)[1]);
  if (duration.indexOf('S') > -1) seconds = parseInt(duration.match(/(\d+)S/)[1]);

  var paddedMinutes = minutes < 10 ? "0" + minutes : minutes;
  var paddedSeconds = seconds < 10 ? "0" + seconds : seconds;

  return hours + ":" + paddedMinutes + ":" + paddedSeconds;
}

/** Parse YYYY-MM-DD string into a UTC Date object. Returns null if invalid. */
function parseInputDate(str) {
  if (!str) return null;
  var parts = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!parts) return null;

  var year  = parseInt(parts[1], 10);
  var month = parseInt(parts[2], 10) - 1;
  var day   = parseInt(parts[3], 10);

  var d = new Date(Date.UTC(year, month, day));

  if (d.getUTCFullYear() !== year || d.getUTCMonth() !== month || d.getUTCDate() !== day) {
    return null;
  }
  return d;
}

// ============================================================
// HELPERS: Country Code Mapping
// ============================================================

function getFullCountryName(code) {
  var countries = {
    "AF": "Afghanistan", "AL": "Albania", "DZ": "Algeria", "AS": "American Samoa", "AD": "Andorra", "AO": "Angola", "AI": "Anguilla", "AQ": "Antarctica", "AG": "Antigua and Barbuda", "AR": "Argentina", "AM": "Armenia", "AW": "Aruba", "AU": "Australia", "AT": "Austria", "AZ": "Azerbaijan", "BS": "Bahamas", "BH": "Bahrain", "BD": "Bangladesh", "BB": "Barbados", "BY": "Belarus", "BE": "Belgium", "BZ": "Belize", "BJ": "Benin", "BM": "Bermuda", "BT": "Bhutan", "BO": "Bolivia", "BA": "Bosnia and Herzegovina", "BW": "Botswana", "BR": "Brazil", "BN": "Brunei", "BG": "Bulgaria", "BF": "Burkina Faso", "BI": "Burundi", "KH": "Cambodia", "CM": "Cameroon", "CA": "Canada", "CV": "Cape Verde", "KY": "Cayman Islands", "CF": "Central African Republic", "TD": "Chad", "CL": "Chile", "CN": "China", "CO": "Colombia", "KM": "Comoros", "CG": "Congo", "CD": "Congo (Dem. Rep.)", "CR": "Costa Rica", "CI": "Cote d'Ivoire", "HR": "Croatia", "CU": "Cuba", "CY": "Cyprus", "CZ": "Czech Republic", "DK": "Denmark", "DJ": "Djibouti", "DM": "Dominica", "DO": "Dominican Republic", "EC": "Ecuador", "EG": "Egypt", "SV": "El Salvador", "GQ": "Equatorial Guinea", "ER": "Eritrea", "EE": "Estonia", "ET": "Ethiopia", "FJ": "Fiji", "FI": "Finland", "FR": "France", "GA": "Gabon", "GM": "Gambia", "GE": "Georgia", "DE": "Germany", "GH": "Ghana", "GR": "Greece", "GL": "Greenland", "GD": "Grenada", "GU": "Guam", "GT": "Guatemala", "GN": "Guinea", "GY": "Guyana", "HT": "Haiti", "HN": "Honduras", "HK": "Hong Kong", "HU": "Hungary", "IS": "Iceland", "IN": "India", "ID": "Indonesia", "IR": "Iran", "IQ": "Iraq", "IE": "Ireland", "IL": "Israel", "IT": "Italy", "JM": "Jamaica", "JP": "Japan", "JO": "Jordan", "KZ": "Kazakhstan", "KE": "Kenya", "KP": "North Korea", "KR": "South Korea", "KW": "Kuwait", "KG": "Kyrgyzstan", "LA": "Laos", "LV": "Latvia", "LB": "Lebanon", "LS": "Lesotho", "LR": "Liberia", "LY": "Libya", "LI": "Liechtenstein", "LT": "Lithuania", "LU": "Luxembourg", "MO": "Macao", "MK": "Macedonia", "MG": "Madagascar", "MW": "Malawi", "MY": "Malaysia", "MV": "Maldives", "ML": "Mali", "MT": "Malta", "MH": "Marshall Islands", "MQ": "Martinique", "MR": "Mauritania", "MU": "Mauritius", "MX": "Mexico", "FM": "Micronesia", "MD": "Moldova", "MC": "Monaco", "MN": "Mongolia", "ME": "Montenegro", "MA": "Morocco", "MZ": "Mozambique", "MM": "Myanmar", "NA": "Namibia", "NR": "Nauru", "NP": "Nepal", "NL": "Netherlands", "NZ": "New Zealand", "NI": "Nicaragua", "NE": "Niger", "NG": "Nigeria", "NO": "Norway", "OM": "Oman", "PK": "Pakistan", "PW": "Palau", "PS": "Palestine", "PA": "Panama", "PG": "Papua New Guinea", "PY": "Paraguay", "PE": "Peru", "PH": "Philippines", "PL": "Poland", "PT": "Portugal", "PR": "Puerto Rico", "QA": "Qatar", "RO": "Romania", "RU": "Russia", "RW": "Rwanda", "SA": "Saudi Arabia", "SN": "Senegal", "RS": "Serbia", "SC": "Seychelles", "SL": "Sierra Leone", "SG": "Singapore", "SK": "Slovakia", "SI": "Slovenia", "SB": "Solomon Islands", "SO": "Somalia", "ZA": "South Africa", "ES": "Spain", "LK": "Sri Lanka", "SD": "Sudan", "SR": "Suriname", "SZ": "Swaziland", "SE": "Sweden", "CH": "Switzerland", "SY": "Syria", "TW": "Taiwan", "TJ": "Tajikistan", "TZ": "Tanzania", "TH": "Thailand", "TL": "Timor-Leste", "TG": "Togo", "TO": "Tonga", "TT": "Trinidad and Tobago", "TN": "Tunisia", "TR": "Turkey", "TM": "Turkmenistan", "UG": "Uganda", "UA": "Ukraine", "AE": "United Arab Emirates", "GB": "United Kingdom", "US": "United States", "UY": "Uruguay", "UZ": "Uzbekistan", "VU": "Vanuatu", "VE": "Venezuela", "VN": "Vietnam", "YE": "Yemen", "ZM": "Zambia", "ZW": "Zimbabwe"
  };
  return countries[code] || code;
}
