# YouTube Data Toolkit 2

A Google Sheets sidebar for collecting YouTube video metadata, playlist videos, channel uploads and a channel's playlist index. Jobs persist between sidebar sessions and write into new output tabs so source cells remain unchanged.

## What changed

- One canonical source/build replaces the ambiguity between two older script versions.
- New sidebar with date inputs, progress counts, pause/resume, cancel, retry and refresh actions.
- Playlist and channel imports filter by **video publication time**, not the date it was added to a playlist. UTC start/end dates are inclusive, including the entire final day. One-sided ranges are supported. Playlist-index mode filters playlist creation dates.
- **Short-video exclusion is enabled by default.** Keeps only videos longer than 180 seconds. YouTube's public Data API does not expose a reliable Shorts classification flag, so this conservative safeguard also excludes short regular videos, live/upcoming broadcasts and unknown durations. It does not infer classification from thumbnails, titles or `#shorts`. Turn it off to retain all lengths.
- Playlist/channel duration exclusions happen before rows are added to the results. Direct video-input jobs retain a FILTERED audit row with the input ID and reason, without the excluded video's metadata.
- Persistent per-user/per-spreadsheet job state, user locks and one API batch per execution.
- Pagination checkpoints with duplicate prevention when a page is replayed after interruption.
- Up to 50 IDs per metadata request, bounded transient retries with backoff, and immediate pause on quota failures. Failed requests count toward the attempt counter.
- Missing statistics stay blank. Zero remains zero. Unavailable videos are not automatically labelled geoblocked.
- Separate original publication timestamps, scheduled/live start timestamps, UTC and local timestamps.
- Protection against formula-like external titles/descriptions being interpreted as spreadsheet formulas.

## Install into a new or existing Google Sheet

1. Open the target sheet, then **Extensions > Apps Script**.
2. Remove the default placeholder and paste `dist/Code.gs` into `Code.gs`.
3. Add an HTML file named **Sidebar**, then paste `dist/Sidebar.html`.
4. In Project Settings enable **Show appsscript.json manifest file in editor**. Replace its content with `dist/appsscript.json`.
5. Under Services add **YouTube Data API v3**. If using a standard Cloud project, enable YouTube Data API v3 in that project too.
6. Save and reload the spreadsheet. Open **YouTube Toolkit > Open workspace**.
7. Complete Google's authorization screen yourself. The manifest requests access to the current spreadsheet, read-only YouTube data and the spreadsheet sidebar. No API key is stored in the code.
8. Paste links, select an import type and date range, then choose **Create collection** and **Run collection**.

**Do not install both original root scripts with the new build.** The old `toolkit.js` and `youtube_toolkit.js` remain in the repository as historical references. Only the three files in `dist/` belong in the Apps Script project.

## Use

For a column of video links, select only its populated data cells and choose **Use selected cells**. For playlists paste a playlist URL/ID. For channels use a handle, channel URL/ID or supported legacy username URL. Every new collection creates its own output tab. The active job belongs to the current Google account and spreadsheet. Other editors can see output tabs according to the spreadsheet's sharing settings.

Pause stops after the current batch. Closing the sidebar stops browser-driven continuation. Reopen and reload to resume. There are no background triggers. Each execution does one list page, one channel resolution or a batch of 50 metadata rows. Start a new job only after completing/cancelling the current one. Cancel preserves collected results.

**Retry unavailable** revisits unavailable video rows. **Refresh metadata** clears old metadata and queues all valid video IDs again. Index jobs require a new import to refresh. If output headers or the output tab are removed, the job stops instead of writing to a different location.

Output fields: input, video ID, row status, title, channel name/ID, publication UTC/local, duration seconds/display, views/likes/comments, broadcast status, scheduled/actual start, country code, description, retrieval UTC and error reason. Playlist-index jobs use the ID column for playlist IDs and the publication column for creation time.

## Limits and data care

- 5,000 input/output rows per collection. Large/wide descriptions can increase Sheets latency.
- The request counter measures this job's attempted supported list calls, not remaining daily account quota. Other apps, jobs and users can share quota. Duration exclusion adds video-detail calls during playlist listing.
- API calls can still fail, be slow or hit Apps Script/service limits. A crashed invocation can repeat requests even when output rows remain idempotent.
- Imports read a changing playlist over multiple pages, not an immutable snapshot. Unknown dates are omitted when date filters are active. Duplicate playlist videos are collapsed to one result.
- UTC/local dates are explicit timestamp text. The project timezone defaults to Asia/Kolkata and can be changed in the manifest.
- Refresh or remove stored API metadata within YouTube's applicable retention requirements. Refresh is manual, and this project does not implement automatic deletion or enforce a retention schedule. Remove output tabs and any exported copies when no longer needed. Revoking access is available through Google Account permissions.
- This collects metadata only. It does not download videos, circumvent geographic restrictions, provide private analytics, determine official Shorts status or create independently inferred engagement metrics.

## Development and validation

Node.js 20+ is sufficient. No npm packages are required.

```sh
npm test
npm run build
npm run benchmark
node scripts/preview.cjs
```

`src/Core.js` contains pure parsing, validation and batching. `src/Code.js` adapts those functions to Apps Script. `src/Sidebar.html` is the UI. `test/harness.cjs` models Apps Script services for offline tests.

The current suite contains 28 tests. The deterministic offline benchmark processes 1,000 mock video inputs using 20 video-list and 20 channel-list attempts, compared with 1,000 video-list calls for one-video-at-a-time retrieval. The **98% video-list call reduction is a mock scenario against an unbatched baseline**, not a measured improvement over the already-batched older `toolkit.js`. It is not a live API speed, quota or business-impact benchmark. See `docs/benchmark.json` and `docs/VALIDATION.md`.

The standalone preview uses clearly labelled sample data. It is not an authenticated Apps Script deployment. Live account authorization, API response differences and real Sheets execution require the installation acceptance checks in the validation guide.

## Official references

- [videos.list](https://developers.google.com/youtube/data/reference/rest/v3/videos/list)
- [Playlist publication timestamps](https://developers.google.com/youtube/data/reference/rest/v3/playlistItems)
- [Three-minute Shorts criteria](https://support.google.com/youtube/answer/15424877)
- [Apps Script quotas](https://developers.google.com/apps-script/guides/services/quotas)
- [YouTube developer policies](https://developers.google.com/youtube/terms/developer-policies)
