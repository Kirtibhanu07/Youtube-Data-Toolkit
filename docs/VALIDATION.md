# Validation and installation acceptance

## Offline verification

Run `npm test`, `npm run build` and `npm run benchmark`. No live YouTube requests are made by these commands. All tests use authored API responses and in-memory service mocks. They validate logic, not Google authorization or production service compatibility.

Coverage includes supported URLs and handles, host validation, exact input order, UTC end-date inclusion, leap dates, missing counts versus zeros, duration parsing, three-minute filtering, new-sheet output isolation, stale callbacks, bounded retries, quota pausing, resume without rewriting successful rows, playlist page replay, repeated-token loops, cancellation, refresh and playlist indexing.

The API-call benchmark's unbatched baseline is defined as one videos.list request per input ID. It does not benchmark the existing batched script, network latency, API accuracy or analyst hours saved. Do not present the mock count reduction as a production KPI.

## Live acceptance before calling it deployed

1. Install only the dist files in a bound Apps Script project and add the YouTube service.
2. Authorize using the intended Google account. Confirm the spreadsheet menu opens the sidebar.
3. Run a few public video URLs whose current metadata you can verify directly on YouTube. Confirm output order and the source sheet remains unchanged.
4. Import a playlist using a narrow UTC publication-date range. Verify a video's publication date is used rather than its playlist insertion date.
5. Keep the duration filter enabled and compare a known sub-three-minute video and a longer video. The filter is conservative and is not official Shorts classification.
6. Pause after a batch, close/reopen the sidebar and resume. Verify no duplicate output rows. Check unavailable entries and empty hidden statistics.
7. Run a channel handle and playlist-index job. Check resolved uploads and playlist IDs.
8. Refresh the completed video collection. Confirm retrieval timestamps change and old fields are cleared until new values arrive.
9. Check the Google project execution log and quota dashboard. The sidebar counter is job-scoped and counts attempts.

Live acceptance has not yet been completed in this workspace. Do not claim a live deployment or successful live collection until these steps pass.
