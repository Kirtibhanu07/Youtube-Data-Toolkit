# Youtube-Data-Toolkit
A high-performance Google Apps Script designed to turn Google Sheets into a professional YouTube data dashboard. It automates the retrieval of video metadata, playlist contents, and channel uploads while handling API quotas efficiently through batch processing.

# Key Features 
- Bulk Metadata Extraction: Fetch titles, durations, views, likes, and comment counts for hundreds of video URLs at once.
- Intelligent Batching: Optimized to process videos in groups of 50, significantly reducing YouTube API quota consumption (10,000 units/day).
- Playlist & Channel Downloader: Export every video from a specific playlist or an entire channel directly into your sheet.
- Smart Date Filtering: Includes a built-in HTML date picker to filter channel uploads by specific UTC date ranges.
- Robust Error Handling: A dedicated "Retry" tool scans for failed rows or geoblocked videos and attempts to re-fetch them without duplicating successful data.
- Timezone Aware: Automatically converts UTC timestamps to both GMT and your local script timezone for easy scheduling analysis.

# 🛠️ Setup Instructions
- 1. Create a Google Sheet: Open a new or existing Google Sheet.
Open Apps Script: Go to Extensions > Apps Script.

- 2. Copy the Code: Paste the provided YT DATA - APR FINALISED.js code into the editor.
Enable YouTube API:

- 3. In the Apps Script editor, click the + next to Services.
Select YouTube Data API v3 and click Add.

- 4. Run: Save the script and refresh your Google Sheet. A new menu "🚀 YouTube Data Toolkit" will appear.
 
# Tech Stack
- Language: Google Apps Script (JavaScript)
- API: YouTube Data API v3.
- UI: HTML5/CSS3 for the custom modal date picker.
