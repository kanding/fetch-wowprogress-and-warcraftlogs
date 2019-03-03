Fetch players looking for a guild on https://www.wowprogress.com/ and their corresponding logs from https://www.warcraftlogs.com/.

Written in Google Apps Script `.gs` and works solely through Google Sheets as an associated script.

![Alt text](https://i.imgur.com/YDTSSl2.png "Drive")

## Install

1. Import the `.xlsx` file from `sheet/` to Google Drive.
2. Open up Tools > Script Editor
3. Copy the contents of script.js into Code.gs

## Notes

This script relies on Google Sheet functions such as `IMPORTHTML` for web scraping, which can be unreliable. There is a hidden sheet which contains a lot of raw data that should only be edited with caution.

### Dev

In the future should be possible to get everything through [Google Drive REST API](https://developers.google.com/drive/api/v2/reference/)

Should be possible to partition players looking for guild in other regions than EU (OC, US).
