Fetch players looking for a guild on https://www.wowprogress.com/ and their corresponding logs from https://www.warcraftlogs.com/.

Written in Google Apps Script `.gs` and works solely through Google Sheets as an associated script.

![Alt text](https://i.imgur.com/YDTSSl2.png "Drive")

## Install

1. Import the sheet [by making a copy](https://docs.google.com/spreadsheets/d/1OxhQvHZgMaS_tMR1TRbIgh8Za-24IDRYxSYRhCJeRLk/copy).
2. Provide a valid WarcraftLogs Public API Key.
3. Click the `<fetch>` button and provide the script with authorization to edit the sheet.
4. If needed copy the contents of script.js into Code.gs in Tools > Script Editor.

## Notes

This script relies on Google Sheet functions such as `IMPORTHTML` for web scraping, which tend to be unreliable. There is a hidden sheet which contains a lot of raw data that should only be edited with caution.

A valid WarcraftLogs Public API Key is also required in order to query for log information. This can be found on your [profile page](https://www.warcraftlogs.com/profile).

### Dev

Should be possible to partition players looking for guild in other regions than EU (OC, US).
