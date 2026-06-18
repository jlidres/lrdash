# Data Source and Snapshot

## Live source

Live data is loaded from the Google Sheet configured in `app.js`:
- `SHEET_SOURCE.sheetId`
- `SHEET_SOURCE.gid`

The app requests the sheet via Google Visualization JSON endpoint.

## Snapshot fallback

When live load fails, the app loads `data/snapshot.js` as fallback.

`data/snapshot.js` defines:
- `window.DASHBOARD_SNAPSHOT.generatedAt`
- `window.DASHBOARD_SNAPSHOT.source`
- `window.DASHBOARD_SNAPSHOT.table`

## Refreshing snapshot

Run:

`node sync-sheet.mjs`

This script downloads the current sheet and overwrites `data/snapshot.js`.
