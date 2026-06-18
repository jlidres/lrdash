# Getting Started

## Requirements

- A modern browser (Chrome, Edge, Firefox, Safari)
- Internet access for live Google Sheet loading and Chart.js CDN

## Run locally

Serve the repository root with any static file server, then open `index.html`.

Example:
- `python -m http.server`
- open `http://localhost:8000/index.html`

## Data loading behavior

1. The app first attempts to load live data from Google Sheets.
2. If live loading fails, it falls back to `data/snapshot.js`.
3. If neither source is available, the dashboard stays in an error state.

## Embed mode

Add `?embed=1` to the URL (or use an iframe context) to switch to compact embed styling.
