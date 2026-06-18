# Maintenance

## Common updates

- Update visual layout in `styles.css`
- Update dashboard behavior in `app.js`
- Update static shell in `index.html`
- Refresh snapshot data via `sync-sheet.mjs`

## Cache/version bump

When updating frontend assets, keep query version strings aligned:
- `index.html` references `styles.css?v=...` and `app.js?v=...`
- `app.js` uses `SNAPSHOT_CACHE_KEY` for `data/snapshot.js`

Update these values when you need cache busting for fresh deploys.

## Data quality checks

Before publishing:
- confirm filters populate correctly
- confirm charts render with live and snapshot data
- confirm CSV export contains filtered rows and headers

## Operational notes

- The dashboard normalizes missing fields to fallback labels (for example `Unspecified`).
- Quantity parsing tolerates formatted numbers with commas.
- Date parsing supports Google Visualization dates, slash dates, and ISO-like strings.
