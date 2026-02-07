# air-condition

Google Apps Script project for logging room conditions to a Google Spreadsheet and toggling a SwitchBot-controlled air conditioner based on temperature thresholds.

## Overview

- `doGet` logs temperature/humidity into the `conditions` sheet or returns JSONP data.
- `turnOnAC` reads the latest temperature and `setting` sheet thresholds to decide whether to toggle heating or cooling.
- Status is stored in the `status` sheet to avoid repeated toggles.

## Spreadsheet setup

The script expects the following sheets in the spreadsheet referenced by `SPREADSHEET_ID` in `src/Code.js`:

- `conditions`: stores timestamp, temperature, humidity. The first row is treated as headers.
- `status`: cell A1 stores the current state (`off`, `hot`, `cool`).
- `setting`: first row contains headers like `hot_buttom`, `hot_up`, `cool_buttom`, `cool_up`, `hot`, `cool`; the second row contains the values.

## Endpoints

### Log sensor data

```
GET /exec?t=<temperature>&h=<humidity>
```

### Update status directly

```
GET /exec?s=status&t=<status>
```

### JSONP export

```
GET /exec?callback=<callback>
```

## OpenAPI

See `openapi.yaml` in the repository root for the OpenAPI 3.0 definition of the `/exec` endpoint.

## Scripts

- `npm test` runs the Vitest suite.
- `npm run build` copies `src/Code.js` into `dist/` for Apps Script deployment.
- `npm run deploy` pushes with clasp.

## Development

1. Update `SPREADSHEET_ID` in `src/Code.js` to match your spreadsheet.
2. Install dependencies with `npm install`.
3. Run tests with `npm test`.
