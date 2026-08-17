# ERPC DSM Latest Data Updater

Checks the ERPC DSM/SRAS/TRAS/SCUC settlement-account page, finds the
latest normal DSM account, downloads its `Data Files 1` Excel file, extracts
`NEA-Bihar` and `NVVN-Nepal`, and inserts new data at the top of Google Sheets.

## Required Google Sheet tabs

- `NEA-Bihar`
- `NVVN-Nepal`
- `_CONTROL`

The script creates missing tabs automatically.

## GitHub Secrets

Add:
- `GOOGLE_SHEET_ID`
- `GOOGLE_SERVICE_ACCOUNT_JSON`

Share the Google Sheet with the service-account email as Editor.

## Run

Actions → ERPC DSM Latest Data Update → Run workflow

The workflow is also scheduled daily at 10:00 AM Nepal time (04:15 UTC).

This package is for the latest/new-data updater. Historical 2024-onward
backfill should be run separately after the latest updater is confirmed.
