# Munich RE Hot Drink Reporting

## Assumptions

- This is a separate desktop web app from the barista tally app in `../munich-hot-drinks`.
- The Apps Script project is either bound to the same Google Sheet as the tally app, or `setHotDrinkSpreadsheetId("SHEET_URL_OR_ID")` is run once.
- The dashboard is designed for desktop reporting and admin use.
- User email is only captured when Google allows it; otherwise the app stores a device/browser label.
- Undo marks the latest active row as `UNDONE` and writes an audit entry, preserving the original row.
- Test data clearing only removes rows whose `Source` or `Device/User` contains `TEST`.

## Files

- `Code.gs` serves the reporting web app, menu, setup, public config and settings admin.
- `Config.gs` contains sheet names, headers, floors, drinks and heatmap buckets.
- `SheetService.gs` provides shared log access and audit helpers.
- `Reporting.gs` builds dashboard summaries, heatmaps, trend data and CSV exports.
- `PdfReport.html` renders the polished client-ready PDF report for the selected date range.
- `Index.html`, `Styles.html`, and `JavaScript.html` contain the desktop reporting and admin UI.

## Deploy

1. Create a separate Apps Script project for reporting, distinct from the tablet tally deployment.
2. Add these files to the reporting Apps Script project.
3. Run `setHotDrinkSpreadsheetId("SHEET_URL_OR_ID")` to connect it to the same Sheet as the tally app, unless it is bound to that Sheet.
4. Run `setupHotDrinkTally` once and approve permissions.
5. Deploy > New deployment > Web app.
6. Set “Execute as” to “Me”.
7. Set access to the reporting/admin audience.
8. Share this reporting URL with desktop users only.

## Sheet Tabs

The setup function creates:

- `Drink_Log`
- `Settings`
- `Dashboard_Data`
- `Audit_Log`

The tally screen is deployed separately from `../munich-hot-drinks`; this app is for reporting, CSV export and settings maintenance.
