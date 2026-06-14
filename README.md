# Hospitality Dashboard

Google Apps Script dashboard for scanning hospitality booking forms from Gmail, writing bookings to a Google Sheet, generating quotes, creating calendar events, and managing booking status.

## Current Shape

- `00_config.gs` contains code defaults and the Settings sheet helpers.
- `02_Schema.gs` defines the booking object and validation.
- `04_Parser.gs` parses Angel Court booking spreadsheets.
- `05_GmailScanner.gs` scans Gmail for `.xlsx` booking forms.
- `06_DataLayer.gs` reads and writes dashboard rows.
- `07_Webapp.gs`, `Index.html`, `Script.html`, `Styles.html`, and `Icons.html` power the dashboard UI.
- `09_QuoteEngine.gs` creates, prints, confirms, and cancels quotes/bookings.
- `10_Calendar.gs` creates CPU calendar events.
- `11_Triggers.gs` contains setup and trigger helpers.

## Settings

The dashboard now has a Settings modal. Opening it seeds a `Settings` sheet with:

- `Key`
- `Value`
- `Section`
- `Label`
- `Type`
- `Notes`

The modal currently covers General, Inbox Scan, Quotes, Calendar, Branding, and Colours. Code defaults remain in `CONFIG`, but saved Settings sheet values take priority for operational fields.

Each tab can be reset to its code defaults from the modal, then saved back to the Settings sheet.

You can also run `initialiseDashboardSettings` from Apps Script, or use the spreadsheet menu:

`Hospitality Dashboard` -> `Initialise settings`

## Importing A Settings Draft

If you have a draft settings sheet, create a tab named `Settings Draft` in the bound spreadsheet.

The importer needs at least these columns, or close equivalents:

- `Key`
- `Value`

For example, `Setting`, `Setting Key`, or `Name` can be used instead of `Key`, and `Current Value`, `Setting Value`, or `Configured Value` can be used instead of `Value`.

Optional planning columns such as `Section`, `Tab`, `Label`, `Type`, and `Notes` can be present, but only the recognised key and value columns are imported. Unknown keys are ignored and reported.

Use the spreadsheet menu:

`Hospitality Dashboard` -> `Import settings draft`

The import still runs through the same validation as the Settings modal.

## Deployment Notes

This project is managed with clasp and points at the Apps Script project in `.clasp.json`.

Useful commands:

```powershell
npx.cmd clasp status
npx.cmd clasp push
```

On this Windows workspace, `npx` may be blocked by PowerShell execution policy, so use `npx.cmd`.

## Before Daily Use

Run through these flows in the deployed web app:

- In Apps Script, run `runDashboardPureTests` and confirm it returns `ok: true`.
- In the deployed web app console, run `runDashboardClientTests()` and confirm it returns `ok: true`.
- In Apps Script, run `testDashboardSettingsValidation` and confirm it returns `{ ok: true }`.
- Open Settings, change a harmless value, save, refresh, and confirm it persists.
- Scan inbox preview and process a small batch.
- Generate a quote.
- Print a quote only when the printer email is confirmed.
- Create a calendar event.
- Confirm a booking email.
- Cancel a booking with and without email/calendar removal.

You can also run `getDashboardLiveTestChecklist` in Apps Script for the live-service checklist covering scanner import, duplicate prevention, quotes, calendar events, settings persistence, email sending, archive behaviour, and filters.
