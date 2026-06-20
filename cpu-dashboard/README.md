# CPU Production Dashboard

A Google Apps Script web app for the central production unit. It reads hospitality
events from multiple Google Calendars, follows the attached quote and original
booking form, caches the parsed order, and presents a chef-first production view.

## What the first version includes

- This-week booking count, total pax, busiest day, next-week forecast and exceptions
- Seven-day calendar with site, service, time and pax visible at a glance
- Site checkboxes and free-text filtering
- Calendar, chronological run-sheet and consolidated production views
- Order drawer with menu lines, dietary information, notes and source documents
- Missing/unreadable attachment warnings
- Hourly refresh trigger
- Demo data when viewed outside Apps Script or before calendars are configured

## Setup

1. Create a new standalone Google Apps Script project attached to a Google Sheet.
2. Copy this folder's files into that project (or configure a separate `.clasp.json`).
3. In Apps Script, enable the advanced Calendar API and Drive API.
4. Run `setupCpuDashboard()` once and approve access.
5. Open the `CPU Settings` sheet. The current central calendar is preconfigured as:

```json
[{"id":"cpux@fikacatering.com","name":"CPU Hospitality Calendar"}]
```

This belongs in `CALENDARS_JSON`. If the calendar is shared under a different
Google Calendar ID, replace the email with that ID.

6. The project includes the authoritative FIKA site email directory. Site assignment
   uses the Calendar event creator/organiser email. `SITES_JSON` is optional and can
   override names, codes, colours or aliases. Run `installCpuSiteDirectory()` if you
   want the full built-in directory written visibly into the Settings sheet.

```json
[
  {
    "name": "One Angel Court",
    "code": "OAC",
    "colour": "#4F34C7",
    "aliases": ["Angel Court", "OAC"],
    "emails": ["seven@fikacatering.com"]
  },
  {
    "name": "London Place",
    "code": "LP",
    "colour": "#EF6C35",
    "aliases": ["LP"]
  }
]
```

7. Run `scanCpuCalendars()` to import the first set of orders.
8. Optionally run `installCpuRefreshTrigger()` for hourly scans.
9. Deploy as a web app with domain-only access.

## Important assumptions

- The current workflow uses the shared `cpux@fikacatering.com` calendar.
- The scanner also supports multiple calendars later through `CALENDARS_JSON`.
- Site identity is read from the quote/event location and matched against
  `SITES_JSON`. When the Calendar event creator/organiser email matches a site's
  `id` or `emails`, that email match takes priority over the written location.
- A quote is sufficient when it contains the production information. The original
  booking form is optional.
- Google Docs and Microsoft Word `.docx` quotes are parsed automatically.
- Calendar titles follow the existing pattern
  `SITE_Company_Service x Pax`, which is used as a safe fallback.
- Google Sheets/XLSX booking forms can also be parsed directly.
- PowerPoint `.pptx` quotes are converted temporarily and read as Google Slides.
- Drive links pasted into the calendar event description are detected even when
  they were not added through Calendar's formal attachment control.
- PDF-only quotes are linked but flagged for manual checking in this version.

## File access requirement

The Google account that deploys/runs this Apps Script must have permission to open
the attached files. Seeing a file link on a shared calendar does not automatically
grant Drive access. For reliable scanning, site teams should share quote and booking
files with the scanner account (or a Google Group containing it) before adding them
to the CPU calendar.

## Recommended next integration step

The strongest long-term contract is to add a compact machine-readable booking JSON
block to the calendar event description when the site dashboard creates the event.
The CPU dashboard can then use that as its primary data and keep attachments as the
human-readable source documents. This avoids relying solely on document layout.
