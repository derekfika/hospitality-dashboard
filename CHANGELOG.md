# Changelog

All notable changes to the Hospitality Dashboard project will be documented in this file.

The project follows a simple versioning approach:

* Major feature releases increment the minor version (e.g. v0.9.0 → v0.10.0)
* Bug fixes and stability improvements increment the patch version (e.g. v0.9.1 → v0.9.11)

---

## v0.9.11 - Code Review & Stability Update

### Fixed

* Calendar bookings now sort chronologically within each day.
* Fixed action button alignment and wrapping issues.
* Fixed calendar control visibility issues.
* Improved parser resilience around date and time handling.
* Improved booking validation behaviour.
* Prevented duplicate original booking form XLSX files from being created when calendar events are recreated.
* Fixed confirmation workflow to prevent accidental duplicate confirmation emails.

### Improved

* Calendar event creation now uses configured location settings rather than hard-coded values.
* Improved multi-site readiness through increased use of Settings sheet configuration.
* Additional code cleanup following external code review.
* Improved overall application stability and maintainability.

### Reviewed

* Full codebase reviewed across:

  * Configuration
  * Parser
  * Gmail scanning
  * Data layer
  * Quote generation
  * Calendar integration
  * Web application
  * Styling
  * Test harness

---

## v0.9.1 - Settings System Release

### Added

* Dynamic Settings modal.
* Settings schema system.
* Settings validation framework.
* Settings persistence to dedicated Settings sheet.
* Settings import and initialisation utilities.
* Calendar week view.
* Inbox Scan modal with progress tracking.
* Expanded test harness.

### Improved

* Dashboard branding and theming now configurable through settings.
* Application configuration moved from hard-coded values to live settings.
* Improved flexibility for future multi-site deployments.

### Fixed

* Various UI alignment and layout issues.
* Calendar rendering improvements.
* General dashboard usability improvements.

---

## v0.9.0 - Initial Dashboard Release

### Added

* Gmail inbox scanning for hospitality booking forms.
* XLSX attachment processing and conversion.
* Booking parser and validation engine.
* Dashboard Data sheet integration.
* Booking lifecycle management.
* Quote generation workflow.
* Calendar event creation workflow.
* Booking confirmation emails.
* Search, filtering and sorting.
* Inline booking editing.
* Status management system.
* Archiving workflow.
* Drive integration for generated documents.

### Notes

* First production-ready release of the Hospitality Dashboard platform.
* Built to automate hospitality booking processing and CPU workflow management.
