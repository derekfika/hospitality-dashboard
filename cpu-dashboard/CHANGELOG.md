# CPU Dashboard Changelog

## 2026.06.20.14

### Changed

- Reworked the top header: Fika wordmark left, CPU title centred and build/load
  status right.
- Moved Reload Sheet, Scan Calendars and Kitchen Mode into week controls.
- Moved Today between previous/next-week arrows.
- Hid the Print button while retaining print code and styles.
- Replaced external logo loading with inline SVG stored in configuration.
- Removed persistent wait-cursor behaviour.

## 2026.06.20.3 – 2026.06.20.13

### Added

- Prep progress bar and quick-prep controls.
- Updated-only, hide-completed and weekday focus filters.
- Dietary production summary.
- Urgent-order highlighting.
- Persistent user preferences and keyboard shortcuts.
- Full-screen Kitchen mode with automatic refresh.
- Dedicated Deliveries view and calendar delivery badges.
- Prepped/Done state with Fika-purple styling.
- Changed-booking tracking and detailed change alerts.
- Chef-focused categorised production board.
- Inline Fika SVG branding.

### Improved

- Static filter/search toolbar.
- Faster incremental calendar scans with larger chunks and batched writes.
- Immediate editor refresh after quote re-reading.
- Product categories configurable through `CPU Settings`.
- Site identity and delivery destination made more prominent.

## 2026.06.19.1 – 2026.06.20.7

### Added

- Initial CPU production dashboard.
- Calendar, run-sheet and production views.
- Site, service and product filtering.
- Calendar scan feedback modal.
- Quote and booking-form attachment parsing.
- Needs Attention workflow and manual quote re-reading.
- Built-in site email directory.

### Parser coverage

Added support for numerous legacy site quote formats, including:

- DOCX text boxes and split XML runs
- Priced packages and per-person calculations
- Quantity-first and `X10` quantity formats
- Product headings followed by price calculations
- Delegate packages and working lunches
- BBQ packages inheriting booking pax
- Vegan/soy/sesame products without misclassifying them as dietary notes
- Short dietary splits such as `3x vege`
- Date/month lines excluded from production products

The parser test suite currently contains 27 regression tests.
