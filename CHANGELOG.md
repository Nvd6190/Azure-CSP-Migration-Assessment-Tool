# Changelog

All notable changes to this project will be documented in this file.

Format: [Semantic Versioning](https://semver.org/) — MAJOR.MINOR.PATCH

---

## [1.1.0] - 2026-05-15

### Added
- Periodic rules re-fetch every 6 hours via `setInterval`
- Live/Static badge fix — correctly shows "Live" when rules are fetched from Microsoft Learn

### Fixed
- Header badge was showing "Static" even when live rules were loaded (source string mismatch: expected `dynamic+static`, actual `microsoft-learn+static`)

### Changed
- README fully updated to reflect all current features, 3 assessment modes, 9 API endpoints, 7 Excel report sheets, and live rules documentation

---

## [1.0.0] - 2026-05-15

### Added
- **Subscription Move Assessment** — Upload Azure resource export, assess each resource against Microsoft's move matrix
- **Region Move Assessment** — Check if resources can be moved to a different Azure region
- **Jio Region Availability** — Check service/VM availability in Jio India West region
- Live rules fetching from Microsoft Learn (750+ subscription + 750+ region rules)
- Disk caching of fetched rules with offline fallback
- Static rules with manually verified corrections and detailed remarks (300+ entries)
- Smart type normalizer with 150+ display name → ARM type mappings
- Fuzzy matching (parent type, suffix, partial match)
- Drag & drop file upload (.xlsx, .xls, .csv, max 10 MB)
- Interactive results grid with clickable summary cards, text search, colour-coded badges
- Rich Excel report output (7 sheets: Assessment Data, Summary Dashboard, Pivot by Provider/RG/Location, Status, Actions)
- Jio data management via Excel upload
- Manual rules refresh button in toolbar
- Docker Compose setup for containerized deployment
- Angular 17 standalone frontend with proxy to Express backend
