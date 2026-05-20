# Changelog

All notable changes to this project will be documented in this file.

Format: [Semantic Versioning](https://semver.org/) — MAJOR.MINOR.PATCH

---

## [1.3.0] - 2026-05-20

### Added
- **GCP to Azure Migration** — New 5th assessment mode that maps Google Cloud Platform resources to Azure equivalents with migration guidance
- **GCP service mapping data** (`gcp-azure-mapping.json`) with 50+ service mappings across 16 categories (Compute, Containers, Serverless, Storage, Database, Analytics, AI/ML, Messaging, Networking, Security, Monitoring, DevOps, IoT, Integration, Governance, Migration)
- **SKU-level recommendations** for Compute Engine (60+ machine types), Cloud SQL (11 tiers), and Memorystore (8 tiers)
- **GCP assessor module** (`assessors/gcpAssessor.js`) — handles googleapis.com URLs, display names, canonical formats, and machine type extraction
- **GCP column detection** in `typeDetector.js` — auto-detects Service/ResourceName/SKU columns from GCP exports
- New API endpoint `POST /api/assess-gcp-json` for GCP resource assessment
- **6-sheet Google-branded Excel report** — Executive Summary, Service Mapping, By Category, By GCP Service, Migration Roadmap, Risk Matrix
- Google blue (#4285F4) themed landing card with GCP icon
- Results table supports GCP-specific columns: GCP Service, Azure Equivalent, Category, Similarity, SKU Recommendation, Migration Notes
- Frontend migration service adds `assessGcpFile()` HTTP method

### Changed
- Landing component now displays 5 mode cards (previously 4)
- App component mode badge includes "GCP → Azure" label
- Upload component shows GCP-specific header and description text
- Results table summary cards and data columns adapt for GCP mode (same layout as AWS with GCP labels)
- Version bumped to 1.3.0 across all packages and environment files

---

## [1.2.1] - 2026-05-20

### Added
- **Enhanced AWS Excel Report** — Completely redesigned with 6 dedicated sheets (up from 5), distinct from subscription/region format
- **Migration Roadmap** sheet — Phased migration plan (Phase 1-4) grouping services by migration wave with strategy, effort, and timeline per service
- **Risk Matrix** sheet — Priority-ranked risk table (P1, P2...) with risk level, impact assessment, strategy, and required actions
- **Migration Readiness Badge** on Executive Summary — auto-rated as "Ready to Migrate", "Needs Planning", or "Complex Migration"
- **Estimated Effort Breakdown** section on Executive Summary with phase timelines (1-2 weeks to 6-12+ weeks)
- **Migration Strategy** and **Effort** columns in Service Mapping data sheet (Rehost/Replatform/Refactor/Re-architect)
- **Example Services** column in By Category sheet showing top services per category
- **Readiness indicator** per namespace in By AWS Namespace sheet (✓ Ready / ○ Partial / ✗ Needs Work)
- **Azure Services** column in namespace sheet showing mapped Azure targets
- Percentage row under KPI cards in Executive Summary
- Namespace names displayed in UPPERCASE for clarity
- Totals row in By AWS Namespace sheet
- Auto-filter on By AWS Namespace sheet

### Changed
- AWS report now generates 6 sheets (was 5): Executive Summary, Service Mapping, By Category, By AWS Namespace, Migration Roadmap, Risk Matrix
- By Category expanded from 5 columns to 8 (added Strategy, Effort, Timeline, Example Services)
- Service Mapping expanded from 8 columns to 11 (added #, Migration Strategy, Effort)
- Executive Summary now includes scope details with namespace names and unique Azure service count
- Next Steps section now uses numbered steps referencing specific sheets
- Visual progress bars in By Category widened to 20 blocks for better resolution
- Version bumped to 1.2.1

---

## [1.2.0] - 2026-05-19

### Added
- **AWS to Azure Migration** — New 4th assessment mode that maps AWS resources to Azure equivalents with similarity scores and migration notes
- New API endpoint `POST /api/assess-aws-json` for AWS resource assessment
- AWS-Azure service mapping data file (`aws-azure-mapping.json`) with comprehensive service mappings
- Landing page updated to 4 assessment cards (new orange-themed AWS card)
- Results table supports AWS-specific columns: AWS Service, Azure Equivalent, Category, Similarity, Migration Notes
- Category badges for AWS mode: Direct Equivalent / Similar / Partial / No Mapping
- Isolated assessor modules — each assessment mode runs independently; errors in one don't crash others

### Changed
- Landing component now displays 4 mode cards (previously 3)
- Migration service extended with `assessAwsResources()` and `getAwsSummary()` methods
- Frontend migration service adds `assessAwsFile()` HTTP method
- Results table summary cards adapt dynamically to show AWS-specific categories
- README updated to document all 4 assessment modes and 10 API endpoints

---

## [1.1.0] - 2026-05-15

### Added
- Periodic rules re-fetch every 6 hours via `setInterval`
- Live/Static badge fix — correctly shows "Live" when rules are fetched from Microsoft Learn
- Version badge displayed in the application topbar (`v1.1.0`)
- Backend health endpoint now returns app version

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
