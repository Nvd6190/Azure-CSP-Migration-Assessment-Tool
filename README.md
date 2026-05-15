# Azure CSP Migration Assessment Tool

![Version](https://img.shields.io/badge/version-1.1.0-blue) ![License](https://img.shields.io/badge/license-MIT-green) ![Node](https://img.shields.io/badge/node-%3E%3D18-brightgreen) ![Angular](https://img.shields.io/badge/angular-17-red)

Upload an Azure resource export (.xlsx) and instantly assess migration readiness across three modes — **Subscription Move**, **Region Move**, and **Jio Region Availability** — powered by 790+ live rules fetched from Microsoft's official documentation.

## Architecture

```
AzureCSP-Migration/
├── backend/                                  # Node.js/Express API
│   ├── src/
│   │   ├── index.js                          # Express server entry point
│   │   ├── routes/upload.js                  # All API endpoints
│   │   ├── services/
│   │   │   ├── migrationService.js           # Core assessment engine + type normalizer
│   │   │   ├── rulesFetcher.js               # Live rules fetcher from Microsoft Learn
│   │   │   └── excelReportBuilder.js         # Rich multi-sheet Excel report generator
│   │   └── data/
│   │       ├── azureMoveMatrix.json          # Static rules (manually verified, 300+ entries)
│   │       ├── jio-availability.json         # Jio India West service & VM availability
│   │       └── learn-rules-cache.json        # Disk cache of fetched Microsoft Learn rules
│   ├── Dockerfile
│   └── package.json
├── frontend/                                 # Angular 17 standalone app
│   ├── src/app/
│   │   ├── app.component.ts                  # Shell with Azure topbar + rules info
│   │   ├── landing/landing.component.ts      # Mode selection (3 assessment cards)
│   │   ├── upload/upload.component.ts        # Drag & drop file upload
│   │   ├── results-table/results-table.component.ts  # Interactive results grid
│   │   ├── models/assessment.model.ts        # TypeScript interfaces
│   │   └── services/migration.service.ts     # Angular HTTP service
│   ├── proxy.conf.json
│   ├── nginx.conf
│   ├── Dockerfile
│   └── package.json
├── docker-compose.yml
└── README.md
```

## Quick Start (Local Development)

### Backend
```bash
cd backend
npm install
npm run dev
```
API runs on http://localhost:3000

### Frontend
```bash
cd frontend
npm install
npx ng serve
```
App runs on http://localhost:4200 (proxies API calls to :3000)

## Quick Start (Docker)

```bash
docker-compose up --build
```
Open http://localhost:4200

## Assessment Modes

| Mode | Output Column | Description |
|------|--------------|-------------|
| **Subscription Move** | `SUBSCRIPTION MOVE SUPPORTED` | Can the resource be moved to a different subscription (CSP migration)? Values: Yes / No / Conditional / Review |
| **Region Move** | `REGION MOVE SUPPORTED` | Can the resource be moved to a different Azure region? Values: Yes / No / Review |
| **Jio Region Availability** | `JIO REGION AVAILABLE` | Is the service/VM available in Jio India West? Detects if resource is in an India region. Values: Yes / No / Review |

## How It Works

1. **Choose Mode** — Select Subscription Move, Region Move, or Jio Region Availability
2. **Upload** — Drag & drop your Azure resource export Excel file
3. **Assess** — Backend reads the RESOURCE TYPE column and matches each resource against 790+ rules from Microsoft's official [resource move support page](https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources)
4. **View** — Results grid shows colour-coded badges: ✅ Yes / ❌ No / ⚠️ Conditional / 🔍 Review
5. **Download** — Get a rich multi-sheet Excel report with assessment data, summary dashboard, pivot tables, and action items

## Features

- **Three assessment modes** — Subscription move, region move, Jio region availability
- **790+ live rules** — Auto-fetched from Microsoft Learn on startup and refreshed every 6 hours
- **Disk caching** — Falls back to cached rules when offline; static JSON overrides for manually verified corrections
- **Smart type normalizer** — 150+ display name → ARM type mappings, parent type matching, suffix matching, fuzzy matching
- **Drag & drop upload** — .xlsx, .xls, .csv supported (max 10 MB)
- **Interactive results grid** — Clickable summary cards, real-time text search, colour-coded status badges
- **Rich Excel reports** — 7 sheets: Assessment Data, Summary Dashboard, Pivot by Provider, Pivot by Resource Group, Pivot by Location, Status Sheet, Action Sheet
- **Jio data management** — Upload updated Jio availability Excel to refresh service/VM data (220+ services, 124 VM series)
- **Manual refresh** — Re-fetch rules from Microsoft Learn via the toolbar button

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| POST | `/api/assess` | Upload Excel, download assessed Excel directly |
| POST | `/api/assess-json` | Upload Excel, get JSON results for subscription move |
| POST | `/api/assess-region-json` | Upload Excel, get JSON results for region move |
| POST | `/api/assess-jio-json` | Upload Excel, get JSON results for Jio availability |
| GET | `/api/download/:id` | Download a previously generated assessment file |
| GET | `/api/rules` | Return current rules, counts, source, and metadata |
| POST | `/api/rules/refresh` | Force re-fetch live rules from Microsoft Learn |
| POST | `/api/jio/refresh` | Upload a new Jio availability Excel to update data |
| GET | `/api/health` | Health check |

## Live Rules

Rules are automatically fetched from the official [Microsoft Learn move support page](https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources) and merged with static overrides:

- **On startup** — fetches live rules and caches to disk
- **Every 6 hours** — periodic re-fetch via `setInterval`
- **Manual refresh** — via the refresh button in the toolbar or `POST /api/rules/refresh`
- **Fallback chain** — Live fetch → Disk cache → Static JSON

Static rules in `azureMoveMatrix.json` always win (manually verified corrections with detailed remarks).

## Excel Input Format

Your uploaded file should have at least these columns (case-insensitive, auto-detected):

| Column | Example |
|--------|---------|
| NAME | my-vm-01 |
| RESOURCE TYPE | Virtual machine |
| RESOURCE GROUP | rg-production |
| LOCATION | eastus |
| SUBSCRIPTION | My Subscription |

The RESOURCE TYPE column accepts both Azure portal display names (e.g. "Virtual machine") and ARM resource type IDs (e.g. "microsoft.compute/virtualmachines"). Previously assessed files can be re-uploaded — existing assessment columns are stripped automatically.

## Excel Output (Report Sheets)

| Sheet | Content |
|-------|---------|
| Assessment Data | All resources with colour-coded status, auto-filters, frozen header |
| Summary Dashboard | Title, metadata, distribution charts, summary table with percentages |
| Pivot by Provider | Breakdown by Azure provider |
| Pivot by Resource Group | Breakdown by resource group |
| Pivot by Location | Breakdown by Azure region |
| Status Sheet | Resources grouped by status (Yes / No / Conditional / Review) |
| Action Sheet | Actionable items grouped by status |

## Build Phases

- **Phase 1** ✅ — Excel upload, rules lookup, enriched Excel download
- **Phase 2** ✅ — Region move assessment + Jio region availability
- **Phase 3** ✅ — Live rules from Microsoft Learn with caching + periodic refresh
- **Phase 4** — Azure login + live validation with Azure move validation APIs

## Version History

See [CHANGELOG.md](CHANGELOG.md) for a detailed list of changes in each version.

| Version | Date | Highlights |
|---------|------|------------|
| 1.1.0 | 2026-05-15 | Periodic 6-hour rule refresh, Live/Static badge fix, README overhaul |
| 1.0.0 | 2026-05-15 | Initial release — 3 assessment modes, live rules, rich Excel reports |
