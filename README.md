# Azure CSP Migration Assessment Tool

Upload an Azure resource export (.xlsx) and instantly see which resources support subscription-to-subscription moves based on Microsoft's official move matrix.

## Architecture

```
AzureCSP-Migration/
├── backend/                              # Node.js/Express API
│   ├── src/
│   │   ├── index.js                      # Express server
│   │   ├── routes/upload.js              # POST /api/assess & /api/assess-json
│   │   ├── services/migrationService.js  # Lookup engine + type normalizer
│   │   └── data/azureMoveMatrix.json     # Microsoft move matrix (120+ rules)
│   ├── Dockerfile
│   └── package.json
├── frontend/                             # Angular 17 standalone app
│   ├── src/app/
│   │   ├── app.component.ts              # Shell with Azure topbar
│   │   ├── upload/                       # Drag & drop upload component
│   │   ├── results-table/                # Filterable results grid
│   │   ├── models/                       # TypeScript interfaces
│   │   └── services/                     # API service
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
ng serve
```
App runs on http://localhost:4200 (proxies API calls to :3000)

## Quick Start (Docker)

```bash
docker-compose up --build
```
Open http://localhost:4200

## How It Works

1. **Upload** — Drag & drop your Azure resource export Excel file
2. **Assess** — Backend reads the TYPE column and matches each resource against 120+ rules from Microsoft's official [resource move support page](https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources)
3. **View** — Results grid shows colour-coded badges: ✅ Yes / ❌ No / ⚠️ Review
4. **Download** — Get the enriched Excel with three new columns: SUBSCRIPTION MOVE SUPPORTED, NORMALIZED TYPE, REMARKS

## Features

- **Drag & drop upload** — .xlsx, .xls, .csv supported (max 10 MB)
- **Smart type normalizer** — Maps Azure portal display names and ARM resource type IDs to the move matrix
- **Filter tabs** — All / Yes / No / Review
- **Search** — Filter by resource name, type, or resource group
- **Summary cards** — Total / Can Move / Cannot Move / Review at a glance
- **Excel download** — Enriched file with assessment results + summary sheet

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| POST | `/api/assess` | Upload Excel, download assessed Excel directly |
| POST | `/api/assess-json` | Upload Excel, get JSON results (used by frontend) |
| GET | `/api/download/:id` | Download a previously generated assessment file |
| GET | `/api/rules` | Return the current move matrix rules |
| GET | `/api/health` | Health check |

## Updating the Move Matrix

Edit `backend/src/data/azureMoveMatrix.json` when Microsoft updates their [resource move support page](https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources).

Each rule:
```json
{
  "subscriptionMove": "Yes | No",
  "remarks": "Optional migration notes"
}
```

## Excel Input Format

Your uploaded file should have at least these columns (case-insensitive):

| Column | Example |
|--------|---------|
| NAME | my-vm-01 |
| TYPE | Virtual machine |
| RESOURCE GROUP | rg-production |
| LOCATION | eastus |
| SUBSCRIPTION | My Subscription |

The TYPE column accepts both Azure portal display names (e.g. "Virtual machine") and ARM resource type IDs (e.g. "microsoft.compute/virtualmachines").

## Build Phases

- **Phase 1** ✅ — Excel upload, rules lookup, enriched Excel download
- **Phase 2** — Add region move supported and remarks columns
- **Phase 3** — Azure login + live validation with Azure move validation APIs
- **Phase 4** — PDF/Excel migration report with dependency grouping
