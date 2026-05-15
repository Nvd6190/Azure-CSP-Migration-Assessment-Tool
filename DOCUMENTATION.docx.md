# Azure CSP Migration Assessment Tool

## Project Documentation

---

## 1. Overview

The **Azure CSP Migration Assessment Tool** is a web-based application that helps Azure Cloud Solution Providers (CSPs) assess whether Azure resources can be:

1. **Moved across subscriptions** (Subscription Move)
2. **Moved across regions** (Region Move)
3. **Available in Jio India West region** (Jio Region Availability)

Users upload an Azure resource export Excel file, and the tool instantly evaluates each resource against Microsoft's move support matrix and Jio availability data, providing a detailed assessment with remarks and downloadable Excel reports.

---

## 2. Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                        FRONTEND (Angular 17)                     │
│                        http://localhost:4200                      │
│                                                                  │
│  ┌──────────┐  ┌──────────────┐  ┌────────────────────────────┐│
│  │ Landing  │  │   Upload     │  │     Results Table           ││
│  │  Page    │──│  Component   │──│  (Summary + Table + Excel)  ││
│  │(3 modes) │  │ (Drag&Drop)  │  │                            ││
│  └──────────┘  └──────────────┘  └────────────────────────────┘│
└───────────────────────────┬─────────────────────────────────────┘
                            │ HTTP (proxy /api → :3000)
┌───────────────────────────▼─────────────────────────────────────┐
│                        BACKEND (Node.js + Express)               │
│                        http://localhost:3000                      │
│                                                                  │
│  ┌─────────────┐  ┌────────────────────┐  ┌──────────────────┐ │
│  │  Routes     │  │  MigrationService  │  │  RulesFetcher    │ │
│  │  /api/*     │──│  (Assessment Core) │──│  (Live CSV from  │ │
│  │             │  │                    │  │   Microsoft)     │ │
│  └─────────────┘  └────────────────────┘  └──────────────────┘ │
│                              │                                   │
│                    ┌─────────┴──────────┐                        │
│                    │     Data Layer     │                        │
│                    │ • azureMoveMatrix  │                        │
│                    │ • jio-availability │                        │
│                    └────────────────────┘                        │
└─────────────────────────────────────────────────────────────────┘
```

---

## 3. Technology Stack

| Layer | Technology | Version |
|-------|-----------|---------|
| Frontend | Angular (Standalone Components) | 17.3 |
| Frontend Styling | SCSS (inline) | - |
| Backend | Node.js + Express | v24 + 4.18 |
| File Parsing | SheetJS (xlsx) | 0.18.5 |
| File Upload | Multer | 1.4.5 |
| Containerization | Docker + Docker Compose | 3.8 |
| Data Source | Microsoft GitHub CSV (live) + Static JSON | - |

---

## 4. Project Structure

```
AzureCSP-Migration/
├── docker-compose.yml          # Container orchestration
├── README.md
│
├── backend/
│   ├── package.json
│   ├── Dockerfile
│   ├── uploads/                # Temporary assessment files (auto-cleaned)
│   └── src/
│       ├── index.js            # Express server entry point
│       ├── routes/
│       │   └── upload.js       # All API routes
│       ├── services/
│       │   ├── migrationService.js   # Core assessment engine
│       │   └── rulesFetcher.js       # CSV parser from Microsoft GitHub
│       └── data/
│           ├── azureMoveMatrix.json  # Static subscription move rules (302 entries)
│           ├── jio-availability.json # Jio region data (220 services, 124 VMs)
│           └── jio-availability.xlsx # Source Excel for Jio data
│
└── frontend/
    ├── package.json
    ├── angular.json
    ├── proxy.conf.json         # Dev proxy: /api → localhost:3000
    ├── Dockerfile
    └── src/
        ├── index.html
        ├── assets/
        │   └── azure-logo.png
        └── app/
            ├── app.component.ts          # Root shell + topbar
            ├── landing/
            │   └── landing.component.ts  # Mode selection (3 cards)
            ├── upload/
            │   └── upload.component.ts   # Drag-and-drop file upload
            ├── results-table/
            │   └── results-table.component.ts  # Assessment results display
            ├── services/
            │   └── migration.service.ts  # HTTP service layer
            └── models/
                └── assessment.model.ts   # TypeScript interfaces
```

---

## 5. Assessment Modes

### 5.1 Subscription Move Assessment

Checks whether each Azure resource type can be moved across subscriptions within the same tenant.

| Field | Description |
|-------|-------------|
| **Data Source** | Microsoft's GitHub CSV (752 rules) + 302 static overrides |
| **Output Columns** | SUBSCRIPTION MOVE SUPPORTED, NORMALIZED TYPE, REMARKS |
| **Possible Values** | Yes, No, Conditional, Review |

### 5.2 Region Move Assessment

Checks whether each Azure resource type can be moved from one Azure region to another.

| Field | Description |
|-------|-------------|
| **Data Source** | Microsoft's GitHub CSV (752 region rules) |
| **Output Columns** | REGION MOVE SUPPORTED, NORMALIZED TYPE, REMARKS |
| **Possible Values** | Yes, No, Review |
| **Remarks** | 130+ resource-specific migration guidance with MS docs links |

### 5.3 Jio India West Region Availability

Checks whether Azure services are available in the Jio India West region (JIO).

| Field | Description |
|-------|-------------|
| **Data Source** | Jio availability Excel (220 services + 124 VM series) |
| **Output Columns** | JIO REGION AVAILABLE, JIO SERVICE NAME, NORMALIZED TYPE, REMARKS |
| **Possible Values** | Yes, No, Review |
| **Updatable** | Yes — via "↑ Jio Data" button in topbar |

---

## 6. API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/api/assess-json` | Subscription move assessment (returns JSON) |
| POST | `/api/assess-region-json` | Region move assessment (returns JSON) |
| POST | `/api/assess-jio-json` | Jio region availability (returns JSON) |
| POST | `/api/assess` | Subscription assessment (direct Excel download) |
| GET | `/api/download/:id` | Download previously generated Excel |
| GET | `/api/rules` | Get current rules metadata |
| POST | `/api/rules/refresh` | Force re-fetch rules from Microsoft CSV |
| POST | `/api/jio/refresh` | Upload new Jio availability Excel |
| GET | `/api/health` | Health check |

### Request Format (Assessment Endpoints)

```
POST /api/assess-json
Content-Type: multipart/form-data

file: <Azure resource export .xlsx file>
```

### Response Format

```json
{
  "summary": {
    "total": 45,
    "yes": 32,
    "no": 8,
    "review": 5,
    "conditional": 0
  },
  "resources": [
    {
      "NAME": "my-vm",
      "TYPE": "microsoft.compute/virtualmachines",
      "RESOURCE GROUP": "rg-prod",
      "LOCATION": "eastus",
      "SUBSCRIPTION MOVE SUPPORTED": "Yes",
      "NORMALIZED TYPE": "microsoft.compute/virtualmachines",
      "REMARKS": "Can be moved across subscriptions..."
    }
  ],
  "downloadId": "assessment-1715425200000"
}
```

---

## 7. Rules Engine

### How Rules Are Loaded

```
Startup:
  1. Load static rules from azureMoveMatrix.json (302 manually verified entries)
  2. Fetch live CSV from Microsoft GitHub (752 subscription + 752 region rules)
  3. Merge: CSV as base → static overrides win (manual corrections)
  4. Load Jio availability from jio-availability.json
```

### Microsoft CSV Sources

| Purpose | URL |
|---------|-----|
| Subscription Move | `https://raw.githubusercontent.com/tfitzmac/resource-capabilities/master/move-support-resources.csv` |
| Region Move | `https://raw.githubusercontent.com/tfitzmac/resource-capabilities/master/move-support-resources-with-regions.csv` |

### Duplicate Handling

Microsoft's CSV contains duplicate entries for some resources. The parser uses a "keep Yes" strategy:
- If a rule already shows `Yes` and a duplicate says `No`, the `Yes` is preserved.

### Type Normalization

The service handles multiple input formats:
- ARM types: `microsoft.compute/virtualmachines` ✓
- Display names: `Virtual Machine` → normalized to ARM type
- Common aliases: `VM`, `AKS`, `NSG` → mapped to ARM types
- Fuzzy matching: Parent type fallback, suffix matching

---

## 8. Jio Data Management

### Monthly Update Process

1. Click **"↑ Jio Data"** button in the top navigation bar
2. Upload the updated Jio availability Excel file
3. System parses the **Services** sheet and **VMs** sheet
4. JSON is regenerated and saved
5. In-memory data reloads immediately — no restart needed

### Expected Excel Format

| Sheet | Required Columns |
|-------|-----------------|
| Services | `Services Names`, `Availability Status` |
| VMs | `VM Series`, `Availability Status`, `Remarks` |

---

## 9. Running the Application

### Development Mode

```powershell
# Terminal 1: Backend
cd C:\Users\NirajDodia\AzureCSP-Migration\backend
npm install
node src/index.js
# → Running on http://localhost:3000

# Terminal 2: Frontend
cd C:\Users\NirajDodia\AzureCSP-Migration\frontend
npm install
npx ng serve --proxy-config proxy.conf.json
# → Running on http://localhost:4200
```

### Docker Mode

```powershell
cd C:\Users\NirajDodia\AzureCSP-Migration
docker-compose up --build
# Backend: http://localhost:3000
# Frontend: http://localhost:4200
```

---

## 10. User Workflow

```
1. Open http://localhost:4200
2. Choose assessment mode:
   ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐
   │  Subscription   │  │   Region Move   │  │  Jio Region     │
   │     Move        │  │                 │  │  Availability   │
   └────────┬────────┘  └────────┬────────┘  └────────┬────────┘
            │                    │                     │
            └────────────────────┴─────────────────────┘
                                 │
3. Upload Azure resource export Excel (.xlsx)
4. View results:
   • Summary cards (Total / Yes / No / Review)
   • Searchable/filterable results table
   • Click cards to filter by status
5. Download assessed Excel report
```

---

## 11. Excel Output

The downloaded Excel contains two sheets:

### Sheet 1: Assessment Results

All original columns from the uploaded file PLUS the assessment columns specific to the chosen mode. Old assessment columns from previous runs are automatically stripped.

### Sheet 2: Summary

| Metric | Value |
|--------|-------|
| Total Resources | 45 |
| Can Move (Yes) | 32 |
| Cannot Move (No) | 8 |
| Needs Review | 5 |
| Assessment Date | 2026-05-11T10:30:00Z |

---

## 12. Key Features

| Feature | Description |
|---------|-------------|
| **Live Rules** | Auto-fetches latest Microsoft CSV on startup |
| **Smart Type Detection** | Handles ARM types, display names, aliases, fuzzy matching |
| **Incremental Assessment** | Re-uploading assessed files strips old columns cleanly |
| **130+ Region Remarks** | Resource-specific migration guidance with documentation links |
| **Jio Monthly Updates** | Upload new Jio Excel anytime via UI — no code changes |
| **Docker Support** | One-command deployment with docker-compose |
| **Responsive UI** | Works on desktop and tablet screens |

---

## 13. Security Notes

- File uploads limited to 10 MB
- Only `.xlsx`, `.xls`, `.csv` extensions accepted
- Uploaded files are cleaned up after processing (or after 10 minutes for downloads)
- No authentication (internal tool) — add authentication if exposed externally
- CORS enabled for development

---

## 14. Troubleshooting

| Issue | Solution |
|-------|----------|
| Backend won't start | Check if port 3000 is in use: `Get-NetTCPConnection -LocalPort 3000` |
| Frontend proxy errors | Ensure backend is running before frontend |
| Rules show as "Static" | Microsoft CSV fetch failed; click refresh button |
| Region move always "Review" | Rules refresh may have failed; check console for errors |
| Jio upload fails | Verify Excel has sheets named "Services" and "VMs" with correct columns |
