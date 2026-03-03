# Dubai ERP - Formula-Based Automation System

> Complete documentation of everything built and configured for the Dubai Civil Engineering ERP system.

---

## 📋 Overview

This project is a **Google Sheets-based ERP system** for Dubai civil engineering project management. It uses a **formula-based automation approach** (like Excel) instead of triggers — all automations run via spreadsheet formulas injected by Google Apps Script.

### Key Files

| File | Purpose |
|------|---------|
| `generate_civil_data.py` | Python script — generates 17 CSV files with realistic Dubai construction data |
| `import_dubai_data.js` | Apps Script — imports CSV files from GitHub into Google Sheets |
| `erp_automation.gs` | Apps Script — formula-based automation engine (775 lines) |
| `erp_dashboard.js` | Apps Script — multi-page operational dashboard |

---

## 📊 Data Generation (`generate_civil_data.py`)

### 17 Tables Generated

| # | Table | Records | Key Fields |
|---|-------|---------|------------|
| 1 | Clients | 30 | client_id, client_name, client_type, status |
| 2 | Employees | 80 | employee_id, full_name, department, status |
| 3 | Contractors | 40 | contractor_id, company_name, specialty, rating |
| 4 | Suppliers | 30 | supplier_id, company_name, category, status |
| 5 | Equipment | 60 | equipment_id, equipment_type, status |
| 6 | Projects | 50 | project_id, project_name, status, health_status |
| 7 | Contracts | 100 | contract_id, contract_type, contract_value_aed |
| 8 | Project_Phases | 300 | phase_id, phase_name, progress_pct |
| 9 | Work_Packages | 400 | package_id, item_description, total_amount_aed |
| 10 | Permits_Approvals | 200 | permit_id, permit_type, expiry_date |
| 11 | Inspections | 500 | inspection_id, inspection_type, result |
| 12 | Safety_Incidents | 150 | incident_id, incident_type, severity |
| 13 | Payment_Applications | 300 | ipc_id, gross_value_aed, net_certified_aed |
| 14 | Variation_Orders | 200 | vo_id, reason, amount_aed |
| 15 | Purchase_Orders | 400 | po_id, po_type, status |
| 16 | Daily_Site_Reports | 500 | report_id, report_date, weather |
| 17 | Project_Documents | 600 | document_id, document_type, discipline |

### Post-Processing
- All empty/NaN cells filled with contextual defaults (68,852 cells filled)
- Defaults: `internal_notes` → "—", `issues` → "None reported", `mmup_license` → "N/A"
- Cross-table metrics calculated (project risk scores, client metrics)
- Full integrity verification (FK checks, governance columns)

### Running
```bash
cd /home/anandhu/projects/dubai-project-management-data
.venv/bin/python generate_civil_data.py
```

---

## ⚙️ ERP Automation Engine (`erp_automation.gs`)

### Architecture: Formula-Based (No Triggers)

All automation uses **spreadsheet formulas**, not `onEdit` triggers. Run `setupERPSystem()` once after importing data.

### 8 Modules

#### Module 0: Expand Sheets
- Inserts 100 extra rows below existing data on each sheet
- Prevents "range out of bounds" errors when applying formulas

#### Module 1: `_Lookups` Helper Sheet
- Auto-built reference sheet with Name+ID lists:
  - Column A-B: Employees (`Ahmed Ali (EMP-0001)` → `EMP-0001`)
  - Column D-E: Clients
  - Column G-H: Contractors
  - Column J-K: Suppliers
  - Column M-N: Projects

#### Module 2: Base Styling
- Dark header (`#2d3748`) with white text
- Alternating row colors (`#f7fafc` / white)
- Inter font, 10pt
- System columns (ID, created_at, record_version) highlighted gray
- Private columns (salary, contract_value) highlighted amber

#### Module 3: Dropdown Validation
- ~50 categorical columns across all 17 tables
- All values verified against actual CSV data (25+ missing values added)
- `setAllowInvalid(true)` — no red marks on existing data
- Clears old validations before applying new ones

#### Module 4: FK Dropdowns
- Foreign key columns get ID-only dropdowns
- Name reference available in `_Lookups` sheet
- Covers: project_id, client_id, contractor_id, employee_id, supplier_id

#### Module 5: Formula Injection (Batch-Optimized)
All formulas use **batch `setFormulas()`** (arrays written at once) instead of cell-by-cell, reducing API calls from ~10,000+ to ~50.

**Calculated Fields:**

| Sheet | Formula Fields |
|-------|---------------|
| Projects | `duration_days`, `elapsed_days`, `remaining_days`, `budget_variance_pct`, `profit`, `health_status` |
| Clients | `lifetime_revenue` (SUMIF), `active_projects` (COUNTIFS), `risk_projects` (COUNTIFS) |
| Contracts | `remaining_value` |
| Work_Packages | `total_amount_aed`, `executed_amount_aed` |
| Permits | `is_expired` (auto-check vs TODAY()) |
| Payments | `net_certified_aed` (gross - retention - deductions) |

**Auto-ID Formulas (new rows only):**
- Pattern: `=IF(trigger<>"", PREFIX & TEXT(COUNTA(...)+1, "0000"), "")`
- Each table has a "trigger column" (the main data field user fills first)

**Timestamp Formulas (new rows only):**
- `created_at` → `=IF(trigger<>"", NOW(), "")`
- `last_updated_at` → `=IF(trigger<>"", NOW(), "")`
- `created_by` → `anandhu7833@gmail.com`
- `record_version` → `1`
- `is_active` → `TRUE`

**Default Value Formulas (new rows only):**

| Table | Defaults |
|-------|----------|
| Clients | status=Active, tags=—, internal_notes=— |
| Employees | status=Active, visa_status=Valid |
| Contractors | status=Active, prequalified=FALSE, rating=C |
| Suppliers | status=Active, approved=TRUE |
| Equipment | status=Available |
| Projects | status=Awarded, completion_pct=0, risk_score=1 |
| Contracts | status=Active, retention=10%, performance_bond=10%, defect_liability=12mo |
| Project_Phases | status=Not Started, progress_pct=0 |
| Work_Packages | executed_qty=0 |
| Permits | status=Pending |
| Inspections | result=Pending, defects_found=0, corrective_action=N/A |
| Safety_Incidents | severity=Low, status=Open, injured_persons=0, lti_days=0 |
| Payment_Applications | status=Submitted, deductions_aed=0 |
| Variation_Orders | status=Draft |
| Purchase_Orders | status=Draft |
| Daily_Site_Reports | weather=Clear, issues="None reported" |
| Project_Documents | status=Draft |

#### Module 6A: Conditional Formatting
- `whenTextEqualTo` rules for every known status value
- Auto-colors new entries when status is selected from dropdown
- Colorblind-safe palette (green/blue/yellow/red/gray/purple)

#### Module 6B: Direct Cell Coloring
- One-time coloring for existing imported data
- Covers status, health_status, severity, result columns

#### Module 7: Auto-Resize (Separate)
- Not in setup flow (causes timeout)
- Available via menu: **Dubai ERP → 📏 Resize All Columns**
- Caps column width at 80–250px

### Color Palette (Colorblind-Safe)

| Status Type | Background | Text | Examples |
|-------------|-----------|------|----------|
| ✅ Success | `#d4edda` | `#155724` | Completed, Active, Approved, Pass, Closed |
| 🔵 In Progress | `#cce5ff` | `#004085` | In Progress, Certified, Under Review, Issued |
| ⚠️ Warning | `#fff3cd` | `#856404` | On Hold, Pending, Draft, Submitted, Medium |
| 🔴 Danger | `#f8d7da` | `#721c24` | Delayed, Rejected, Fail, Expired, High, Red |
| ⚪ Inactive | `#e2e3e5` | `#383d41` | Cancelled, Inactive, Terminated, Superseded |
| 🟣 Special | `#e8daef` | `#4a235a` | Awarded, Not Started |

### Menu Items

| Menu Item | Function |
|-----------|----------|
| ⚙️ Setup ERP System | `setupERPSystem()` — full setup |
| 📥 Import All Data | `importAllData()` — imports CSVs |
| 🔄 Rebuild Lookups | `rebuildLookups()` — refreshes _Lookups sheet |
| 🎨 Reapply Formatting | `reapplyFormatting()` — styling + dropdowns + colors |
| 🔄 Recolor Status Cells | `recolorStatus()` — re-colors status cells |
| 📏 Resize All Columns | `resizeColumnsMenu()` — auto-resize (separate to avoid timeout) |

---

## 🔧 Dropdown Values

All dropdown values verified against actual CSV data. Key additions made:

| Table | Values Added (were in CSV but missing from dropdowns) |
|-------|------------------------------------------------------|
| Employees | Commercial (dept), Inactive (status), Employment Visa, Golden Visa, Mission Visa |
| Contractors | MEP Works, Road Works |
| Contracts | Closed |
| Project_Phases | Excavation & Shoring, Finishes, Handover |
| Permits | Renewal Required |
| Inspections | Passed, Failed, Conditional |
| Safety_Incidents | Environmental Spill, Fire Incident, First Aid, Lost Time Injury, Medical Treatment, Near Miss, Property Damage |

---

## 🚀 Deployment Steps

### First Time Setup
1. Create a new Google Sheet
2. Open **Extensions → Apps Script**
3. Add `import_dubai_data.js` and `erp_automation.gs` as separate script files
4. Run `importAllData()` to pull CSV data from GitHub
5. Run `setupERPSystem()` to apply all formulas, dropdowns, and styling
6. Optionally run **📏 Resize All Columns** from the menu

### After Data Changes
- Run `setupERPSystem()` again to reapply formulas to new data
- Or use individual menu items for specific refreshes

---

## 🐛 Known Issues & Solutions

| Issue | Solution |
|-------|----------|
| Execution timeout (>6 min) | Auto-resize removed from setup; batch `setFormulas()` used |
| Red marks on dropdowns | `clearDataValidations()` before applying + `setAllowInvalid(true)` |
| Range out of bounds | `expandAllSheets()` runs first, inserts rows before any formulas |
| FK dropdown stores name | Changed to ID-only dropdown; name reference in `_Lookups` sheet |
| Missing dropdown values | 25+ values added after scanning all CSVs for mismatches |

---

## 📂 Git History (Key Commits)

1. **Initial data generator** — 17 tables with Dubai construction data
2. **Import engine** — CSV import from GitHub with memory optimization
3. **Visual redesign** — Colorblind-safe palette, dark headers, elegant styling
4. **Formula-based automation** — Complete rewrite from triggers to formulas
5. **Batch optimization** — setFormulas() arrays for speed
6. **Dropdown fixes** — 25+ missing values, no red marks
7. **Auto-fill defaults** — Status, tags, notes for all 17 tables

---

## 🔑 Configuration

```javascript
ADMIN: 'anandhu7833@gmail.com'
EXTRA_ROWS: 100
LOOKUP_SHEET: '_Lookups'
```

### Environment
- **Python**: 3.x with `pandas`, `faker` (in `.venv`)
- **Apps Script**: V8 runtime
- **Repository**: github.com/Anandhuvimalan/dubai-project-management-data
