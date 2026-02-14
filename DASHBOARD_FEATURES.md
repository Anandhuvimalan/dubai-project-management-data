# Dubai ERP — Operations Command Center v7.0
# Dashboard Features & Capabilities

---

## Overview

A fully automated operational dashboard built on Google Sheets that provides real-time visibility into **100 projects** across **17 integrated data modules**. The dashboard includes **12 live KPIs**, **23 analytical charts**, **15 detailed data tables**, and **smart filtering** — all computed in real-time from live project data.

---

## Smart Filtering System

| Feature | Description |
|---------|-------------|
| **Project Filter** | Select any specific project to instantly filter all KPIs, charts, and tables to that project's data |
| **Project Type Filter** | Filter by type (High-Rise, Villa, Infrastructure, Mixed-Use, etc.) to compare categories |
| **Status Filter** | Filter by status (In Progress, Delayed, Completed, On Hold) to focus on specific project stages |
| **One-Click Reset** | Checkbox that instantly clears all filters back to portfolio-wide view |
| **Auto-Refresh** | All filter changes automatically trigger a full dashboard rebuild — no manual refresh needed |

---

## Live KPI Cards (12 Metrics)

### Portfolio Health
| KPI | What It Tells You |
|-----|------------------|
| **Active Projects** | Number of currently in-progress projects |
| **Delayed Projects** | Number of projects flagged as delayed — immediate attention needed |
| **Total Contract Value** | Combined value of all contracts in the filtered portfolio (AED millions) |
| **Portfolio Completion %** | Weighted average completion across all projects — weighted by contract size |

### Financial Health
| KPI | What It Tells You |
|-----|------------------|
| **Budget Overrun %** | How much actual spend exceeds planned budget — turns red above 10% |
| **Pending Payments** | Total value of submitted but uncertified payment applications (AED millions) |
| **Approved Variation Orders** | Total value of approved scope changes (AED millions) |
| **Cost Performance Index** | Earned value efficiency — below 0.8 signals serious cost overrun |

### Operations Intelligence
| KPI | What It Tells You |
|-----|------------------|
| **Open Safety Incidents** | Unresolved safety incidents — turns red above 10 cases |
| **Inspection Pass Rate** | Percentage of inspections that passed — indicates quality control effectiveness |
| **Expiring Permits (<30 days)** | Permits that need renewal within the next month |
| **Rented Equipment Cost/Day** | Daily rental expenditure — highlights ongoing operational costs |

---

## Analytical Charts (23 Visualizations)

### Project Overview
| Chart | What It Shows |
|-------|-------------|
| **Project Status Distribution** | Pie chart showing the split of projects by status |
| **Project Completion Progress** | Top 15 projects ranked by completion percentage |
| **Cost per SQM by Project** | Construction cost benchmarking — contract value ÷ built-up area |
| **Client Portfolio Value** | Top 10 clients ranked by total project value |

### Contractor & Workforce
| Chart | What It Shows |
|-------|-------------|
| **Contractor Rating by Specialty** | Average performance rating across contractor specialties |
| **Top 10 Contractors by Contracts** | Most active contractors by contract count |
| **Manpower Trend (Direct + Subcon)** | Monthly direct and subcontractor workforce numbers |
| **Supplier Category Distribution** | Breakdown of suppliers by material/service category |

### Budget & Payments
| Chart | What It Shows |
|-------|-------------|
| **Budget Allocation by Phase** | How budget is distributed across project phases |
| **Budget vs Actual by Type** | Planned vs actual spending comparison by project type |
| **PO Monthly Spend Trend** | Purchase order spending pattern across months |
| **Payment Status Pipeline** | Distribution of payments by certification status |

### Equipment & Assets
| Chart | What It Shows |
|-------|-------------|
| **Fleet Ownership Split** | Rented vs owned equipment ratio |
| **Daily Rate: Rented vs Owned** | Cost comparison between rented and owned equipment by type |
| **Equipment Utilization** | In-use percentage for rented vs owned equipment |
| **Document Status Distribution** | Approval pipeline for project documents |

### Risk & Compliance
| Chart | What It Shows |
|-------|-------------|
| **VO Status Distribution** | Approved, pending, and rejected variation orders |
| **VO Risk by Project Type** | Which project types have the highest average VO values |
| **Incidents by Severity** | Safety incident breakdown by severity level |
| **Inspection Pass Rate by Type** | Quality pass rates compared across project types |
| **Duration: Planned vs Actual** | Schedule performance — are projects finishing on time? |
| **CPI by Project Type** | Cost efficiency comparison across project types |
| **Safety Incidents by Type** | Which project types generate the most safety incidents |

---

## Data Tables (15 Operational Panels)

### Schedule & Progress
| Table | What It Delivers |
|-------|-----------------|
| **New & Recent Projects** | Highlights projects started in the last 90 days with age badges (🟢 New / 🟡 Maturing / 🔴 Aging) |
| **Phase Progress Tracker** | Flags phases past their planned end date — shows days behind with CRITICAL/BEHIND alerts |

### Contractor Intelligence
| Table | What It Delivers |
|-------|-----------------|
| **Contractor Recommendations** | Composite scoring based on contract history, rating, and specialty — ranks A/B/C/D with color badges |

### Financial Control
| Table | What It Delivers |
|-------|-----------------|
| **Budget Health Monitor** | Per-project planned vs actual cost with variance — flags UNDER/ON TRACK/WARNING/OVER BUDGET |
| **Cash Flow & Payment Tracking** | Gross → Retained → Deducted → Net Certified breakdown per project |
| **Retention Money Tracker** | Retention held per project with DLP period and release eligibility status |

### Equipment & Procurement
| Table | What It Delivers |
|-------|-----------------|
| **Equipment Rental Alerts** | Rented equipment sorted by daily cost — flags HIGH/MEDIUM/LOW cost items |
| **PO Procurement Pipeline** | Pending purchase orders sorted by value with supplier and delivery status |

### Risk Management
| Table | What It Delivers |
|-------|-----------------|
| **VO Impact Tracker** | Variation orders ranked by financial impact — CRITICAL (>500K) / HIGH (>200K) / NORMAL |
| **LAD Risk Calculator** | Liquidated damages exposure: LAD/day × delay days = total financial risk per contract |

### HSE & Compliance
| Table | What It Delivers |
|-------|-----------------|
| **Safety Incidents** | Recent incidents sorted by date with severity and resolution status |
| **Inspection Defects** | Failed inspections with defect count and corrective actions needed |
| **Permits Expiry Tracker** | Permits sorted by urgency — 🔴 EXPIRED / 🟡 Expiring Soon / 🟢 Valid |
| **Document Submission Status** | Pending and rejected documents requiring attention |

### Executive Summary
| Table | What It Delivers |
|-------|-----------------|
| **Decision Matrix** | One-row-per-type summary combining project count, completion %, CPI, VO rate, inspection pass rate, and overall risk level |

---

## Data Integration

The dashboard pulls from all **17 data modules** in real-time:

| Module | Key Data Points |
|--------|----------------|
| Projects | Status, type, completion, contract value, GFA, dates |
| Contracts | FIDIC type, contract value, LAD rates, retention %, completion dates |
| Contractors | Company name, specialty, rating, prequalification |
| Work Packages | Planned cost, actual cost per package |
| Payment Applications | Gross amount, certified amount, deductions, status |
| Variation Orders | VO value, approval status, impact classification |
| Safety Incidents | Severity, status, incident type, project linkage |
| Inspections | Result, inspection type, defect count, corrective actions |
| Permits & Approvals | Permit type, expiry date, issuing authority |
| Equipment | Type, ownership, daily rate, utilization status |
| Project Phases | Phase name, planned dates, actual dates, budget |
| Purchase Orders | Supplier, amount, delivery status |
| Project Documents | Document type, submission status |
| Clients | Client name, type, portfolio value |
| Daily Site Reports | Manpower counts, weather, work description |
| Suppliers | Category, company name, materials supplied |
| Employees | Department, designation, visa status |

---

## Key Technical Highlights

- **Fully automated** — one-click deployment via Apps Script menu
- **Filter-reactive** — every KPI and table responds to filter changes in real-time
- **Color-coded alerts** — red/yellow/green thresholds on all KPIs and table rows for instant visual assessment
- **Cross-module analysis** — tables join data from multiple sheets (e.g., LAD calculator joins Contracts + Projects + payment data)
- **No manual data entry** — all values computed from existing project data

---

*Operations Command Center v7.0 — Built for Dubai Project Management*
