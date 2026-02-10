# Dubai Project Management Data Overview

This dataset models a construction company in Dubai that runs multiple projects end‑to‑end. It looks like the data you would find in a real ERP: commercial setup, contracts, procurement, execution, safety, quality, payments, and documents. The goal is to track every project from client onboarding to completion, including budget, progress, risks, and compliance.

Think of it as a full lifecycle picture:
- A client requests a project.
- The project is created with dates, value, and a manager.
- Contracts and subcontracts are signed.
- Work is broken into phases and work packages.
- Teams report daily activity and equipment usage.
- Inspections and safety events are logged.
- Permits and documents are tracked for compliance.
- Payments and variations are managed as work progresses.

Below is a simple, detailed explanation for each CSV and how it can be used in an ERP dashboard.

---

## 1) `clients.csv`
**What it is:**
A list of project owners (clients). It contains legal details, contact persons, and payment terms.

**Why it matters:**
This is your customer master data. It drives credit checks, billing terms, and contract setup.

**How to use it in an ERP dashboard:**
- Client list with key contacts and payment terms.
- Credit exposure vs. credit limit.
- Active projects per client.

---

## 2) `projects.csv`
**What it is:**
The master list of projects, including type, location, value, dates, status, size, and assigned project manager.

**Why it matters:**
This is the core table that connects everything else. Almost all other CSVs link to `project_id`.

**How to use it in an ERP dashboard:**
- Project portfolio view (status, completion %, value).
- Project timeline (start/end, delays).
- Project KPIs (cost, safety, quality, cashflow).

---

## 3) `employees.csv`
**What it is:**
Company employees with department, role, visa status, and basic HR data.

**Why it matters:**
Employees are referenced across daily reports, inspections, and documents.

**How to use it in an ERP dashboard:**
- Staff directory by department/role.
- Project manager and inspector accountability.
- Resource planning by availability/roles.

---

## 4) `contractors.csv`
**What it is:**
A list of subcontractors and specialist contractors, including qualifications and ratings.

**Why it matters:**
Used when creating subcontracts and tracking subcontractor performance.

**How to use it in an ERP dashboard:**
- Contractor performance scorecards (rating, prequalified status).
- Active subcontractors by project.
- Compliance checks (licenses, registrations).

---

## 5) `contracts.csv`
**What it is:**
All contracts for projects, including main contract and subcontracts.

**Why it matters:**
Defines project value, timelines, and contractual terms (retention, LAD, bonds, etc.).

**How to use it in an ERP dashboard:**
- Contract value vs. project budget.
- Active vs. closed contracts.
- Risk summary (LADs, delays, defect liability periods).

---

## 6) `project_phases.csv`
**What it is:**
Project timeline broken into phases, with planned vs actual dates and progress.

**Why it matters:**
This is how you track schedule performance phase by phase.

**How to use it in an ERP dashboard:**
- Phase Gantt view (planned vs actual).
- Phase progress and delay tracking.
- Phase budget consumption.

---

## 7) `work_packages.csv`
**What it is:**
Detailed work items (like a bill of quantities). Each line has planned quantity, rate, and executed quantity.

**Why it matters:**
This is the base for measuring progress and cost at item level.

**How to use it in an ERP dashboard:**
- Progress tracking by item (planned vs executed).
- Cost tracking (budget vs executed amount).
- Top cost items or biggest remaining work.

---

## 8) `daily_site_reports.csv`
**What it is:**
Daily site activity logs: manpower, equipment count, weather, and work description.

**Why it matters:**
This is the operational heartbeat of the site. It supports productivity analysis and delay claims.

**How to use it in an ERP dashboard:**
- Daily manpower and equipment trends.
- Productivity tracking by day/project.
- Weather impact analysis.

---

## 9) `inspections.csv`
**What it is:**
Quality control inspections with results, defects, and corrective actions.

**Why it matters:**
Quality issues can cause delays and rework costs. Inspection data shows compliance health.

**How to use it in an ERP dashboard:**
- Pass/fail rate by project or inspection type.
- Open defects requiring corrective action.
- Reinspection trends and turnaround time.

---

## 10) `safety_incidents.csv`
**What it is:**
Health and Safety incident logs including severity, causes, and actions.

**Why it matters:**
Safety performance is critical and often reported to leadership and regulators.

**How to use it in an ERP dashboard:**
- Incident frequency and severity trends.
- Lost time injury (LTI) days tracking.
- Root cause analysis and corrective action status.

---

## 11) `permits_approvals.csv`
**What it is:**
Regulatory permits/approvals (DDA, RTA, Civil Defense, etc.) with dates and fees.

**Why it matters:**
Work cannot proceed without valid approvals. Expired permits cause delays.

**How to use it in an ERP dashboard:**
- Permit status tracker (approved/pending/expired).
- Upcoming expiry alerts.
- Permit cost summary per project.

---

## 12) `project_documents.csv`
**What it is:**
Document control register: drawings, reports, minutes, etc.

**Why it matters:**
A strong document trail is required for approvals, QA/QC, and claims.

**How to use it in an ERP dashboard:**
- Document status (for review/approved/superseded).
- Latest revision tracking by discipline.
- Document volume and turnaround time.

---

## 13) `purchase_orders.csv`
**What it is:**
Procurement orders for materials or services.

**Why it matters:**
Procurement cost and delivery timing affect budget and schedule.

**How to use it in an ERP dashboard:**
- PO value by project and supplier.
- Delivery status (issued, delivered, invoiced, paid).
- Procurement lead time monitoring.

---

## 14) `suppliers.csv`
**What it is:**
Supplier master list with categories, lead time, and approval status.

**Why it matters:**
Used for vendor selection, compliance checks, and purchase orders.

**How to use it in an ERP dashboard:**
- Approved vs unapproved suppliers.
- Supplier lead time by category.
- Supplier distribution by material type.

---

## 15) `payment_applications.csv`
**What it is:**
Interim Payment Certificates (IPC) showing progress billing and payment status.

**Why it matters:**
This tracks cash inflow and client payments during the project lifecycle.

**How to use it in an ERP dashboard:**
- Certified vs paid amounts.
- Outstanding receivables per project.
- Cashflow trend by month.

---

## 16) `variation_orders.csv`
**What it is:**
Change orders that adjust contract scope, cost, or time.

**Why it matters:**
Variations are common and must be tracked to protect revenue and schedule.

**How to use it in an ERP dashboard:**
- Total approved variation value per project.
- Pending vs rejected variation count.
- Time impact days summary.

---

## 17) `equipment.csv`
**What it is:**
Equipment inventory with ownership, rental rate, availability, and service date.

**Why it matters:**
Equipment costs and availability affect daily productivity.

**How to use it in an ERP dashboard:**
- Equipment utilization by project.
- Owned vs rented cost comparison.
- Maintenance due alerts.

---

# Summary: What the full data is about
This dataset gives a complete view of a construction contractor’s operations in Dubai. It includes:
- **Commercial setup** (clients, contracts, variations).
- **Project execution** (phases, work packages, daily reports).
- **Procurement & resources** (suppliers, purchase orders, equipment).
- **Quality & safety** (inspections, safety incidents).
- **Compliance & documentation** (permits, documents).
- **Cashflow** (payment applications).

For an ERP dashboard, you can build modules for Portfolio, Contracts, Procurement, Execution, Safety, Quality, Compliance, and Finance — all tied together by `project_id`.
