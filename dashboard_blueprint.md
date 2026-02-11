# Dubai Construction ERP Dashboard - Functional Specification

## 1. Executive Summary
**Objective:** To implement a centralized ERP dashboard for the Dubai construction portfolio, integrating 17 existing data modules into a single analytical view.

**Core Data Sources:**
*   **Financials:** Contracts, Payment Applications, Purchase Orders, Variation Orders.
*   **Operations:** Projects, Daily Reports, Equipment, Work Packages.
*   **Quality & Safety:** Inspections, Safety Incidents.
*   **Stakeholders:** Contractors, Suppliers, Employees, Clients.

---

## 2. Implementation Roadmap

To build this ERP dashboard, the following core activities are required:

### Phase 1: Data Architecture
1.  **Data Integration:** Establish specific query logic to aggregated CSV data from all 17 sources.
2.  **Global Filtering:** Implement a dynamic `project_id` slicer that cascades across all metrics.
3.  **Relational Mapping:** Link `project_id` across disparate datasets (e.g., matching Safety Incidents to Project Names).

### Phase 2: Logic & Calculation
1.  **Financial Aggregation:** Develop formulas to sum "Certified" and "Paid" amounts distinct from "Submitted".
2.  **Risk Modeling:** Create logic to flag projects where `Actual Progress` < `Planned Progress`.
3.  **Performance Scoring:** Calculate weighted averages for Contractor Ratings and Supplier Lead Times.

### Phase 3: Visualization & Reporting
1.  **Chart Construction:** Build the 16 specific visualizations defined below.
2.  **KPI Dashboarding:** Implement the 8 headline metrics.
3.  **Tabular Reporting:** Construct the detailed "Project Health Matrix".

---

## 3. Required Dashboard Components

The dashboard must implement the following specific visualizations and metrics.

### A. Headline KPIs (The "Pulse")
These 8 metrics must be calculated and displayed prominently:
1.  **Total Contract Value (AED):** Sum of active contracts.
2.  **Certified Payments (AED):** Validated work done.
3.  **Approved VO Value (AED):** Cost of scope changes.
4.  **Total PO Value (AED):** Procurement spend.
5.  **Active Projects Count:** Operational workload.
6.  **Inspection Pass Rate (%):** Quality benchmark.
7.  **Open Safety Incidents:** HSE risk indicator.
8.  **Equipment Utilization (%):** Asset efficiency.

### B. Financial Performance Charts
These charts are required to analyze budget and cash flow:
1.  **Budget Utilization (Donut Chart):**
    *   *Data:* Certified Amount + Paid Amount vs. Remaining Budget.
2.  **Payment Cashflow Status (Column Chart):**
    *   *Data:* Total value grouped by status: Certified, Paid, Submitted, Rejected.
3.  **Contract Value by Type (Bar Chart):**
    *   *Data:* Total contract value grouped by discipline (Civil, MEP, etc.).
4.  **Variation Order Trend (Line Chart):**
    *   *Data:* Cumulative VO value plotted monthly over the project lifecycle.

### C. Operational & Project Charts
These charts track schedule and progress:
1.  **Project Status Distribution (Bar Chart):**
    *   *Data:* Count of projects by status (In Progress, On Hold, Completed).
2.  **Project Type Mix (Pie Chart):**
    *   *Data:* Portfolio breakdown by sector (Residential, Infra, Commercial).
3.  **Manpower Trend Analysis (Stacked Area Chart):**
    *   *Data:* Monthly average headcount split by Direct Labor vs. Subcontractors.
4.  **Inspection Result Analysis (Column Chart):**
    *   *Data:* Count of inspections by result (Pass, Fail, Partial).

### D. Supply Chain & Procurement Charts
These charts monitor vendor performance:
1.  **PO Status Breakdown (Column Chart):**
    *   *Data:* Purchase Order value split by status (Issued, Closed, Draft).
2.  **VO Reason Analysis (Bar Chart):**
    *   *Data:* Variation costs grouped by root cause (Design Change, Site Condition, etc.).
3.  **Permit Status Overview (Donut Chart):**
    *   *Data:* Count of active vs. expired permits.
4.  **Supplier Lead Time Analysis (Scatter/Bar):**
    *   *Data:* Average days (Delivery Date - Issue Date) by Supplier Category.

### E. HSE, Quality & Resource Charts
These charts track safety and assets:
1.  **Safety Incidents by Type (Bar Chart):**
    *   *Data:* Incident counts by category (Near Miss, LTI, First Aid).
2.  **Safety Severity Distribution (Donut Chart):**
    *   *Data:* Incidents grouped by severity level (High, Medium, Low).
3.  **Equipment Fleet Status (Pie Chart):**
    *   *Data:* Equipment count by status (Available, In Use, Maintenance).
4.  **Contractor Performance (Bar Chart):**
    *   *Data:* Average performance rating for top 10 active contractors.

---

## 4. Project Health Matrix (Detailed Table)

A comprehensive data table is required to list the Top 10 Projects with the following columns:
*   **Project Name & ID**
*   **Current Status** (with R/A/G indicators)
*   **Completion %** (Visual progress bar)
*   **Financials:** Contract Value vs. Certified Amount
*   **Risk Indicators:** Open VOs count, Open Safety Incidents count

---

## 5. Functional Data Requirements

To support the above, the system must perform these specific data operations:
*   **Time-Intelligence:** Ability to group `daily_site_reports` and `payment_applications` by Month/Year.
*   **Cross-Referencing:** Validating `contractor_id` in the `projects` table against the `contractors` master list.
*   **Threshold Monitoring:** Automatically flagging "High Risk" if Safety Incidents > 0 or Schedule Slippage > 10%.
