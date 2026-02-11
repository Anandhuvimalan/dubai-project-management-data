# Dubai Construction ERP Dashboard - Technical Design Specification v2.0

## 1. Executive Summary
**Objective:** To provide a comprehensive, real-time view of construction project health, financial performance, and operational safety for executive decision-making.
**Target Audience:** Project Managers, Operations Directors, C-Level Executives.
**Core Value:** Unification of 17 disparate data sources into a single "Command Center" for the Dubai construction portfolio.

---

## 2. Visual Layout & Wireframe

The dashboard follows a strict **104-column grid** layout to ensure density without clutter. Below is the visual structure:

<div style="border: 2px solid #333; padding: 20px; background: #f0f4f8; border-radius: 8px;">
  <!-- Header -->
  <div style="background: #fff; border: 1px solid #ccc; padding: 15px; margin-bottom: 20px; text-align: center; font-weight: bold;">
    HEADER AREA: Logo | Title | Last Updated | [ GLOBAL PROJECT FILTER ]
  </div>

  <!-- KPI Area -->
  <div style="display: flex; gap: 10px; margin-bottom: 20px;">
    <div style="flex: 1; background: #fff; padding: 20px; border-left: 5px solid #6366F1; text-align: center;">FINANCIAL KPI</div>
    <div style="flex: 1; background: #fff; padding: 20px; border-left: 5px solid #10B981; text-align: center;">PAYMENT KPI</div>
    <div style="flex: 1; background: #fff; padding: 20px; border-left: 5px solid #F59E0B; text-align: center;">VO KPI</div>
    <div style="flex: 1; background: #fff; padding: 20px; border-left: 5px solid #3B82F6; text-align: center;">PO KPI</div>
  </div>

  <div style="display: flex; gap: 10px; margin-bottom: 20px;">
    <div style="flex: 1; background: #fff; padding: 20px; border-left: 5px solid #14B8A6; text-align: center;">ACTIVE PROJECTS</div>
    <div style="flex: 1; background: #fff; padding: 20px; border-left: 5px solid #8B5CF6; text-align: center;">QUALITY PASS %</div>
    <div style="flex: 1; background: #fff; padding: 20px; border-left: 5px solid #EF4444; text-align: center;">SAFETY INCIDENTS</div>
    <div style="flex: 1; background: #fff; padding: 20px; border-left: 5px solid #F97316; text-align: center;">RESOURCE UTIL %</div>
  </div>

  <!-- Charts Area - 2 Columns -->
  <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px;">
    <div style="background: #fff; height: 200px; border: 1px dashed #999; display: flex; align-items: center; justify-content: center;">Financial Charts (Budget & Cashflow)</div>
    <div style="background: #fff; height: 200px; border: 1px dashed #999; display: flex; align-items: center; justify-content: center;">Operations Charts (Progress & Status)</div>
    <div style="background: #fff; height: 200px; border: 1px dashed #999; display: flex; align-items: center; justify-content: center;">Procurement Charts (POs & Supply Chain)</div>
    <div style="background: #fff; height: 200px; border: 1px dashed #999; display: flex; align-items: center; justify-content: center;">HSE & Quality Charts</div>
  </div>

  <!-- Footer/Table Area -->
  <div style="background: #fff; border: 1px solid #ccc; padding: 15px; text-align: center;">
    PROJECT HEALTH SUMMARY TABLE (Top 10 Projects Detailed View)
  </div>
</div>

---

## 3. Detailed Component Specifications

### 3.1 Global Navigation & Control
*   **Project Slicer (L2)**: The "Engine" of the dashboard. Selecting a project here instantly filters:
    *   *All* KPI cards.
    *   *All* 16 Charts.
    *   *The* Health Table remains static to provide context.

### 3.2 Advanced Metrics (The "New Essentials")

To meet modern project management standards, the following advanced metrics are integrated:

#### A. Contractor Performance Module
*   **Performance Scorecard**: Weighted average of `rating` from `contractors.csv`.
*   **Defect Density**: Number of defects found (from `inspections.csv`) divided by `work_pages` completed.
*   **Metric**: *"Top 5 Performing Contractors"* vs *"Contractors At Risk"*.

#### B. Supply Chain Intelligence
*   **Lead Time Variance**: Planned delivery vs Actual delivery (from `purchase_orders.csv`).
*   **Vendor Reliance**: % of Total PO Value awarded to top 3 suppliers (Risk assessment).
*   **Visual**: Scatter plot of *Vendor Value* vs *Lead Time*.

#### C. Equipment & Asset Efficiency
*   **Idle Cost Analysis**: Value of equipment with status `Available` (Not earning revenue).
*   **Maintenance Alert**: Count of equipment where `last_service_date` > 180 days.
*   **Visual**: Gauge Chart for *Fleet Utilization %*.

#### D. Project Stages & Forecasting
*   **Phase Slippage**: Comparison of `actual_end` vs `planned_end` from `project_phases.csv`.
*   **Forecasted Completion**: Linear extrapolation of progress based on current burn rate.
*   **Visual**: Grantt Chart Overlay (Planned vs Actual).

---

## 4. Visual Identity & User Experience

### Interactive Elements
1.  **Hover Effects**: Chart tooltips display precise values (e.g., "PO #1234: 50,000 AED").
2.  **Drill-Down**: Clicking a slice in the "Project Status" pie chart filters the Health Table below.
3.  **Conditional Formatting**:
    *   **Progress Bars**: Visual bars inside the table for "Completion %".
    *   **Traffic Lights**: Status indicators (Green/Amber/Red) for quick health assessment.

### Color Semantics details
*   **Revenue/Value**: Indigo/Purple (Royal, confident).
*   **Operations/Activity**: Teal/Cyan (Vibrant, active).
*   **Risk/Issues**: Red/Orange (Alarming, draws attention).
*   **Safety/Environment**: Emerald/Green (Natural, safe).

---

## 5. Implementation Data Model

### Data Sources & Transformations
| Metric Group | Primary CSV Source | Transformation Logic |
| :--- | :--- | :--- |
| **Financials** | `contracts.csv`, `payment_applications.csv` | `SUMIF` ProjectID, `SUMPRODUCT` (Certified + Paid). |
| **Schedule** | `project_phases.csv` | `DATEDIF` (Planned Start, Actual Start). |
| **Quality** | `inspections.csv` | `COUNTIF` (Result = "Passed") / `COUNTA`. |
| **Safety** | `safety_incidents.csv` | Incident Rate = (Incidents / Manpower) * 1000. |
| **Supply Chain** | `suppliers.csv`, `purchase_orders.csv` | Lead Time = Delivery Date - Issue Date. |

---

## 6. Risk & Governance (New Section)

### Risk Matrix (Derived)
Risks are calculated dynamically based on project parameters:
*   **Schedule Risk**: If `Actual Progress` < `Planned Progress` by > 10%.
*   **Cost Risk**: If `Approved VO Value` > 15% of `Original Contract Value`.
*   **Quality Risk**: If `Inspection Pass Rate` < 85%.

### Quality Assurance (NCRs)
*   Visual tracking of Non-Conformance Reports (NCRs).
*   Tracking of "Corrective Actions" (from `inspections.csv` > `corrective_action` column).
