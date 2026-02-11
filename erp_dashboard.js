/**
 * ============================================================================
 *  DUBAI CONSTRUCTION ERP DASHBOARD - Google Apps Script
 * ============================================================================
 *  A comprehensive project management dashboard for Dubai construction
 *  operations. Reads from 17 CSV-sourced Google Sheets tabs and renders
 *  an interactive, filterable dashboard with KPIs, charts, and health table.
 *
 *  Sheet tabs: Projects, Contracts, Payment_Applications, Safety_Incidents,
 *  Daily_Site_Reports, Inspections, Purchase_Orders, Variation_Orders,
 *  Permits_Approvals, Project_Documents, Project_Phases, Work_Packages,
 *  Equipment, Employees, Clients, Contractors, Suppliers
 *
 *  Author : ERP Dashboard Generator
 *  Version: 3.0
 *  Date   : 2026-02-11
 * ============================================================================
 */

// ---------------------------------------------------------------------------
//  THEME CONSTANTS
// ---------------------------------------------------------------------------
var THEME = {
  bg:            '#F8FAFC',
  cardBg:        '#FFFFFF',
  sidebarBg:     '#0F172A',
  sidebarText:   '#94A3B8',
  sidebarActive: '#818CF8',
  headerText:    '#0F172A',
  subText:       '#64748B',
  border:        '#E2E8F0',
  accent:        '#6366F1',
  success:       '#10B981',
  warning:       '#F59E0B',
  danger:        '#EF4444',
  info:          '#3B82F6',
  teal:          '#14B8A6',
  purple:        '#8B5CF6',
  pink:          '#EC4899',
  orange:        '#F97316',
  palette: [
    '#6366F1', '#8B5CF6', '#EC4899', '#F43F5E', '#F97316',
    '#EAB308', '#22C55E', '#14B8A6', '#06B6D4', '#3B82F6',
    '#A78BFA', '#FB923C'
  ],
  font: 'Inter, Roboto, sans-serif'
};

// Dashboard sheet name
var DASH_NAME = 'ERP Dashboard';

// ---------------------------------------------------------------------------
//  MENU
// ---------------------------------------------------------------------------
/**
 * Adds the ERP Dashboard menu to the spreadsheet UI on open.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ERP Dashboard')
    .addItem('Create / Refresh Dashboard', 'createERPDashboard')
    .addToUi();
}

// ---------------------------------------------------------------------------
//  MAIN ENTRY POINT
// ---------------------------------------------------------------------------
/**
 * Main function: deletes any existing dashboard sheet, creates a fresh one,
 * and populates it with layout, KPIs, data engine, charts, and health table.
 */
function createERPDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Delete old dashboard if it exists
  var old = ss.getSheetByName(DASH_NAME);
  if (old) {
    ss.deleteSheet(old);
    SpreadsheetApp.flush();
  }

  // Create new dashboard sheet at position 0
  var sh = ss.insertSheet(DASH_NAME, 0);
  SpreadsheetApp.flush();

  // Build all components in sequence
  setupLayout_(sh);
  buildSidebar_(sh);
  buildHeader_(sh);
  setupProjectFilter_(sh);
  buildDataEngine_(sh);
  buildKPIRow1_(sh);
  buildKPIRow2_(sh);
  buildSections_(sh);
  buildAllCharts_(sh);
  buildHealthTable_(sh);

  // Hide helper columns Z (col 26) through CZ (col 104)
  sh.hideColumns(26, 79);

  // Final flush and set active cell
  SpreadsheetApp.flush();
  sh.setActiveSelection('C2');

  ui.alert(
    'ERP Dashboard',
    'Dashboard created successfully!\n\nUse the filter at L2 to select a specific project.',
    ui.ButtonSet.OK
  );
}

// ---------------------------------------------------------------------------
//  LAYOUT
// ---------------------------------------------------------------------------
/**
 * Sets column widths, row heights, freezes, and background colours for the
 * entire dashboard grid.
 *
 * Column plan:
 *   A (72px) = Sidebar icons
 *   B (25px) = Padding
 *   C-E, G-I, K-M, O-Q (115px each) = Content columns
 *   F, J, N (20px) = Gap columns
 *   R-Y (40px) = Extra space
 *   Z-CZ (100px) = Hidden helper columns
 */
function setupLayout_(sh) {
  // Ensure we have enough columns and rows to work with
  if (sh.getMaxColumns() < 104) {
    sh.insertColumnsAfter(sh.getMaxColumns(), 104 - sh.getMaxColumns());
  }
  if (sh.getMaxRows() < 160) {
    sh.insertRowsAfter(sh.getMaxRows(), 160 - sh.getMaxRows());
  }

  // -- Column widths --
  sh.setColumnWidth(1, 72);   // A: Sidebar
  sh.setColumnWidth(2, 25);   // B: Padding

  // Content columns: 115px each
  var contentCols = [3, 4, 5, 7, 8, 9, 11, 12, 13, 15, 16, 17];
  for (var i = 0; i < contentCols.length; i++) {
    sh.setColumnWidth(contentCols[i], 115);
  }
  // Gap columns: 20px
  sh.setColumnWidth(6, 20);
  sh.setColumnWidth(10, 20);
  sh.setColumnWidth(14, 20);

  // Extra columns R-Y
  for (var c = 18; c <= 25; c++) {
    sh.setColumnWidth(c, 40);
  }

  // Helper columns Z-CZ
  for (var c = 26; c <= 104; c++) {
    sh.setColumnWidth(c, 100);
  }

  // -- Row heights --
  // Default height for all rows
  for (var r = 1; r <= 160; r++) {
    sh.setRowHeight(r, 22);
  }
  // Special rows
  sh.setRowHeight(1, 8);    // Top padding
  sh.setRowHeight(2, 40);   // Header / title
  sh.setRowHeight(3, 24);   // Subtitle
  sh.setRowHeight(4, 12);   // Spacer below header
  sh.setRowHeight(12, 12);  // Spacer below KPIs

  // -- Backgrounds --
  // Main content background
  sh.getRange(1, 1, 160, 17).setBackground(THEME.bg);

  // Sidebar background (column A, full height)
  sh.getRange(1, 1, 160, 1).setBackground(THEME.sidebarBg);

  // Padding column B
  sh.getRange(1, 2, 160, 1).setBackground(THEME.bg);

  // Gap columns F, J, N match main background
  sh.getRange(1, 6, 160, 1).setBackground(THEME.bg);
  sh.getRange(1, 10, 160, 1).setBackground(THEME.bg);
  sh.getRange(1, 14, 160, 1).setBackground(THEME.bg);

  // Freeze sidebar + header
  sh.setFrozenColumns(1);
  sh.setFrozenRows(3);
}

// ---------------------------------------------------------------------------
//  SIDEBAR
// ---------------------------------------------------------------------------
/**
 * Renders sidebar navigation icons in column A using Unicode glyphs.
 * The first item is highlighted as "active".
 */
function buildSidebar_(sh) {
  var items = [
    { row: 2,   icon: '\u2302', active: true  },  // Home
    { row: 5,   icon: '\u2261', active: false },  // KPIs
    { row: 13,  icon: '\u0024', active: false },  // Finance
    { row: 43,  icon: '\u2692', active: false },  // Projects
    { row: 74,  icon: '\u2606', active: false },  // Procurement
    { row: 105, icon: '\u2665', active: false },  // HSE
    { row: 136, icon: '\u2611', active: false }   // Health
  ];

  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    var cell = sh.getRange(it.row, 1);
    cell.setValue(it.icon)
        .setFontFamily(THEME.font)
        .setFontSize(16)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setFontColor(it.active ? THEME.sidebarActive : THEME.sidebarText);

    if (it.active) {
      cell.setBackground('#1E293B').setFontWeight('bold');
    }
  }
}

// ---------------------------------------------------------------------------
//  HEADER
// ---------------------------------------------------------------------------
/**
 * Builds the dashboard header: title, subtitle, timestamp, and filter label.
 */
function buildHeader_(sh) {
  // Title (C2:I2)
  var titleRange = sh.getRange('C2:I2');
  titleRange.merge()
    .setValue('Dubai Construction ERP Dashboard')
    .setFontFamily(THEME.font)
    .setFontSize(18)
    .setFontWeight('bold')
    .setFontColor(THEME.headerText)
    .setVerticalAlignment('middle')
    .setBackground(THEME.bg);

  // Subtitle (C3:I3)
  var subRange = sh.getRange('C3:I3');
  subRange.merge()
    .setValue('Real-time project management overview across 17 data modules')
    .setFontFamily(THEME.font)
    .setFontSize(10)
    .setFontColor(THEME.subText)
    .setVerticalAlignment('middle')
    .setBackground(THEME.bg);

  // Last updated (O2:Q2)
  var updRange = sh.getRange('O2:Q2');
  updRange.merge()
    .setFormula('="Last updated: "&TEXT(NOW(),"DD-MMM-YYYY HH:MM")')
    .setFontFamily(THEME.font)
    .setFontSize(9)
    .setFontColor(THEME.subText)
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle')
    .setBackground(THEME.bg);

  // Filter label (K2)
  sh.getRange('K2')
    .setValue('Project Filter:')
    .setFontFamily(THEME.font)
    .setFontSize(10)
    .setFontWeight('bold')
    .setFontColor(THEME.headerText)
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle')
    .setBackground(THEME.bg);
}

// ---------------------------------------------------------------------------
//  PROJECT FILTER
// ---------------------------------------------------------------------------
/**
 * Creates the project filter dropdown at L2:N2, populates the project list
 * in hidden column CZ (col 104), and sets up the two dynamic WHERE-clause
 * helpers in Z2 and Z3.
 */
function setupProjectFilter_(sh) {
  // Merge and style the dropdown cell
  var filterCell = sh.getRange('L2:N2');
  filterCell.merge()
    .setValue('All Projects')
    .setFontFamily(THEME.font)
    .setFontSize(11)
    .setFontWeight('bold')
    .setFontColor(THEME.accent)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(THEME.cardBg)
    .setBorder(true, true, true, true, false, false,
               THEME.accent, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Get unique project IDs and write to column CZ (104)
  var projectIds = getUniqueProjectIds_();
  sh.getRange(1, 104).setValue('All Projects');
  for (var i = 0; i < projectIds.length; i++) {
    sh.getRange(i + 2, 104).setValue(projectIds[i]);
  }

  // Data validation dropdown on L2
  var listRange = sh.getRange(1, 104, projectIds.length + 1, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(listRange, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange('L2').setDataValidation(rule);

  // Z2: dynamic WHERE clause for sheets where project_id is in column B
  sh.getRange('Z2')
    .setFormula('=IF($L$2="All Projects","", " AND B = \'"&$L$2&"\'")')
    .setFontSize(8).setFontColor('#999999');

  // Z3: dynamic WHERE clause for Projects sheet where project_id is in column A
  sh.getRange('Z3')
    .setFormula('=IF($L$2="All Projects","", " AND A = \'"&$L$2&"\'")')
    .setFontSize(8).setFontColor('#999999');
}

/**
 * Reads unique project IDs from the Projects sheet column A.
 * @return {string[]} Sorted array of project IDs.
 */
function getUniqueProjectIds_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var projSheet = ss.getSheetByName('Projects');
  if (!projSheet) return [];

  var lastRow = projSheet.getLastRow();
  if (lastRow < 2) return [];

  var data = projSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var seen = {};
  for (var i = 0; i < data.length; i++) {
    var val = String(data[i][0]).trim();
    if (val !== '' && val !== 'undefined') {
      seen[val] = true;
    }
  }
  return Object.keys(seen).sort();
}

// ---------------------------------------------------------------------------
//  KPI CARD BUILDER (reusable)
// ---------------------------------------------------------------------------
/**
 * Draws a single KPI card occupying 3 rows x 3 columns.
 *
 * @param {Sheet}  sh         Dashboard sheet
 * @param {number} col        Starting column (e.g. 3 for column C)
 * @param {number} row        Starting row
 * @param {string} title      Card title
 * @param {string} formula    Formula string for the main value
 * @param {string} fmt        Number format (e.g. '#,##0.0')
 * @param {string} sparkRange Range for sparkline data, or null
 * @param {string} color      Accent colour for top border
 */
function drawCard_(sh, col, row, title, formula, fmt, sparkRange, color) {
  // -- Title row (row, col -> col+2) --
  var titleCell = sh.getRange(row, col, 1, 3);
  titleCell.merge()
    .setValue(title)
    .setFontFamily(THEME.font)
    .setFontSize(9)
    .setFontColor(THEME.subText)
    .setFontWeight('normal')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('bottom')
    .setBackground(THEME.cardBg);

  // Card border on title
  titleCell.setBorder(true, true, null, true, false, false,
                      THEME.border, SpreadsheetApp.BorderStyle.SOLID);
  // Coloured top accent border (overrides top)
  sh.getRange(row, col, 1, 3)
    .setBorder(true, null, null, null, null, null,
               color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // -- Value row (row+1) --
  var valueCell = sh.getRange(row + 1, col, 1, 3);
  valueCell.merge()
    .setFormula(formula)
    .setNumberFormat(fmt)
    .setFontFamily(THEME.font)
    .setFontSize(20)
    .setFontWeight('bold')
    .setFontColor(THEME.headerText)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(THEME.cardBg);
  valueCell.setBorder(null, true, null, true, false, false,
                      THEME.border, SpreadsheetApp.BorderStyle.SOLID);

  // -- Sparkline / bottom row (row+2) --
  var bottomCell = sh.getRange(row + 2, col, 1, 3);
  bottomCell.merge()
    .setBackground(THEME.cardBg)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('top');
  bottomCell.setBorder(null, true, true, true, false, false,
                       THEME.border, SpreadsheetApp.BorderStyle.SOLID);

  if (sparkRange) {
    bottomCell.setFormula(
      '=SPARKLINE(' + sparkRange + ', {"charttype","line";' +
      '"color","' + color + '";' +
      '"linewidth",2})'
    );
  } else {
    bottomCell.setValue(' ')
      .setFontSize(8)
      .setFontColor(THEME.subText);
  }
}

// ---------------------------------------------------------------------------
//  KPI ROW 1 (rows 5-7) - Financial metrics
// ---------------------------------------------------------------------------
/**
 * First row of four KPI cards: contract value, certified payments,
 * approved VOs, and total PO value (all in millions AED).
 */
function buildKPIRow1_(sh) {
  // Card 1: Total Contract Value (M AED)
  drawCard_(sh, 3, 5,
    'Total Contract Value (M AED)',
    '=IFERROR(IF($L$2="All Projects",' +
      'SUMPRODUCT((Contracts!O1:O664<>"")*(Contracts!F1:F664))/1000000,' +
      'SUMPRODUCT((Contracts!B1:B664=$L$2)*(Contracts!F1:F664))/1000000), 0)',
    '#,##0.0',
    'AA2:AA13',
    THEME.accent
  );

  // Card 2: Certified Payments (M AED)
  drawCard_(sh, 7, 5,
    'Certified Payments (M AED)',
    '=IFERROR(IF($L$2="All Projects",' +
      'SUMPRODUCT((Payment_Applications!L1:L823="Certified")*(Payment_Applications!K1:K823)' +
      '+(Payment_Applications!L1:L823="Paid")*(Payment_Applications!K1:K823))/1000000,' +
      'SUMPRODUCT((Payment_Applications!B1:B823=$L$2)*' +
      '((Payment_Applications!L1:L823="Certified")+(Payment_Applications!L1:L823="Paid"))*' +
      '(Payment_Applications!K1:K823))/1000000), 0)',
    '#,##0.0',
    'AB2:AB13',
    THEME.success
  );

  // Card 3: Approved VO Value (M AED)
  drawCard_(sh, 11, 5,
    'Approved VO Value (M AED)',
    '=IFERROR(IF($L$2="All Projects",' +
      'SUMPRODUCT((Variation_Orders!I1:I515="Approved")*(Variation_Orders!G1:G515))/1000000,' +
      'SUMPRODUCT((Variation_Orders!B1:B515=$L$2)*' +
      '(Variation_Orders!I1:I515="Approved")*(Variation_Orders!G1:G515))/1000000), 0)',
    '#,##0.0',
    'AC2:AC13',
    THEME.warning
  );

  // Card 4: Total PO Value (M AED) -- H=amount_aed
  drawCard_(sh, 15, 5,
    'Total PO Value (M AED)',
    '=IFERROR(IF($L$2="All Projects",' +
      'SUM(Purchase_Orders!H2:H2745)/1000000,' +
      'SUMIF(Purchase_Orders!B2:B2745,$L$2,Purchase_Orders!H2:H2745)/1000000), 0)',
    '#,##0.0',
    'AD2:AD13',
    THEME.info
  );
}

// ---------------------------------------------------------------------------
//  KPI ROW 2 (rows 9-11) - Operational metrics
// ---------------------------------------------------------------------------
/**
 * Second row of four KPI cards: active projects, inspection pass rate,
 * open safety incidents, equipment utilisation.
 */
function buildKPIRow2_(sh) {
  // Card 5: Active Projects (count of "In Progress")
  drawCard_(sh, 3, 9,
    'Active Projects',
    '=IFERROR(IF($L$2="All Projects",' +
      'COUNTIF(Projects!L2:L101,"In Progress"),' +
      'IF(INDEX(Projects!L:L,MATCH($L$2,Projects!A:A,0))="In Progress",1,0)), 0)',
    '#,##0',
    null,
    THEME.teal
  );

  // Card 6: Inspection Pass Rate (%)
  drawCard_(sh, 7, 9,
    'Inspection Pass Rate (%)',
    '=IFERROR(IF($L$2="All Projects",' +
      'COUNTIF(Inspections!G2:G3974,"Passed")/COUNTA(Inspections!G2:G3974)*100,' +
      'COUNTIFS(Inspections!B2:B3974,$L$2,Inspections!G2:G3974,"Passed")/' +
      'COUNTIF(Inspections!B2:B3974,$L$2)*100), 0)',
    '#,##0.0',
    null,
    THEME.purple
  );

  // Card 7: Open Safety Incidents
  drawCard_(sh, 11, 9,
    'Open Safety Incidents',
    '=IFERROR(IF($L$2="All Projects",' +
      'COUNTIF(Safety_Incidents!L2:L696,"Open"),' +
      'COUNTIFS(Safety_Incidents!B2:B696,$L$2,Safety_Incidents!L2:L696,"Open")), 0)',
    '#,##0',
    null,
    THEME.danger
  );

  // Card 8: Equipment Utilization (%) - no project filter (Equipment has no project col)
  drawCard_(sh, 15, 9,
    'Equipment Utilization (%)',
    '=IFERROR(COUNTIF(Equipment!G2:G151,"In Use")/COUNTA(Equipment!G2:G151)*100, 0)',
    '#,##0.0',
    null,
    THEME.orange
  );
}

// ---------------------------------------------------------------------------
//  DATA ENGINE - Hidden helper columns (Z+)
// ---------------------------------------------------------------------------
/**
 * Populates all QUERY formulas and computed data in hidden columns that
 * feed sparklines and charts. All use $L$2 project filter via Z2/Z3.
 *
 * Column map:
 *   AA-AD (27-30) : Sparkline monthly trends
 *   AI-AJ (35-36) : Budget utilisation (pie)
 *   AK-AL (37-38) : Payment status breakdown
 *   AM-AN (39-40) : Contract type split
 *   AO-AP (41-42) : Project status
 *   AQ-AR (43-44) : Project type distribution
 *   AS-AT (45-46) : PO by status
 *   AU-AV (47-48) : VO by reason
 *   AW-AX (49-50) : VO value trend (monthly)
 *   AY-AZ (51-52) : Permit status
 *   BA-BB (53-54) : Inspection results
 *   BC-BD (55-56) : Safety incident types
 *   BE-BF (57-58) : Safety severity
 *   BG-BI (59-61) : Manpower trend (3 cols)
 *   BJ-BK (62-63) : Equipment status
 *   BL-BM (64-65) : Document types
 *   BN-BO (66-67) : Employee departments
 */
function buildDataEngine_(sh) {

  // =====================================================================
  //  SPARKLINE DATA (AA-AD, cols 27-30) - 12 monthly data points
  // =====================================================================

  // AA: Monthly contract values by signed_date month
  sh.getRange('AA1').setValue('Contract_Monthly');
  sh.getRange('AA2').setFormula(
    '=IFERROR(QUERY(Contracts!A:O,' +
    '"SELECT MONTH(G)+1, SUM(F) WHERE G IS NOT NULL' +
    '" & $Z$2 & " GROUP BY MONTH(G)+1 ORDER BY MONTH(G)+1' +
    ' LABEL MONTH(G)+1 \'\', SUM(F) \'\'", 1), {"N/A",0})'
  );

  // AB: Monthly certified+paid payments by submission_date month
  sh.getRange('AB1').setValue('Certified_Monthly');
  sh.getRange('AB2').setFormula(
    '=IFERROR(QUERY(Payment_Applications!A:M,' +
    '"SELECT MONTH(F)+1, SUM(K) WHERE F IS NOT NULL' +
    ' AND (L=\'Certified\' OR L=\'Paid\')' +
    '" & $Z$2 & " GROUP BY MONTH(F)+1 ORDER BY MONTH(F)+1' +
    ' LABEL MONTH(F)+1 \'\', SUM(K) \'\'", 1), {"N/A",0})'
  );

  // AC: Monthly approved VO values by submitted_date month
  sh.getRange('AC1').setValue('VO_Monthly');
  sh.getRange('AC2').setFormula(
    '=IFERROR(QUERY(Variation_Orders!A:K,' +
    '"SELECT MONTH(F)+1, SUM(G) WHERE F IS NOT NULL AND I=\'Approved\'' +
    '" & $Z$2 & " GROUP BY MONTH(F)+1 ORDER BY MONTH(F)+1' +
    ' LABEL MONTH(F)+1 \'\', SUM(G) \'\'", 1), {"N/A",0})'
  );

  // AD: Monthly safety incident counts by incident_date month
  sh.getRange('AD1').setValue('Incidents_Monthly');
  sh.getRange('AD2').setFormula(
    '=IFERROR(QUERY(Safety_Incidents!A:L,' +
    '"SELECT MONTH(C)+1, COUNT(A) WHERE C IS NOT NULL' +
    '" & $Z$2 & " GROUP BY MONTH(C)+1 ORDER BY MONTH(C)+1' +
    ' LABEL MONTH(C)+1 \'\', COUNT(A) \'\'", 1), {"N/A",0})'
  );

  // =====================================================================
  //  AI-AJ (cols 35-36): Budget Utilisation donut - 2 rows
  //  Uses SUMPRODUCT (not QUERY) for certified+paid vs remaining
  // =====================================================================
  sh.getRange('AI1').setValue('Category');
  sh.getRange('AJ1').setValue('Value');
  sh.getRange('AI2').setValue('Certified & Paid');
  sh.getRange('AI3').setValue('Remaining');

  sh.getRange('AJ2').setFormula(
    '=IFERROR(IF($L$2="All Projects",' +
    'SUMPRODUCT((Payment_Applications!L1:L823="Certified")*' +
    '(Payment_Applications!K1:K823))+' +
    'SUMPRODUCT((Payment_Applications!L1:L823="Paid")*' +
    '(Payment_Applications!K1:K823)),' +
    'SUMPRODUCT((Payment_Applications!B1:B823=$L$2)*' +
    '(Payment_Applications!L1:L823="Certified")*' +
    '(Payment_Applications!K1:K823))+' +
    'SUMPRODUCT((Payment_Applications!B1:B823=$L$2)*' +
    '(Payment_Applications!L1:L823="Paid")*' +
    '(Payment_Applications!K1:K823))),0)'
  );

  sh.getRange('AJ3').setFormula(
    '=IFERROR(IF($L$2="All Projects",' +
    'SUM(Contracts!F2:F664)-AJ2,' +
    'SUMIF(Contracts!B2:B664,$L$2,Contracts!F2:F664)-AJ2),0)'
  );

  // =====================================================================
  //  AK-AL (cols 37-38): Payment Status Breakdown
  //  Payment_Applications: L=status, K=net_certified_aed
  // =====================================================================
  sh.getRange('AK1').setValue('Payment_Status');
  sh.getRange('AL1').setValue('Amount_AED');
  sh.getRange('AK2').setFormula(
    '=IFERROR(QUERY(Payment_Applications!A:M,' +
    '"SELECT L, SUM(K) WHERE L IS NOT NULL' +
    '" & $Z$2 & " GROUP BY L ORDER BY SUM(K) DESC' +
    ' LABEL L \'\', SUM(K) \'\'", 1), {"No Data",0})'
  );

  // =====================================================================
  //  AM-AN (cols 39-40): Contract Type Split
  //  Contracts: C=contract_type, F=contract_value_aed
  // =====================================================================
  sh.getRange('AM1').setValue('Contract_Type');
  sh.getRange('AN1').setValue('Total_Value');
  sh.getRange('AM2').setFormula(
    '=IFERROR(QUERY(Contracts!A:O,' +
    '"SELECT C, SUM(F) WHERE C IS NOT NULL' +
    '" & $Z$2 & " GROUP BY C ORDER BY SUM(F) DESC' +
    ' LABEL C \'\', SUM(F) \'\'", 1), {"No Data",0})'
  );

  // =====================================================================
  //  AO-AP (cols 41-42): Project Status Distribution
  //  Projects: L=status -- uses Z3 (project in col A)
  // =====================================================================
  sh.getRange('AO1').setValue('Project_Status');
  sh.getRange('AP1').setValue('Count');
  sh.getRange('AO2').setFormula(
    '=IFERROR(QUERY(Projects!A:P,' +
    '"SELECT L, COUNT(L) WHERE L IS NOT NULL' +
    '" & $Z$3 & " GROUP BY L ORDER BY COUNT(L) DESC' +
    ' LABEL L \'\', COUNT(L) \'\'", 1), {"No Data",0})'
  );

  // =====================================================================
  //  AQ-AR (cols 43-44): Project Type Distribution
  //  Projects: C=project_type -- uses Z3
  // =====================================================================
  sh.getRange('AQ1').setValue('Project_Type');
  sh.getRange('AR1').setValue('Count');
  sh.getRange('AQ2').setFormula(
    '=IFERROR(QUERY(Projects!A:P,' +
    '"SELECT C, COUNT(C) WHERE C IS NOT NULL' +
    '" & $Z$3 & " GROUP BY C ORDER BY COUNT(C) DESC' +
    ' LABEL C \'\', COUNT(C) \'\'", 1), {"No Data",0})'
  );

  // =====================================================================
  //  AS-AT (cols 45-46): Purchase Orders by Status
  //  Purchase_Orders: I=status, H=amount_aed
  // =====================================================================
  sh.getRange('AS1').setValue('PO_Status');
  sh.getRange('AT1').setValue('Amount_AED');
  sh.getRange('AS2').setFormula(
    '=IFERROR(QUERY(Purchase_Orders!A:J,' +
    '"SELECT I, SUM(H) WHERE I IS NOT NULL' +
    '" & $Z$2 & " GROUP BY I ORDER BY SUM(H) DESC' +
    ' LABEL I \'\', SUM(H) \'\'", 1), {"No Data",0})'
  );

  // =====================================================================
  //  AU-AV (cols 47-48): Variation Orders by Reason
  //  Variation_Orders: E=reason, G=vo_value_aed
  // =====================================================================
  sh.getRange('AU1').setValue('VO_Reason');
  sh.getRange('AV1').setValue('Total_Value');
  sh.getRange('AU2').setFormula(
    '=IFERROR(QUERY(Variation_Orders!A:K,' +
    '"SELECT E, SUM(G) WHERE E IS NOT NULL' +
    '" & $Z$2 & " GROUP BY E ORDER BY SUM(G) DESC' +
    ' LABEL E \'\', SUM(G) \'\'", 1), {"No Data",0})'
  );

  // =====================================================================
  //  AW-AX (cols 49-50): VO Value Trend (Monthly)
  //  Variation_Orders: F=submitted_date, G=vo_value_aed
  // =====================================================================
  sh.getRange('AW1').setValue('Month');
  sh.getRange('AX1').setValue('VO_Value');
  sh.getRange('AW2').setFormula(
    '=IFERROR(QUERY(Variation_Orders!A:K,' +
    '"SELECT MONTH(F)+1, SUM(G) WHERE F IS NOT NULL' +
    '" & $Z$2 & " GROUP BY MONTH(F)+1 ORDER BY MONTH(F)+1' +
    ' LABEL MONTH(F)+1 \'\', SUM(G) \'\'", 1), {"N/A",0})'
  );

  // =====================================================================
  //  AY-AZ (cols 51-52): Permit Status
  //  Permits_Approvals: H=status
  // =====================================================================
  sh.getRange('AY1').setValue('Permit_Status');
  sh.getRange('AZ1').setValue('Count');
  sh.getRange('AY2').setFormula(
    '=IFERROR(QUERY(Permits_Approvals!A:J,' +
    '"SELECT H, COUNT(H) WHERE H IS NOT NULL' +
    '" & $Z$2 & " GROUP BY H ORDER BY COUNT(H) DESC' +
    ' LABEL H \'\', COUNT(H) \'\'", 1), {"No Data",0})'
  );

  // =====================================================================
  //  BA-BB (cols 53-54): Inspection Results
  //  Inspections: G=result
  // =====================================================================
  sh.getRange('BA1').setValue('Insp_Result');
  sh.getRange('BB1').setValue('Count');
  sh.getRange('BA2').setFormula(
    '=IFERROR(QUERY(Inspections!A:K,' +
    '"SELECT G, COUNT(G) WHERE G IS NOT NULL' +
    '" & $Z$2 & " GROUP BY G ORDER BY COUNT(G) DESC' +
    ' LABEL G \'\', COUNT(G) \'\'", 1), {"No Data",0})'
  );

  // =====================================================================
  //  BC-BD (cols 55-56): Safety Incident Types
  //  Safety_Incidents: D=incident_type
  // =====================================================================
  sh.getRange('BC1').setValue('Incident_Type');
  sh.getRange('BD1').setValue('Count');
  sh.getRange('BC2').setFormula(
    '=IFERROR(QUERY(Safety_Incidents!A:L,' +
    '"SELECT D, COUNT(D) WHERE D IS NOT NULL' +
    '" & $Z$2 & " GROUP BY D ORDER BY COUNT(D) DESC' +
    ' LABEL D \'\', COUNT(D) \'\'", 1), {"No Data",0})'
  );

  // =====================================================================
  //  BE-BF (cols 57-58): Safety Severity
  //  Safety_Incidents: E=severity
  // =====================================================================
  sh.getRange('BE1').setValue('Severity');
  sh.getRange('BF1').setValue('Count');
  sh.getRange('BE2').setFormula(
    '=IFERROR(QUERY(Safety_Incidents!A:L,' +
    '"SELECT E, COUNT(E) WHERE E IS NOT NULL' +
    '" & $Z$2 & " GROUP BY E ORDER BY COUNT(E) DESC' +
    ' LABEL E \'\', COUNT(E) \'\'", 1), {"No Data",0})'
  );

  // =====================================================================
  //  BG-BI (cols 59-61): Manpower Trend (Monthly from Daily_Site_Reports)
  //  Daily_Site_Reports: C=report_date, G=manpower_direct, H=manpower_subcon
  // =====================================================================
  sh.getRange('BG1').setValue('Month');
  sh.getRange('BH1').setValue('Direct');
  sh.getRange('BI1').setValue('Subcon');
  sh.getRange('BG2').setFormula(
    '=IFERROR(QUERY(Daily_Site_Reports!A:L,' +
    '"SELECT MONTH(C)+1, AVG(G), AVG(H) WHERE C IS NOT NULL' +
    '" & $Z$2 & " GROUP BY MONTH(C)+1 ORDER BY MONTH(C)+1' +
    ' LABEL MONTH(C)+1 \'\', AVG(G) \'\', AVG(H) \'\'", 1), {"N/A",0,0})'
  );

  // =====================================================================
  //  BJ-BK (cols 62-63): Equipment Status
  //  Equipment: G=status -- NO project filter (Equipment has no project col)
  // =====================================================================
  sh.getRange('BJ1').setValue('Equip_Status');
  sh.getRange('BK1').setValue('Count');
  sh.getRange('BJ2').setFormula(
    '=IFERROR(QUERY(Equipment!A:J,' +
    '"SELECT G, COUNT(G) WHERE G IS NOT NULL' +
    ' GROUP BY G ORDER BY COUNT(G) DESC' +
    ' LABEL G \'\', COUNT(G) \'\'", 1), {"No Data",0})'
  );

  // =====================================================================
  //  BL-BM (cols 64-65): Document Types
  //  Project_Documents: C=document_type (only PRJ-001 data)
  // =====================================================================
  sh.getRange('BL1').setValue('Doc_Type');
  sh.getRange('BM1').setValue('Count');
  sh.getRange('BL2').setFormula(
    '=IFERROR(QUERY(Project_Documents!A:K,' +
    '"SELECT C, COUNT(C) WHERE C IS NOT NULL' +
    '" & $Z$2 & " GROUP BY C ORDER BY COUNT(C) DESC' +
    ' LABEL C \'\', COUNT(C) \'\'", 1), {"No Data",0})'
  );

  // =====================================================================
  //  BN-BO (cols 66-67): Employees by Department
  //  Employees: C=department -- NO project filter
  // =====================================================================
  sh.getRange('BN1').setValue('Department');
  sh.getRange('BO1').setValue('Count');
  sh.getRange('BN2').setFormula(
    '=IFERROR(QUERY(Employees!A:K,' +
    '"SELECT C, COUNT(C) WHERE C IS NOT NULL' +
    ' GROUP BY C ORDER BY COUNT(C) DESC' +
    ' LABEL C \'\', COUNT(C) \'\'", 1), {"No Data",0})'
  );
}

// ---------------------------------------------------------------------------
//  SECTION HEADERS
// ---------------------------------------------------------------------------
/**
 * Builds the five section header rows with styled titles and accent underlines.
 */
function buildSections_(sh) {
  var sections = [
    { row: 13,  title: 'FINANCIAL OVERVIEW' },
    { row: 43,  title: 'PROJECT & OPERATIONS' },
    { row: 74,  title: 'PROCUREMENT & RISK' },
    { row: 105, title: 'HSE & RESOURCES' },
    { row: 136, title: 'PROJECT HEALTH SUMMARY' }
  ];

  for (var i = 0; i < sections.length; i++) {
    var s = sections[i];
    sh.setRowHeight(s.row, 32);

    var hdr = sh.getRange(s.row, 3, 1, 15);
    hdr.merge()
       .setValue('\u25C8  ' + s.title)
       .setFontFamily(THEME.font)
       .setFontSize(13)
       .setFontWeight('bold')
       .setFontColor(THEME.headerText)
       .setVerticalAlignment('middle')
       .setBackground(THEME.bg);

    // Accent underline
    hdr.setBorder(null, null, true, null, null, null,
                  THEME.accent, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
}

// ---------------------------------------------------------------------------
//  CHART BUILDER (reusable)
// ---------------------------------------------------------------------------
/**
 * Creates and inserts a single embedded chart on the dashboard.
 *
 * @param {Sheet}  sh    Dashboard sheet
 * @param {string} type  'PIE','COLUMN','BAR','LINE','AREA'
 * @param {string} range Data range string e.g. 'AI1:AJ3'
 * @param {number} row   Anchor row
 * @param {number} col   Anchor column
 * @param {string} title Chart title
 * @param {Object} opts  Optional overrides: {pieHole, width, height, colors,
 *                        stacked, legend}
 */
function addChart_(sh, type, range, row, col, title, opts) {
  opts = opts || {};
  var width  = opts.width  || 680;
  var height = opts.height || 300;
  var colors = opts.colors || THEME.palette;

  // Map string type to Charts.ChartType enum
  var chartType;
  switch (type) {
    case 'PIE':    chartType = Charts.ChartType.PIE;    break;
    case 'COLUMN': chartType = Charts.ChartType.COLUMN; break;
    case 'BAR':    chartType = Charts.ChartType.BAR;    break;
    case 'LINE':   chartType = Charts.ChartType.LINE;   break;
    case 'AREA':   chartType = Charts.ChartType.AREA;   break;
    default:       chartType = Charts.ChartType.COLUMN;
  }

  var dataRange = sh.getRange(range);

  var builder = sh.newChart()
    .setChartType(chartType)
    .addRange(dataRange)
    .setPosition(row, col, 0, 0)
    .setOption('title', title)
    .setOption('titleTextStyle', {
      color: THEME.headerText,
      fontName: THEME.font,
      fontSize: 12,
      bold: true
    })
    .setOption('fontName', THEME.font)
    .setOption('backgroundColor', { fill: THEME.cardBg })
    .setOption('chartArea', {
      left: '12%', top: '15%', width: '76%', height: '68%'
    })
    .setOption('legend', {
      position: opts.legend || 'right',
      textStyle: {
        color: THEME.subText,
        fontName: THEME.font,
        fontSize: 10
      }
    })
    .setOption('colors', colors)
    .setOption('width', width)
    .setOption('height', height)
    .setOption('hAxis', {
      textStyle: { color: THEME.subText, fontName: THEME.font, fontSize: 9 },
      gridlines: { color: THEME.border },
      baselineColor: THEME.border
    })
    .setOption('vAxis', {
      textStyle: { color: THEME.subText, fontName: THEME.font, fontSize: 9 },
      gridlines: { color: THEME.border },
      baselineColor: THEME.border,
      format: 'short'
    })
    .setOption('tooltip', {
      textStyle: { fontName: THEME.font, fontSize: 11 }
    })
    .setOption('animation', {
      startup: true, duration: 500, easing: 'out'
    })
    .setOption('useFirstColumnAsDomain', true);

  // -- Type-specific options --

  if (type === 'PIE') {
    if (opts.pieHole !== undefined) {
      builder.setOption('pieHole', opts.pieHole);
    }
    builder.setOption('pieSliceTextStyle', {
      fontName: THEME.font, fontSize: 10, color: '#FFFFFF'
    });
    builder.setOption('chartArea', {
      left: '5%', top: '15%', width: '90%', height: '72%'
    });
  }

  if (type === 'AREA') {
    builder.setOption('areaOpacity', 0.3);
    builder.setOption('isStacked', opts.stacked || false);
  }

  if ((type === 'COLUMN' || type === 'BAR') && opts.stacked) {
    builder.setOption('isStacked', true);
  }

  sh.insertChart(builder.build());
}

// ---------------------------------------------------------------------------
//  ALL 16 CHARTS
// ---------------------------------------------------------------------------
/**
 * Creates all 16 dashboard charts across four visual sections.
 *
 * Left charts : position (row, 3),  width 680, height 300
 * Right charts: position (row, 11), width 660, height 300
 */
function buildAllCharts_(sh) {

  // =====================================================================
  //  SECTION 1: FINANCIAL OVERVIEW (header at row 13)
  // =====================================================================

  // 1. Budget Utilisation - Donut PIE (row 14, left)
  addChart_(sh, 'PIE', 'AI1:AJ3', 14, 3,
    'Budget Utilization', {
      pieHole: 0.6,
      width: 680,
      height: 300,
      colors: [THEME.success, THEME.border],
      legend: 'right'
    }
  );

  // 2. Payment Cashflow by Status - COLUMN (row 14, right)
  addChart_(sh, 'COLUMN', 'AK1:AL15', 14, 11,
    'Payment Cashflow by Status', {
      width: 660,
      height: 300,
      colors: [THEME.accent, THEME.info, THEME.success, THEME.warning],
      legend: 'none'
    }
  );

  // 3. Contract Value by Type - BAR (row 29, left)
  addChart_(sh, 'BAR', 'AM1:AN15', 29, 3,
    'Contract Value by Type', {
      width: 680,
      height: 300,
      colors: [THEME.purple],
      legend: 'none'
    }
  );

  // 4. VO Value Trend (Monthly) - LINE (row 29, right)
  addChart_(sh, 'LINE', 'AW1:AX15', 29, 11,
    'Variation Order Value Trend (Monthly)', {
      width: 660,
      height: 300,
      colors: [THEME.warning],
      legend: 'none'
    }
  );

  // =====================================================================
  //  SECTION 2: PROJECT & OPERATIONS (header at row 43)
  // =====================================================================

  // 5. Project Status Distribution - BAR (row 44, left)
  addChart_(sh, 'BAR', 'AO1:AP15', 44, 3,
    'Project Status Distribution', {
      width: 680,
      height: 300,
      colors: [THEME.accent],
      legend: 'none'
    }
  );

  // 6. Project Type Distribution - PIE (row 44, right)
  addChart_(sh, 'PIE', 'AQ1:AR15', 44, 11,
    'Project Type Distribution', {
      width: 660,
      height: 300,
      legend: 'right'
    }
  );

  // 7. Manpower Trend (Monthly Avg) - AREA (row 59, left)
  addChart_(sh, 'AREA', 'BG1:BI15', 59, 3,
    'Manpower Trend (Monthly Avg)', {
      width: 680,
      height: 300,
      colors: [THEME.info, THEME.teal],
      stacked: true,
      legend: 'top'
    }
  );

  // 8. Inspection Results - COLUMN (row 59, right)
  addChart_(sh, 'COLUMN', 'BA1:BB15', 59, 11,
    'Inspection Results', {
      width: 660,
      height: 300,
      colors: [THEME.success, THEME.warning, THEME.danger],
      legend: 'none'
    }
  );

  // =====================================================================
  //  SECTION 3: PROCUREMENT & RISK (header at row 74)
  // =====================================================================

  // 9. Purchase Orders by Status - COLUMN (row 75, left)
  addChart_(sh, 'COLUMN', 'AS1:AT15', 75, 3,
    'Purchase Orders by Status', {
      width: 680,
      height: 300,
      colors: [THEME.info],
      legend: 'none'
    }
  );

  // 10. Variation Orders by Reason - BAR (row 75, right)
  addChart_(sh, 'BAR', 'AU1:AV15', 75, 11,
    'Variation Orders by Reason', {
      width: 660,
      height: 300,
      colors: [THEME.orange],
      legend: 'none'
    }
  );

  // 11. Permit Status Overview - Donut PIE (row 90, left)
  addChart_(sh, 'PIE', 'AY1:AZ15', 90, 3,
    'Permit Status Overview', {
      pieHole: 0.5,
      width: 680,
      height: 300,
      colors: [THEME.success, THEME.warning, THEME.danger],
      legend: 'right'
    }
  );

  // 12. Document Types Distribution - BAR (row 90, right)
  addChart_(sh, 'BAR', 'BL1:BM15', 90, 11,
    'Document Types Distribution', {
      width: 660,
      height: 300,
      colors: [THEME.purple],
      legend: 'none'
    }
  );

  // =====================================================================
  //  SECTION 4: HSE & RESOURCES (header at row 105)
  // =====================================================================

  // 13. Safety Incidents by Type - BAR (row 106, left)
  addChart_(sh, 'BAR', 'BC1:BD15', 106, 3,
    'Safety Incidents by Type', {
      width: 680,
      height: 300,
      colors: [THEME.danger],
      legend: 'none'
    }
  );

  // 14. Safety Incident Severity - Donut PIE (row 106, right)
  addChart_(sh, 'PIE', 'BE1:BF15', 106, 11,
    'Safety Incident Severity', {
      pieHole: 0.4,
      width: 660,
      height: 300,
      colors: [THEME.danger, THEME.warning, THEME.success],
      legend: 'right'
    }
  );

  // 15. Equipment Status - PIE (row 121, left)
  addChart_(sh, 'PIE', 'BJ1:BK15', 121, 3,
    'Equipment Status', {
      width: 680,
      height: 300,
      colors: [THEME.success, THEME.info, THEME.warning],
      legend: 'right'
    }
  );

  // 16. Employees by Department - COLUMN (row 121, right)
  addChart_(sh, 'COLUMN', 'BN1:BO15', 121, 11,
    'Employees by Department', {
      width: 660,
      height: 300,
      colors: [THEME.accent],
      legend: 'none'
    }
  );
}

// ---------------------------------------------------------------------------
//  HEALTH TABLE (row 136+)
// ---------------------------------------------------------------------------
/**
 * Builds the Project Health Summary table below the charts.
 * Rows 137 (headers) and 138-147 (top 10 projects) with conditional
 * formatting for status and completion percentage.
 */
function buildHealthTable_(sh) {
  // Section header at row 136 is already built in buildSections_()

  // -- Column headers at row 137 --
  var headers = [
    'Project ID', 'Project Name', 'Status', 'Completion %',
    'Contract Value', 'Certified Amt', 'Open VOs', 'Open Incidents'
  ];

  // Column spans: each header occupies a range of columns across C-Q
  var spans = [
    [3, 4],    // C-D : Project ID
    [5, 6],    // E-F : Project Name
    [7, 8],    // G-H : Status
    [9, 10],   // I-J : Completion %
    [11, 12],  // K-L : Contract Value
    [13, 14],  // M-N : Certified Amt
    [15, 15],  // O   : Open VOs
    [16, 17]   // P-Q : Open Incidents
  ];

  sh.setRowHeight(137, 30);

  for (var h = 0; h < headers.length; h++) {
    var startCol = spans[h][0];
    var endCol   = spans[h][1];
    var hRange   = sh.getRange(137, startCol, 1, endCol - startCol + 1);
    hRange.merge()
      .setValue(headers[h])
      .setFontFamily(THEME.font)
      .setFontSize(10)
      .setFontWeight('bold')
      .setFontColor('#FFFFFF')
      .setBackground(THEME.sidebarBg)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true,
                 THEME.sidebarBg, SpreadsheetApp.BorderStyle.SOLID);
  }

  // -- Data rows 138-147 (top 10 projects from Projects sheet) --
  for (var r = 0; r < 10; r++) {
    var dataRow = 138 + r;
    var projRow = r + 2;  // Row in Projects sheet (1-indexed, row 1 is header)
    var rowBg   = (r % 2 === 0) ? THEME.cardBg : '#F1F5F9';

    sh.setRowHeight(dataRow, 26);

    // Helper: apply common cell styling
    var styleCell_ = function(range, align) {
      range.setFontFamily(THEME.font)
           .setFontSize(9)
           .setFontColor(THEME.headerText)
           .setHorizontalAlignment(align || 'center')
           .setVerticalAlignment('middle')
           .setBackground(rowBg)
           .setBorder(true, true, true, true, false, false,
                      THEME.border, SpreadsheetApp.BorderStyle.SOLID);
    };

    // Project ID (C-D)
    var idR = sh.getRange(dataRow, 3, 1, 2);
    idR.merge().setFormula('=IFERROR(INDEX(Projects!A:A,' + projRow + '),"")');
    styleCell_(idR, 'center');

    // Project Name (E-F)
    var nmR = sh.getRange(dataRow, 5, 1, 2);
    nmR.merge().setFormula('=IFERROR(INDEX(Projects!B:B,' + projRow + '),"")');
    styleCell_(nmR, 'left');

    // Status (G-H)
    var stR = sh.getRange(dataRow, 7, 1, 2);
    stR.merge().setFormula('=IFERROR(INDEX(Projects!L:L,' + projRow + '),"")');
    styleCell_(stR, 'center');
    stR.setFontWeight('bold');

    // Completion % (I-J)
    var cpR = sh.getRange(dataRow, 9, 1, 2);
    cpR.merge().setFormula('=IFERROR(INDEX(Projects!M:M,' + projRow + '),"")');
    cpR.setNumberFormat('#,##0"%"');
    styleCell_(cpR, 'center');

    // Contract Value (K-L) - sum of contracts for this project
    var cvR = sh.getRange(dataRow, 11, 1, 2);
    cvR.merge().setFormula(
      '=IFERROR(SUMIF(Contracts!B:B,INDEX(Projects!A:A,' + projRow + '),Contracts!F:F),"")'
    );
    cvR.setNumberFormat('#,##0');
    styleCell_(cvR, 'right');

    // Certified Amt (M-N) - sum of certified+paid net amounts
    var caR = sh.getRange(dataRow, 13, 1, 2);
    caR.merge().setFormula(
      '=IFERROR(SUMPRODUCT(' +
      '(Payment_Applications!B$1:B$823=INDEX(Projects!A:A,' + projRow + '))*' +
      '((Payment_Applications!L$1:L$823="Certified")+(Payment_Applications!L$1:L$823="Paid"))*' +
      '(Payment_Applications!K$1:K$823)),"")'
    );
    caR.setNumberFormat('#,##0');
    styleCell_(caR, 'right');

    // Open VOs (O) - count of VOs not Approved and not Rejected
    var voR = sh.getRange(dataRow, 15, 1, 1);
    voR.setFormula(
      '=IFERROR(COUNTIFS(Variation_Orders!B:B,INDEX(Projects!A:A,' + projRow + '),' +
      'Variation_Orders!I:I,"<>Approved",Variation_Orders!I:I,"<>Rejected"),"")'
    );
    voR.setNumberFormat('#,##0');
    styleCell_(voR, 'center');

    // Open Incidents (P-Q)
    var incR = sh.getRange(dataRow, 16, 1, 2);
    incR.merge().setFormula(
      '=IFERROR(COUNTIFS(Safety_Incidents!B:B,INDEX(Projects!A:A,' + projRow + '),' +
      'Safety_Incidents!L:L,"Open"),"")'
    );
    incR.setNumberFormat('#,##0');
    styleCell_(incR, 'center');
  }

  // -- Conditional formatting --
  var statusRange = sh.getRange('G138:H147');
  var compRange   = sh.getRange('I138:J147');

  // Status colours
  var ruleCompleted = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Completed')
    .setBackground('#D1FAE5').setFontColor('#065F46')
    .setRanges([statusRange]).build();

  var ruleInProgress = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('In Progress')
    .setBackground('#FEF3C7').setFontColor('#92400E')
    .setRanges([statusRange]).build();

  var ruleDelayed = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Delayed')
    .setBackground('#FFEDD5').setFontColor('#9A3412')
    .setRanges([statusRange]).build();

  var ruleOnHold = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('On Hold')
    .setBackground('#FEE2E2').setFontColor('#991B1B')
    .setRanges([statusRange]).build();

  // Completion % colour scale
  var ruleCompLow = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(30)
    .setBackground('#FEE2E2').setFontColor('#991B1B')
    .setRanges([compRange]).build();

  var ruleCompMid = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(30, 69)
    .setBackground('#FEF3C7').setFontColor('#92400E')
    .setRanges([compRange]).build();

  var ruleCompHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(70)
    .setBackground('#D1FAE5').setFontColor('#065F46')
    .setRanges([compRange]).build();

  var rules = sh.getConditionalFormatRules();
  rules.push(ruleCompleted, ruleInProgress, ruleDelayed, ruleOnHold);
  rules.push(ruleCompLow, ruleCompMid, ruleCompHigh);
  sh.setConditionalFormatRules(rules);

  // -- Footer note --
  var footerRange = sh.getRange(149, 3, 1, 15);
  footerRange.merge()
    .setValue('Data sourced from 17 integrated modules  |  Contract values in AED  |  Use project filter dropdown to drill down')
    .setFontFamily(THEME.font)
    .setFontSize(8)
    .setFontColor(THEME.subText)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(THEME.bg)
    .setFontStyle('italic');
}
