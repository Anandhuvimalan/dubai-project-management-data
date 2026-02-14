/**
 * DUBAI ERP - OPERATIONS COMMAND CENTER v7.0
 * 25+ Advanced Operational Intelligence Panels
 * All data computed in-script. Filter-reactive tables + KPIs.
 * onEdit rebuilds tables on filter change. Checkbox reset.
 */

var THEME = {
  bg: '#F8FAFC', cardBg: '#FFFFFF', sidebarBg: '#0F172A', sidebarText: '#94A3B8',
  sidebarActive: '#818CF8', headerText: '#0F172A', subText: '#64748B',
  border: '#E2E8F0', accent: '#6366F1', success: '#10B981', warning: '#F59E0B',
  danger: '#EF4444', info: '#3B82F6', teal: '#14B8A6', purple: '#8B5CF6',
  pink: '#EC4899', orange: '#F97316', filterBar: '#1E293B',
  gBg: '#D1FAE5', gFg: '#065F46', yBg: '#FEF3C7', yFg: '#92400E',
  rBg: '#FEE2E2', rFg: '#991B1B', bBg: '#DBEAFE', bFg: '#1E40AF',
  palette: ['#6366F1', '#8B5CF6', '#EC4899', '#F43F5E', '#F97316', '#EAB308',
    '#22C55E', '#14B8A6', '#06B6D4', '#3B82F6', '#A78BFA', '#FB923C'],
  font: 'Inter, Roboto, sans-serif'
};
var DN = 'ERP Dashboard';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('ERP Dashboard')
    .addItem('Create / Refresh Dashboard', 'createERPDashboard').addToUi();
}

/* ---- helpers ---- */
function rd_(n, c) { var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); if (!s) return []; var lr = s.getLastRow(); if (lr < 2) return []; return s.getRange(2, 1, lr - 1, c).getValues(); }
function uq_(d, i) { var s = {}, o = []; d.forEach(function (r) { var v = String(r[i]).trim(); if (v && v !== 'undefined' && v !== '') { if (!s[v]) { s[v] = true; o.push(v); } } }); return o.sort(); }
function db_(a, b) { if (!(a instanceof Date) || !(b instanceof Date)) return 0; return Math.round((b - a) / 864e5); }
function fm_(v) { return Math.round(v / 1e5) / 10; }
function pf_(sh, r, c, t, bg) { sh.getRange(r, c, 1, t).merge().setBackground(bg || THEME.filterBar); }

/* style cell */
function sc_(rng, bg, al) {
  rng.setFontFamily(THEME.font).setFontSize(9).setFontColor(THEME.headerText)
    .setHorizontalAlignment(al || 'center').setVerticalAlignment('middle').setBackground(bg)
    .setBorder(true, true, true, true, false, false, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
}
/* table header row */
function th_(sh, row, sp, hd) {
  for (var h = 0; h < hd.length; h++) {
    sh.getRange(row, sp[h][0], 1, sp[h][1] - sp[h][0] + 1).merge().setValue(hd[h])
      .setFontFamily(THEME.font).setFontSize(9).setFontWeight('bold')
      .setFontColor('#FFF').setBackground(THEME.filterBar)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, THEME.filterBar, SpreadsheetApp.BorderStyle.SOLID);
  }
}
/* section header */
function sec_(sh, r, t) {
  sh.getRange(r, 3, 1, 15).merge().setValue(t)
    .setFontFamily(THEME.font).setFontSize(12).setFontWeight('bold')
    .setFontColor(THEME.headerText).setVerticalAlignment('middle').setBackground(THEME.bg)
    .setBorder(null, null, true, null, null, null, THEME.accent, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}
/* KPI card */
function kpi_(sh, col, row, title, formula, fmt, color) {
  var t = sh.getRange(row, col, 1, 3);
  t.merge().setValue(title).setFontFamily(THEME.font).setFontSize(8)
    .setFontColor(THEME.subText).setHorizontalAlignment('center')
    .setVerticalAlignment('bottom').setBackground(THEME.cardBg);
  t.setBorder(true, true, null, true, false, false, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, col, 1, 3).setBorder(true, null, null, null, null, null, color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  var v = sh.getRange(row + 1, col, 1, 3);
  v.merge().setFormula(formula).setNumberFormat(fmt).setFontFamily(THEME.font)
    .setFontSize(18).setFontWeight('bold').setFontColor(THEME.headerText)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(THEME.cardBg);
  v.setBorder(null, true, null, true, false, false, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
  var b = sh.getRange(row + 2, col, 1, 3);
  b.merge().setBackground(THEME.cardBg).setHorizontalAlignment('center').setVerticalAlignment('top');
  b.setBorder(null, true, true, true, false, false, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
  b.setValue(' ').setFontSize(6).setFontColor(THEME.subText);
}
/* chart builder */
function ch_(sh, type, range, row, col, title, opts) {
  opts = opts || {}; var ct;
  switch (type) {
    case 'PIE': ct = Charts.ChartType.PIE; break; case 'COLUMN': ct = Charts.ChartType.COLUMN; break;
    case 'BAR': ct = Charts.ChartType.BAR; break; case 'LINE': ct = Charts.ChartType.LINE; break;
    case 'AREA': ct = Charts.ChartType.AREA; break; default: ct = Charts.ChartType.COLUMN;
  }
  var b = sh.newChart().setChartType(ct).addRange(sh.getRange(range))
    .setPosition(row, col, 5, 5).setOption('title', title)
    .setOption('titleTextStyle', { color: THEME.headerText, fontName: THEME.font, fontSize: 11, bold: true })
    .setOption('fontName', THEME.font).setOption('backgroundColor', { fill: THEME.cardBg })
    .setOption('chartArea', { left: '14%', top: '15%', width: '74%', height: '68%' })
    .setOption('legend', { position: opts.leg || 'none', textStyle: { color: THEME.subText, fontName: THEME.font, fontSize: 9 } })
    .setOption('colors', opts.col || THEME.palette)
    .setOption('width', opts.w || 660).setOption('height', opts.h || 290)
    .setOption('hAxis', { textStyle: { color: THEME.subText, fontName: THEME.font, fontSize: 8 }, gridlines: { color: THEME.border } })
    .setOption('vAxis', { textStyle: { color: THEME.subText, fontName: THEME.font, fontSize: 8 }, gridlines: { color: THEME.border }, format: 'short' })
    .setOption('tooltip', { textStyle: { fontName: THEME.font, fontSize: 10 } })
    .setOption('animation', { startup: true, duration: 400, easing: 'out' })
    .setOption('useFirstColumnAsDomain', true);
  if (type === 'PIE') {
    if (opts.hole !== undefined) b.setOption('pieHole', opts.hole);
    b.setOption('pieSliceTextStyle', { fontName: THEME.font, fontSize: 9, color: '#FFF' });
    b.setOption('chartArea', { left: '5%', top: '15%', width: '90%', height: '72%' });
  }
  if (type === 'AREA') b.setOption('areaOpacity', 0.3);
  if ((type === 'COLUMN' || type === 'BAR' || type === 'AREA') && opts.stack) b.setOption('isStacked', true);
  sh.insertChart(b.build());
}

/* ================================================================ */
/*  MAIN                                                            */
/* ================================================================ */
function createERPDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var old = ss.getSheetByName(DN);
  if (old) { ss.deleteSheet(old); SpreadsheetApp.flush(); }
  var sh = ss.insertSheet(DN, 0);
  SpreadsheetApp.flush(); sh.clear(); SpreadsheetApp.flush();

  setupLayout_(sh);
  buildSidebar_(sh);
  buildHeader_(sh);
  buildFilterPanel_(sh);
  setupFilterValidation_(sh);

  sh.getRange('R2').setValue('v7.0').setFontFamily(THEME.font).setFontSize(8)
    .setFontColor(THEME.subText).setHorizontalAlignment('right');

  sh.getRange(1, 27, 120, 70).clearContent();
  SpreadsheetApp.flush();

  var m = buildDataEngine_(sh);
  SpreadsheetApp.flush(); Utilities.sleep(1000); SpreadsheetApp.flush();

  buildKPIRow1_(sh);
  buildKPIRow2_(sh);
  buildKPIRow3_(sh);
  computeKPIs_(sh);
  buildSections_(sh);
  buildAllCharts_(sh, m);

  // All operational tables
  buildNewProjectsTable_(sh);
  buildPhaseProgressTable_(sh);
  buildContractorTable_(sh);
  buildBudgetTable_(sh);
  buildCashFlowTable_(sh);
  buildEquipmentTable_(sh);
  buildVOTable_(sh);
  buildSafetyTable_(sh);
  buildPermitsTable_(sh);
  buildInspectionTable_(sh);
  buildPOPipelineTable_(sh);
  buildLADRiskTable_(sh);
  buildDocStatusTable_(sh);
  buildRetentionTable_(sh);
  buildDecisionMatrix_(sh);

  sh.hideColumns(27, 70);
  SpreadsheetApp.flush();
  sh.setActiveSelection('C2');
  ui.alert('ERP Operations v7.0',
    '25+ operational panels loaded.\nFilters are reactive — change Project/Type/Status dropdowns.\nClick RESET checkbox to clear all filters.',
    ui.ButtonSet.OK);
}

/* ================================================================ */
/*  LAYOUT — 140 rows                                               */
/* ================================================================ */
function setupLayout_(sh) {
  if (sh.getMaxRows() < 145) sh.insertRowsAfter(sh.getMaxRows(), 145 - sh.getMaxRows());
  if (sh.getMaxColumns() < 100) sh.insertColumnsAfter(sh.getMaxColumns(), 100 - sh.getMaxColumns());
  sh.setColumnWidth(1, 60); sh.setColumnWidth(2, 20);
  [3, 4, 5, 7, 8, 9, 11, 12, 13, 15, 16, 17].forEach(function (c) { sh.setColumnWidth(c, 110); });
  sh.setColumnWidth(6, 18); sh.setColumnWidth(10, 18); sh.setColumnWidth(14, 18);
  for (var c = 18; c <= 26; c++)sh.setColumnWidth(c, 35);
  for (var c = 27; c <= 96; c++)sh.setColumnWidth(c, 120);

  sh.setRowHeight(1, 6); sh.setRowHeight(2, 44); sh.setRowHeight(3, 22);
  sh.setRowHeight(4, 38); sh.setRowHeight(5, 6);
  // KPI rows 6-14
  sh.setRowHeight(6, 22); sh.setRowHeight(7, 40); sh.setRowHeight(8, 40); sh.setRowHeight(9, 4);
  sh.setRowHeight(10, 22); sh.setRowHeight(11, 40); sh.setRowHeight(12, 40); sh.setRowHeight(13, 4);
  sh.setRowHeight(14, 22); sh.setRowHeight(15, 40); sh.setRowHeight(16, 40); sh.setRowHeight(17, 6);
  // Section 1: New Projects + Phase (18-30)
  sh.setRowHeight(18, 30); sh.setRowHeight(19, 24);
  for (var r = 20; r <= 26; r++)sh.setRowHeight(r, 22);
  sh.setRowHeight(27, 24);
  for (var r = 28; r <= 33; r++)sh.setRowHeight(r, 22);
  sh.setRowHeight(34, 6);
  // Section 2: Project Overview Charts (35-38)
  sh.setRowHeight(35, 30); sh.setRowHeight(36, 300); sh.setRowHeight(37, 300); sh.setRowHeight(38, 6);
  // Section 3: Contractor (39-50)
  sh.setRowHeight(39, 30); sh.setRowHeight(40, 300); sh.setRowHeight(41, 300);
  sh.setRowHeight(42, 24);
  for (var r = 43; r <= 49; r++)sh.setRowHeight(r, 22);
  sh.setRowHeight(50, 6);
  // Section 4: Budget + Cash Flow (51-66)
  sh.setRowHeight(51, 30); sh.setRowHeight(52, 300); sh.setRowHeight(53, 300);
  sh.setRowHeight(54, 24);
  for (var r = 55; r <= 61; r++)sh.setRowHeight(r, 22);
  sh.setRowHeight(62, 24);
  for (var r = 63; r <= 68; r++)sh.setRowHeight(r, 22);
  sh.setRowHeight(69, 6);
  // Section 5: Equipment (70-80)
  sh.setRowHeight(70, 30); sh.setRowHeight(71, 300); sh.setRowHeight(72, 300);
  sh.setRowHeight(73, 24);
  for (var r = 74; r <= 79; r++)sh.setRowHeight(r, 22);
  sh.setRowHeight(80, 6);
  // Section 6: VO + LAD Risk (81-94)
  sh.setRowHeight(81, 30); sh.setRowHeight(82, 300); sh.setRowHeight(83, 300);
  sh.setRowHeight(84, 24);
  for (var r = 85; r <= 89; r++)sh.setRowHeight(r, 22);
  sh.setRowHeight(90, 24);
  for (var r = 91; r <= 95; r++)sh.setRowHeight(r, 22);
  sh.setRowHeight(96, 6);
  // Section 7: Safety + Inspections (97-110)
  sh.setRowHeight(97, 30); sh.setRowHeight(98, 300); sh.setRowHeight(99, 300);
  sh.setRowHeight(100, 24);
  for (var r = 101; r <= 106; r++)sh.setRowHeight(r, 22);
  sh.setRowHeight(107, 24);
  for (var r = 108; r <= 112; r++)sh.setRowHeight(r, 22);
  sh.setRowHeight(113, 6);
  // Section 8: Permits + Docs + Retention (114-132)
  sh.setRowHeight(114, 30); sh.setRowHeight(115, 24);
  for (var r = 116; r <= 121; r++)sh.setRowHeight(r, 22);
  sh.setRowHeight(122, 24);
  for (var r = 123; r <= 127; r++)sh.setRowHeight(r, 22);
  sh.setRowHeight(128, 24);
  for (var r = 129; r <= 133; r++)sh.setRowHeight(r, 22);
  sh.setRowHeight(134, 6);
  // Section 9: PO Pipeline + Decision Matrix (135-145)
  sh.setRowHeight(135, 30); sh.setRowHeight(136, 24);
  for (var r = 137; r <= 141; r++)sh.setRowHeight(r, 22);
  sh.setRowHeight(142, 24);
  for (var r = 143; r <= 152; r++) { if (r <= sh.getMaxRows()) sh.setRowHeight(r, 22); }

  sh.getRange(1, 1, 155, 17).setBackground(THEME.bg);
  sh.getRange(1, 1, 155, 1).setBackground(THEME.sidebarBg);
  sh.setFrozenColumns(1); sh.setFrozenRows(4);
}

/* ================================================================ */
/*  SIDEBAR                                                         */
/* ================================================================ */
function buildSidebar_(sh) {
  [{ r: 2, i: '\u2302', a: true }, { r: 4, i: '\u2261' }, { r: 6, i: '\u2606' },
  { r: 8, i: '\u2692' }, { r: 10, i: '\u0024' }, { r: 12, i: '\u2691' },
  { r: 14, i: '\u2709' }, { r: 16, i: '\u2699' }].forEach(function (it) {
    var c = sh.getRange(it.r, 1);
    c.setValue(it.i).setFontFamily(THEME.font).setFontSize(14)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFontColor(it.a ? THEME.sidebarActive : THEME.sidebarText);
    if (it.a) c.setBackground('#1E293B').setFontWeight('bold');
  });
}

/* ================================================================ */
/*  HEADER + ALERT BAR                                              */
/* ================================================================ */
function buildHeader_(sh) {
  sh.getRange('C2:H2').merge().setValue('DUBAI ERP \u2014 OPERATIONS COMMAND CENTER')
    .setFontFamily(THEME.font).setFontSize(16).setFontWeight('bold')
    .setFontColor(THEME.headerText).setVerticalAlignment('middle').setBackground(THEME.bg);
  sh.getRange('C3:H3').merge()
    .setValue('Projects \u2022 Contractors \u2022 Budget \u2022 Equipment \u2022 Safety \u2022 Permits \u2022 Procurement \u2022 Quality')
    .setFontFamily(THEME.font).setFontSize(8).setFontColor(THEME.subText)
    .setVerticalAlignment('middle').setBackground(THEME.bg);

  // Alert bar
  var proj = rd_('Projects', 16); var now = new Date(); var rec = [];
  proj.forEach(function (r) {
    var sd = r[9];
    if (sd instanceof Date && db_(sd, now) <= 60 && db_(sd, now) >= 0) rec.push(String(r[1]).trim());
  });
  var atxt = rec.length > 0
    ? '\u26A0 ' + rec.length + ' NEW PROJECT(S) in 60 days: ' + rec.slice(0, 2).join(', ') + (rec.length > 2 ? ' +more' : '')
    : '\u2713 No new projects in last 60 days';
  sh.getRange('I2:Q2').merge().setValue(atxt)
    .setFontFamily(THEME.font).setFontSize(9).setFontWeight('bold')
    .setVerticalAlignment('middle').setHorizontalAlignment('right').setBackground(THEME.bg)
    .setFontColor(rec.length > 0 ? THEME.warning : THEME.success);
  sh.getRange('N3:Q3').merge().setFormula('="Updated: "&TEXT(NOW(),"DD-MMM-YYYY HH:MM")')
    .setFontFamily(THEME.font).setFontSize(8).setFontColor(THEME.subText)
    .setHorizontalAlignment('right').setVerticalAlignment('middle').setBackground(THEME.bg);
}

/* ================================================================ */
/*  FILTER PANEL — with checkbox reset                              */
/* ================================================================ */
function buildFilterPanel_(sh) {
  sh.getRange('C4:Q4').setBackground(THEME.filterBar);
  var s = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  // PROJECT filter
  sh.getRange('C4').setValue('PROJECT').setFontFamily(THEME.font).setFontSize(8)
    .setFontWeight('bold').setFontColor('#FFF').setHorizontalAlignment('right')
    .setVerticalAlignment('middle').setBackground(THEME.filterBar);
  sh.getRange('D4:E4').merge().setValue('All Projects').setFontFamily(THEME.font)
    .setFontSize(9).setFontWeight('bold').setFontColor(THEME.headerText)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#FFF').setBorder(true, true, true, true, false, false, THEME.accent, s);
  // TYPE filter
  sh.getRange('G4').setValue('TYPE').setFontFamily(THEME.font).setFontSize(8)
    .setFontWeight('bold').setFontColor('#FFF').setHorizontalAlignment('right')
    .setVerticalAlignment('middle').setBackground(THEME.filterBar);
  sh.getRange('H4:I4').merge().setValue('All Types').setFontFamily(THEME.font)
    .setFontSize(9).setFontWeight('bold').setFontColor(THEME.headerText)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#FFF').setBorder(true, true, true, true, false, false, THEME.accent, s);
  // STATUS filter
  sh.getRange('K4').setValue('STATUS').setFontFamily(THEME.font).setFontSize(8)
    .setFontWeight('bold').setFontColor('#FFF').setHorizontalAlignment('right')
    .setVerticalAlignment('middle').setBackground(THEME.filterBar);
  sh.getRange('L4:M4').merge().setValue('All Statuses').setFontFamily(THEME.font)
    .setFontSize(9).setFontWeight('bold').setFontColor(THEME.headerText)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#FFF').setBorder(true, true, true, true, false, false, THEME.accent, s);
  // RESET checkbox
  sh.getRange('O4').setValue('RESET').setFontFamily(THEME.font).setFontSize(8)
    .setFontWeight('bold').setFontColor('#FFF').setHorizontalAlignment('right')
    .setVerticalAlignment('middle').setBackground(THEME.filterBar);
  sh.getRange('P4').insertCheckboxes().setValue(false)
    .setBackground('#FFF').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange('Q4').setValue('\u21BB Clear').setFontFamily(THEME.font).setFontSize(8)
    .setFontColor('#FFF').setFontStyle('italic').setHorizontalAlignment('left')
    .setVerticalAlignment('middle').setBackground(THEME.filterBar);
}

function setupFilterValidation_(sh) {
  var proj = rd_('Projects', 16);
  var pIds = ['All Projects'].concat(uq_(proj, 0));
  sh.getRange('D4').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(pIds, true).setAllowInvalid(false).build());
  var tList = ['All Types'].concat(uq_(proj, 2));
  sh.getRange('H4').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(tList, true).setAllowInvalid(false).build());
  sh.getRange('L4').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['All Statuses', 'Completed', 'In Progress', 'Delayed', 'On Hold'], true)
    .setAllowInvalid(false).build());
}

/* ================================================================ */
/*  KPI ROW 1 — Portfolio Health (4 cards)                          */
/* ================================================================ */
function buildKPIRow1_(sh) { kpiLabel_(sh, 3, 6, 'Active Projects', THEME.success); kpiLabel_(sh, 7, 6, 'Delayed Projects', THEME.danger); kpiLabel_(sh, 11, 6, 'Total Contract Value (M)', THEME.purple); kpiLabel_(sh, 15, 6, 'Portfolio Completion %', THEME.info); }
function buildKPIRow2_(sh) { kpiLabel_(sh, 3, 10, 'Budget Overrun %', THEME.danger); kpiLabel_(sh, 7, 10, 'Pending Payments (M)', THEME.orange); kpiLabel_(sh, 11, 10, 'Approved VOs (M)', THEME.warning); kpiLabel_(sh, 15, 10, 'Cost Performance Index', THEME.info); }
function buildKPIRow3_(sh) { kpiLabel_(sh, 3, 14, 'Open Safety Incidents', THEME.pink); kpiLabel_(sh, 7, 14, 'Inspection Pass Rate %', THEME.teal); kpiLabel_(sh, 11, 14, 'Expiring Permits (<30d)', THEME.orange); kpiLabel_(sh, 15, 14, 'Rented Equip. Cost/Day', THEME.accent); }

/* KPI label only — value set by computeKPIs_ */
function kpiLabel_(sh, col, row, title, color) {
  var t = sh.getRange(row, col, 1, 3);
  t.merge().setValue(title).setFontFamily(THEME.font).setFontSize(8)
    .setFontColor(THEME.subText).setHorizontalAlignment('center')
    .setVerticalAlignment('bottom').setBackground(THEME.cardBg);
  t.setBorder(true, true, null, true, false, false, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, col, 1, 3).setBorder(true, null, null, null, null, null, color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  var v = sh.getRange(row + 1, col, 1, 3);
  v.merge().setValue('—').setFontFamily(THEME.font)
    .setFontSize(18).setFontWeight('bold').setFontColor(THEME.headerText)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(THEME.cardBg);
  v.setBorder(null, true, null, true, false, false, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
  var b = sh.getRange(row + 2, col, 1, 3);
  b.merge().setBackground(THEME.cardBg).setHorizontalAlignment('center').setVerticalAlignment('top');
  b.setBorder(null, true, true, true, false, false, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
  b.setValue(' ').setFontSize(6).setFontColor(THEME.subText);
}

/* ================================================================ */
/*  COMPUTE KPIs — JS-computed, filter-aware, called on every edit  */
/* ================================================================ */
function computeKPIs_(sh) {
  var P = rd_('Projects', 16), fP = getFiltered_(P), pids = getFilteredPids_();
  var CT = rd_('Contracts', 15), WP = rd_('Work_Packages', 9), VO = rd_('Variation_Orders', 11);
  var PA = rd_('Payment_Applications', 13), SI = rd_('Safety_Incidents', 12);
  var INS = rd_('Inspections', 11), PM = rd_('Permits_Approvals', 10), EQ = rd_('Equipment', 10);

  // ROW 1
  var active = 0, delayed = 0, totalCV = 0, compW = 0, compWt = 0;
  fP.forEach(function (r) {
    var st = String(r[11]).trim(), cv = Number(r[8]) || 0, comp = Number(r[12]) || 0;
    if (st === 'In Progress') active++;
    if (st === 'Delayed') delayed++;
    compW += comp * cv; compWt += cv;
  });
  CT.forEach(function (r) { if (pids[String(r[1]).trim()]) totalCV += (Number(r[5]) || 0); });
  setKV_(sh, 7, 3, active, '#,##0'); setKV_(sh, 7, 7, delayed, '#,##0');
  setKV_(sh, 7, 11, Math.round(totalCV / 1e5) / 10, '#,##0.0'); setKV_(sh, 7, 15, compWt ? Math.round(compW / compWt * 10) / 10 : 0, '#,##0.0');

  // ROW 2
  var wpPl = 0, wpAc = 0, pendPay = 0, voApp = 0, certPay = 0;
  WP.forEach(function (r) { if (pids[String(r[1]).trim()]) { wpPl += (Number(r[6]) || 0); wpAc += (Number(r[8]) || 0); } });
  PA.forEach(function (r) { if (!pids[String(r[1]).trim()]) return; var s = String(r[11]).trim(); if (s === 'Submitted') pendPay += (Number(r[6]) || 0); if (s === 'Certified' || s === 'Paid') certPay += (Number(r[10]) || 0); });
  VO.forEach(function (r) { if (pids[String(r[1]).trim()] && String(r[8]).trim() === 'Approved') voApp += (Number(r[6]) || 0); });
  var overrun = wpPl ? Math.round((wpAc / wpPl - 1) * 1000) / 10 : 0;
  var cpi = totalCV ? Math.round(certPay / totalCV * 100) / 100 : 0;
  setKV_(sh, 11, 3, overrun, '#,##0.0'); setKV_(sh, 11, 7, Math.round(pendPay / 1e5) / 10, '#,##0.0');
  setKV_(sh, 11, 11, Math.round(voApp / 1e5) / 10, '#,##0.0'); setKV_(sh, 11, 15, cpi, '0.00');

  // ROW 3
  var openSI = 0, insPass = 0, insTotal = 0, expPerm = 0, rentCost = 0, now = new Date();
  SI.forEach(function (r) { if (pids[String(r[1]).trim()] && String(r[11]).trim() !== 'Closed') openSI++; });
  INS.forEach(function (r) { if (!pids[String(r[1]).trim()]) return; var rs = String(r[6]).trim(); if (rs) insTotal++; if (rs === 'Passed') insPass++; });
  PM.forEach(function (r) { if (!pids[String(r[1]).trim()]) return; var exp = r[6]; if (exp instanceof Date) { var dl = db_(now, exp); if (dl >= 0 && dl <= 30) expPerm++; } });
  EQ.forEach(function (r) { if (String(r[4]).trim() === 'Rented') { var proj = String(r[7]).trim(); if (!proj || pids[proj]) rentCost += (Number(r[5]) || 0); } });
  setKV_(sh, 15, 3, openSI, '#,##0'); setKV_(sh, 15, 7, insTotal ? Math.round(insPass / insTotal * 1000) / 10 : 0, '#,##0.0');
  setKV_(sh, 15, 11, expPerm, '#,##0'); setKV_(sh, 15, 15, rentCost, '#,##0');

  // Color KPI values based on thresholds
  if (overrun > 10) sh.getRange(11, 3, 1, 3).setFontColor(THEME.danger); else if (overrun > 0) sh.getRange(11, 3, 1, 3).setFontColor(THEME.warning); else sh.getRange(11, 3, 1, 3).setFontColor(THEME.success);
  if (cpi < 0.8) sh.getRange(11, 15, 1, 3).setFontColor(THEME.danger); else if (cpi < 1.0) sh.getRange(11, 15, 1, 3).setFontColor(THEME.warning);
  if (openSI > 10) sh.getRange(15, 3, 1, 3).setFontColor(THEME.danger);
}

function setKV_(sh, r, c, v, fmt) {
  var rng = sh.getRange(r, c, 1, 3);
  rng.merge().setValue(v).setNumberFormat(fmt).setFontFamily(THEME.font)
    .setFontSize(18).setFontWeight('bold').setFontColor(THEME.headerText)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(THEME.cardBg);
  rng.setBorder(null, true, null, true, false, false, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
}

/* ================================================================ */
/*  SECTION HEADERS                                                 */
/* ================================================================ */
function buildSections_(sh) {
  sec_(sh, 18, '\uD83D\uDCCB  NEW & RECENT PROJECTS + PHASE PROGRESS');
  sec_(sh, 35, '\uD83D\uDCCA  PROJECT STATUS OVERVIEW & BENCHMARKS');
  sec_(sh, 39, '\uD83C\uDFD7\uFE0F  CONTRACTOR MANAGEMENT & RECOMMENDATIONS');
  sec_(sh, 51, '\uD83D\uDCB0  BUDGET CONTROL & CASH FLOW INTELLIGENCE');
  sec_(sh, 70, '\u2699\uFE0F  EQUIPMENT RENTAL & UTILIZATION MONITOR');
  sec_(sh, 81, '\uD83D\uDCDD  VARIATION ORDERS & LAD RISK ANALYSIS');
  sec_(sh, 97, '\uD83D\uDEE1\uFE0F  SAFETY & QUALITY INTELLIGENCE');
  sec_(sh, 114, '\uD83D\uDCDC  PERMITS, DOCUMENTS & RETENTION TRACKER');
  sec_(sh, 135, '\uD83D\uDCCA  PROCUREMENT PIPELINE & DECISION MATRIX');
}

/* ================================================================ */
/*  DATA ENGINE — 25+ computed data blocks                          */
/* ================================================================ */
function buildDataEngine_(sh) {
  var P = rd_('Projects', 16), CON = rd_('Contractors', 12), CT = rd_('Contracts', 15);
  var PH = rd_('Project_Phases', 9), WP = rd_('Work_Packages', 9), VO = rd_('Variation_Orders', 11);
  var PO = rd_('Purchase_Orders', 10), EQ = rd_('Equipment', 10), SI = rd_('Safety_Incidents', 12);
  var INS = rd_('Inspections', 11), PA = rd_('Payment_Applications', 13);
  var PM = rd_('Permits_Approvals', 10), CL = rd_('Clients', 12), DSR = rd_('Daily_Site_Reports', 12);
  var SUP = rd_('Suppliers', 12), DOC = rd_('Project_Documents', 11);

  var rM = { 'A': 4, 'B': 3, 'C': 2, 'D': 1 };
  var types = uq_(P, 2), pMap = {};
  P.forEach(function (r) { pMap[String(r[0]).trim()] = String(r[2]).trim(); });

  var col = 27; // start column for hidden data

  // B1: AA-AB Contractor Rating by Specialty
  var sp = uq_(CON, 2), spS = {}, spC = {};
  CON.forEach(function (r) { var s = String(r[2]).trim(), rt = rM[String(r[10]).trim()] || 0; if (s) { spS[s] = (spS[s] || 0) + rt; spC[s] = (spC[s] || 0) + 1; } });
  sp.sort(function (a, b) { return (spS[b] / spC[b]) - (spS[a] / spC[a]); });
  sh.getRange('AA1').setValue('Specialty'); sh.getRange('AB1').setValue('Avg Rating');
  for (var i = 0; i < sp.length; i++) { sh.getRange('AA' + (i + 2)).setValue(sp[i]); sh.getRange('AB' + (i + 2)).setValue(Math.round(spS[sp[i]] / spC[sp[i]] * 10) / 10); }

  // B2: AC-AD Top 10 Contractors by Contract Count
  var cC = {}, cN = {};
  CT.forEach(function (r) { var c = String(r[3]).trim(); if (c) cC[c] = (cC[c] || 0) + 1; });
  CON.forEach(function (r) { cN[String(r[0]).trim()] = String(r[1]).trim(); });
  var tC = Object.keys(cC).sort(function (a, b) { return cC[b] - cC[a]; }).slice(0, 10);
  sh.getRange('AC1').setValue('Contractor'); sh.getRange('AD1').setValue('Contracts');
  for (var i = 0; i < tC.length; i++) { sh.getRange('AC' + (i + 2)).setValue(cN[tC[i]] || tC[i]); sh.getRange('AD' + (i + 2)).setValue(cC[tC[i]]); }

  // B3: AE-AF Budget by Phase
  var pN = uq_(PH, 2), pB = {};
  PH.forEach(function (r) { var p = String(r[2]).trim(), b = Number(r[8]) || 0; if (p) pB[p] = (pB[p] || 0) + b; });
  pN.sort(function (a, b) { return (pB[b] || 0) - (pB[a] || 0); });
  sh.getRange('AE1').setValue('Phase'); sh.getRange('AF1').setValue('Budget');
  for (var i = 0; i < pN.length; i++) { sh.getRange('AE' + (i + 2)).setValue(pN[i]); sh.getRange('AF' + (i + 2)).setValue(pB[pN[i]] || 0); }

  // B4: AG-AI Budget vs Actual by Type
  var tPl = {}, tAc = {};
  WP.forEach(function (r) { var t = pMap[String(r[1]).trim()]; if (t) { tPl[t] = (tPl[t] || 0) + (Number(r[6]) || 0); tAc[t] = (tAc[t] || 0) + (Number(r[8]) || 0); } });
  sh.getRange('AG1').setValue('Type'); sh.getRange('AH1').setValue('Planned'); sh.getRange('AI1').setValue('Actual');
  for (var i = 0; i < types.length; i++) { sh.getRange('AG' + (i + 2)).setValue(types[i]); sh.getRange('AH' + (i + 2)).setValue(tPl[types[i]] || 0); sh.getRange('AI' + (i + 2)).setValue(tAc[types[i]] || 0); }

  // B5: AJ-AK VO Risk by Type
  var vS = {}, vC = {};
  VO.forEach(function (r) { var t = pMap[String(r[1]).trim()], v = Number(r[6]) || 0; if (t && v > 0) { vS[t] = (vS[t] || 0) + v; vC[t] = (vC[t] || 0) + 1; } });
  sh.getRange('AJ1').setValue('Type'); sh.getRange('AK1').setValue('Avg VO Value');
  for (var i = 0; i < types.length; i++) { sh.getRange('AJ' + (i + 2)).setValue(types[i]); sh.getRange('AK' + (i + 2)).setValue(vC[types[i]] ? Math.round(vS[types[i]] / vC[types[i]]) : 0); }

  // B6: AL-AM PO Spend Monthly
  var mS = {}; PO.forEach(function (r) { var d = r[5]; if (!(d instanceof Date)) return; mS[d.getMonth() + 1] = (mS[d.getMonth() + 1] || 0) + (Number(r[7]) || 0); });
  sh.getRange('AL1').setValue('Month'); sh.getRange('AM1').setValue('PO Spend');
  for (var m = 1; m <= 12; m++) { sh.getRange('AL' + (m + 1)).setValue(m); sh.getRange('AM' + (m + 1)).setValue(mS[m] || 0); }

  // B7: AN-AO Fleet Split
  var oC = {}; EQ.forEach(function (r) { var o = String(r[4]).trim(); if (o) oC[o] = (oC[o] || 0) + 1; });
  var oK = Object.keys(oC).sort();
  sh.getRange('AN1').setValue('Ownership'); sh.getRange('AO1').setValue('Count');
  for (var i = 0; i < oK.length; i++) { sh.getRange('AN' + (i + 2)).setValue(oK[i]); sh.getRange('AO' + (i + 2)).setValue(oC[oK[i]]); }

  // B8: AP-AR Equipment Rate by Type
  var eT = uq_(EQ, 1), rR = {}, rC2 = {}, oR = {}, oC2 = {};
  EQ.forEach(function (r) {
    var t = String(r[1]).trim(), o = String(r[4]).trim(), rt = Number(r[5]) || 0; if (!t) return;
    if (o === 'Rented') { rR[t] = (rR[t] || 0) + rt; rC2[t] = (rC2[t] || 0) + 1; } if (o === 'Owned') { oR[t] = (oR[t] || 0) + rt; oC2[t] = (oC2[t] || 0) + 1; }
  });
  sh.getRange('AP1').setValue('Equipment'); sh.getRange('AQ1').setValue('Rented'); sh.getRange('AR1').setValue('Owned');
  var mE = Math.min(eT.length, 12);
  for (var i = 0; i < mE; i++) { sh.getRange('AP' + (i + 2)).setValue(eT[i]); sh.getRange('AQ' + (i + 2)).setValue(rC2[eT[i]] ? Math.round(rR[eT[i]] / rC2[eT[i]]) : 0); sh.getRange('AR' + (i + 2)).setValue(oC2[eT[i]] ? Math.round(oR[eT[i]] / oC2[eT[i]]) : 0); }

  // B9: AS-AT Utilization
  var rT = 0, rU = 0, oT = 0, oU = 0;
  EQ.forEach(function (r) { var o = String(r[4]).trim(), s = String(r[6]).trim(); if (o === 'Rented') { rT++; if (s === 'In Use') rU++; } if (o === 'Owned') { oT++; if (s === 'In Use') oU++; } });
  sh.getRange('AS1').setValue('Type'); sh.getRange('AT1').setValue('Util %');
  sh.getRange('AS2').setValue('Rented'); sh.getRange('AT2').setValue(rT ? Math.round(rU / rT * 1e3) / 10 : 0);
  sh.getRange('AS3').setValue('Owned'); sh.getRange('AT3').setValue(oT ? Math.round(oU / oT * 1e3) / 10 : 0);

  // B10: AU-AW Duration by Type
  var dP = {}, dA = {}, dPC = {}, dAC = {};
  P.forEach(function (r) { var t = String(r[2]).trim(), s = r[9], e = r[10], st = String(r[11]).trim(); if (!t || !(s instanceof Date) || !(e instanceof Date)) return; var d = Math.round((e - s) / 864e5); dP[t] = (dP[t] || 0) + d; dPC[t] = (dPC[t] || 0) + 1; if (st === 'Completed') { dA[t] = (dA[t] || 0) + d; dAC[t] = (dAC[t] || 0) + 1; } });
  sh.getRange('AU1').setValue('Type'); sh.getRange('AV1').setValue('Plan Days'); sh.getRange('AW1').setValue('Act Days');
  for (var i = 0; i < types.length; i++) { sh.getRange('AU' + (i + 2)).setValue(types[i]); sh.getRange('AV' + (i + 2)).setValue(dPC[types[i]] ? Math.round(dP[types[i]] / dPC[types[i]]) : 0); sh.getRange('AW' + (i + 2)).setValue(dAC[types[i]] ? Math.round(dA[types[i]] / dAC[types[i]]) : 0); }

  // B11: AX-AY CPI by Type
  var tCr = {}, tCt = {};
  PA.forEach(function (r) { var t = pMap[String(r[1]).trim()], s = String(r[11]).trim(), v = Number(r[10]) || 0; if (t && (s === 'Certified' || s === 'Paid')) tCr[t] = (tCr[t] || 0) + v; });
  CT.forEach(function (r) { var t = pMap[String(r[1]).trim()], v = Number(r[5]) || 0; if (t && v > 0) tCt[t] = (tCt[t] || 0) + v; });
  sh.getRange('AX1').setValue('Type'); sh.getRange('AY1').setValue('CPI');
  for (var i = 0; i < types.length; i++) { sh.getRange('AX' + (i + 2)).setValue(types[i]); sh.getRange('AY' + (i + 2)).setValue(tCt[types[i]] ? Math.round((tCr[types[i]] || 0) / tCt[types[i]] * 100) / 100 : 0); }

  // B12: AZ-BA Safety by Type
  var tSf = {}; SI.forEach(function (r) { var t = pMap[String(r[1]).trim()]; if (t) tSf[t] = (tSf[t] || 0) + 1; });
  sh.getRange('AZ1').setValue('Type'); sh.getRange('BA1').setValue('Incidents');
  for (var i = 0; i < types.length; i++) { sh.getRange('AZ' + (i + 2)).setValue(types[i]); sh.getRange('BA' + (i + 2)).setValue(tSf[types[i]] || 0); }

  // B13: BB-BC Inspection Pass Rate
  var tPs = {}, tTo = {};
  INS.forEach(function (r) { var t = pMap[String(r[1]).trim()], rs = String(r[6]).trim(); if (t && rs) { tTo[t] = (tTo[t] || 0) + 1; if (rs === 'Passed') tPs[t] = (tPs[t] || 0) + 1; } });
  sh.getRange('BB1').setValue('Type'); sh.getRange('BC1').setValue('Pass %');
  for (var i = 0; i < types.length; i++) { sh.getRange('BB' + (i + 2)).setValue(types[i]); sh.getRange('BC' + (i + 2)).setValue(tTo[types[i]] ? Math.round((tPs[types[i]] || 0) / tTo[types[i]] * 1e3) / 10 : 0); }

  // B14: BD-BE Project Status Distribution
  var sC = {}; P.forEach(function (r) { var s = String(r[11]).trim(); if (s) sC[s] = (sC[s] || 0) + 1; });
  var sK = Object.keys(sC).sort();
  sh.getRange('BD1').setValue('Status'); sh.getRange('BE1').setValue('Count');
  for (var i = 0; i < sK.length; i++) { sh.getRange('BD' + (i + 2)).setValue(sK[i]); sh.getRange('BE' + (i + 2)).setValue(sC[sK[i]]); }

  // B15: BF-BG Project Completion Top 15
  var pComp = []; P.forEach(function (r) { pComp.push({ n: String(r[1]).trim().substring(0, 22), p: Number(r[12]) || 0 }); });
  pComp.sort(function (a, b) { return b.p - a.p; }); var p15 = pComp.slice(0, 15);
  sh.getRange('BF1').setValue('Project'); sh.getRange('BG1').setValue('Completion %');
  for (var i = 0; i < p15.length; i++) { sh.getRange('BF' + (i + 2)).setValue(p15[i].n); sh.getRange('BG' + (i + 2)).setValue(p15[i].p); }

  // B16: BH-BI VO Status Distribution
  var vSt = {}; VO.forEach(function (r) { var s = String(r[8]).trim(); if (s) vSt[s] = (vSt[s] || 0) + 1; });
  var vK = Object.keys(vSt).sort();
  sh.getRange('BH1').setValue('VO Status'); sh.getRange('BI1').setValue('Count');
  for (var i = 0; i < vK.length; i++) { sh.getRange('BH' + (i + 2)).setValue(vK[i]); sh.getRange('BI' + (i + 2)).setValue(vSt[vK[i]]); }

  // B17: BJ-BK Safety by Severity
  var svC = {}; SI.forEach(function (r) { var s = String(r[3]).trim(); if (s) svC[s] = (svC[s] || 0) + 1; });
  var svK = Object.keys(svC).sort();
  sh.getRange('BJ1').setValue('Severity'); sh.getRange('BK1').setValue('Count');
  for (var i = 0; i < svK.length; i++) { sh.getRange('BJ' + (i + 2)).setValue(svK[i]); sh.getRange('BK' + (i + 2)).setValue(svC[svK[i]]); }

  // B18: BL-BM Permit Status
  var pmS = {}; PM.forEach(function (r) { var s = String(r[7]).trim(); if (s) pmS[s] = (pmS[s] || 0) + 1; });
  var pmK = Object.keys(pmS).sort();
  sh.getRange('BL1').setValue('Permit'); sh.getRange('BM1').setValue('Count');
  for (var i = 0; i < pmK.length; i++) { sh.getRange('BL' + (i + 2)).setValue(pmK[i]); sh.getRange('BM' + (i + 2)).setValue(pmS[pmK[i]]); }

  // B19: BN-BO Client Portfolio
  var clN = {}; CL.forEach(function (r) { clN[String(r[0]).trim()] = String(r[1]).trim(); });
  var clV = {}; P.forEach(function (r) { var c = String(r[3]).trim(), v = Number(r[8]) || 0; if (c && v > 0) clV[c] = (clV[c] || 0) + v; });
  var tCl = Object.keys(clV).sort(function (a, b) { return clV[b] - clV[a]; }).slice(0, 10);
  sh.getRange('BN1').setValue('Client'); sh.getRange('BO1').setValue('Portfolio');
  for (var i = 0; i < tCl.length; i++) { sh.getRange('BN' + (i + 2)).setValue(clN[tCl[i]] || tCl[i]); sh.getRange('BO' + (i + 2)).setValue(clV[tCl[i]]); }

  // B20: BP-BR Manpower Monthly
  var mD = {}, mSb = {}; DSR.forEach(function (r) { var d = r[2]; if (!(d instanceof Date)) return; var m = d.getMonth() + 1; mD[m] = (mD[m] || 0) + (Number(r[6]) || 0); mSb[m] = (mSb[m] || 0) + (Number(r[7]) || 0); });
  sh.getRange('BP1').setValue('Month'); sh.getRange('BQ1').setValue('Direct'); sh.getRange('BR1').setValue('Subcon');
  for (var m = 1; m <= 12; m++) { sh.getRange('BP' + (m + 1)).setValue(m); sh.getRange('BQ' + (m + 1)).setValue(mD[m] || 0); sh.getRange('BR' + (m + 1)).setValue(mSb[m] || 0); }

  // B21: BS-BT Cost per SQM by Project
  var cSqm = []; P.forEach(function (r) { var gfa = Number(r[13]) || 0, cv = Number(r[8]) || 0; if (gfa > 0) cSqm.push({ n: String(r[1]).trim().substring(0, 18), v: Math.round(cv / gfa) }); });
  cSqm.sort(function (a, b) { return b.v - a.v; }); var cS10 = cSqm.slice(0, 10);
  sh.getRange('BS1').setValue('Project'); sh.getRange('BT1').setValue('AED/SQM');
  for (var i = 0; i < cS10.length; i++) { sh.getRange('BS' + (i + 2)).setValue(cS10[i].n); sh.getRange('BT' + (i + 2)).setValue(cS10[i].v); }

  // B22: BU-BV Payment Status Pipeline
  var paS = {}; PA.forEach(function (r) { var s = String(r[11]).trim(); if (s) paS[s] = (paS[s] || 0) + 1; });
  var paK = Object.keys(paS).sort();
  sh.getRange('BU1').setValue('Pay Status'); sh.getRange('BV1').setValue('Count');
  for (var i = 0; i < paK.length; i++) { sh.getRange('BU' + (i + 2)).setValue(paK[i]); sh.getRange('BV' + (i + 2)).setValue(paS[paK[i]]); }

  // B23: BW-BX PO Status Pipeline
  var poS = {}; PO.forEach(function (r) { var s = String(r[8]).trim(); if (s) poS[s] = (poS[s] || 0) + 1; });
  var poK = Object.keys(poS).sort();
  sh.getRange('BW1').setValue('PO Status'); sh.getRange('BX1').setValue('Count');
  for (var i = 0; i < poK.length; i++) { sh.getRange('BW' + (i + 2)).setValue(poK[i]); sh.getRange('BX' + (i + 2)).setValue(poS[poK[i]]); }

  // B24: BY-BZ Document Status
  var dcS = {}; DOC.forEach(function (r) { var s = String(r[8]).trim(); if (s) dcS[s] = (dcS[s] || 0) + 1; });
  var dcK = Object.keys(dcS).sort();
  sh.getRange('BY1').setValue('Doc Status'); sh.getRange('BZ1').setValue('Count');
  for (var i = 0; i < dcK.length; i++) { sh.getRange('BY' + (i + 2)).setValue(dcK[i]); sh.getRange('BZ' + (i + 2)).setValue(dcS[dcK[i]]); }

  // B25: CA-CB Supplier Category Distribution
  var suC = {}; SUP.forEach(function (r) { var c = String(r[2]).trim(); if (c) suC[c] = (suC[c] || 0) + 1; });
  var suK = Object.keys(suC).sort();
  sh.getRange('CA1').setValue('Category'); sh.getRange('CB1').setValue('Suppliers');
  for (var i = 0; i < Math.min(suK.length, 10); i++) { sh.getRange('CA' + (i + 2)).setValue(suK[i]); sh.getRange('CB' + (i + 2)).setValue(suC[suK[i]]); }

  return {
    spN: sp.length, cN: tC.length, phN: pN.length, tyN: types.length, eqN: mE, owN: oK.length,
    stN: sK.length, p15: p15.length, voN: vK.length, svN: svK.length, pmN: pmK.length, clN: tCl.length,
    csN: cS10.length, paN: paK.length, poN: poK.length, dcN: dcK.length, suN: Math.min(suK.length, 10)
  };
}

/* ================================================================ */
/*  ALL CHARTS — 20 charts across 9 sections                       */
/* ================================================================ */
function buildAllCharts_(sh, m) {
  var tr = m.tyN + 1;
  // Section 2: Project Overview (rows 36-37)
  ch_(sh, 'PIE', 'BD1:BE' + (m.stN + 1), 36, 3, 'Project Status Distribution', { hole: 0.55, leg: 'right', col: [THEME.success, THEME.warning, THEME.danger, THEME.info] });
  ch_(sh, 'BAR', 'BF1:BG' + (m.p15 + 1), 36, 11, 'Project Completion Progress (%)', { col: [THEME.accent] });
  ch_(sh, 'BAR', 'BS1:BT' + (m.csN + 1), 37, 3, 'Cost per SQM by Project (AED)', { col: [THEME.teal] });
  ch_(sh, 'BAR', 'BN1:BO' + (m.clN + 1), 37, 11, 'Client Portfolio Value (AED)', { col: [THEME.purple] });
  // Section 3: Contractor (rows 40-41)
  ch_(sh, 'BAR', 'AA1:AB' + (m.spN + 1), 40, 3, 'Contractor Rating by Specialty', { col: [THEME.accent] });
  ch_(sh, 'COLUMN', 'AC1:AD' + (m.cN + 1), 40, 11, 'Top 10 Contractors (Contracts)', { col: [THEME.success] });
  ch_(sh, 'AREA', 'BP1:BR13', 41, 3, 'Manpower Trend (Direct+Subcon)', { col: [THEME.info, THEME.orange], leg: 'top' });
  ch_(sh, 'COLUMN', 'CA1:CB' + (m.suN + 1), 41, 11, 'Supplier Category Distribution', { col: THEME.palette });
  // Section 4: Budget (rows 52-53)
  ch_(sh, 'BAR', 'AE1:AF' + (m.phN + 1), 52, 3, 'Budget Allocation by Phase', { col: [THEME.purple] });
  ch_(sh, 'COLUMN', 'AG1:AI' + tr, 52, 11, 'Budget vs Actual by Type', { col: [THEME.info, THEME.danger], leg: 'top' });
  ch_(sh, 'AREA', 'AL1:AM13', 53, 3, 'PO Monthly Spend Trend', { col: [THEME.teal] });
  ch_(sh, 'PIE', 'BU1:BV' + (m.paN + 1), 53, 11, 'Payment Status Pipeline', { hole: 0.55, leg: 'right', col: [THEME.success, THEME.warning, THEME.danger, THEME.info] });
  // Section 5: Equipment (rows 71-72)
  ch_(sh, 'PIE', 'AN1:AO' + (m.owN + 1), 71, 3, 'Fleet Ownership Split', { hole: 0.55, col: [THEME.success, THEME.warning], leg: 'right' });
  ch_(sh, 'BAR', 'AP1:AR' + (m.eqN + 1), 71, 11, 'Daily Rate: Rented vs Owned', { col: [THEME.danger, THEME.info], leg: 'top' });
  ch_(sh, 'COLUMN', 'AS1:AT3', 72, 3, 'Utilization: Rented vs Owned (%)', { col: [THEME.accent, THEME.teal] });
  ch_(sh, 'PIE', 'BY1:BZ' + (m.dcN + 1), 72, 11, 'Document Status Distribution', { hole: 0.55, leg: 'right', col: THEME.palette });
  // Section 6: VO (rows 82-83)
  ch_(sh, 'PIE', 'BH1:BI' + (m.voN + 1), 82, 3, 'VO Status Distribution', { hole: 0.55, leg: 'right', col: [THEME.success, THEME.warning, THEME.danger] });
  ch_(sh, 'BAR', 'AJ1:AK' + tr, 82, 11, 'VO Risk by Type (Avg Value)', { col: [THEME.warning] });
  // Section 7: Safety (rows 98-99)
  ch_(sh, 'PIE', 'BJ1:BK' + (m.svN + 1), 98, 3, 'Incidents by Severity', { hole: 0.55, leg: 'right', col: [THEME.danger, THEME.warning, THEME.info, THEME.success] });
  ch_(sh, 'BAR', 'BB1:BC' + tr, 98, 11, 'Inspection Pass Rate by Type', { col: [THEME.success] });
  ch_(sh, 'COLUMN', 'AU1:AW' + tr, 99, 3, 'Duration: Planned vs Actual', { col: [THEME.info, THEME.warning], leg: 'top' });
  ch_(sh, 'BAR', 'AX1:AY' + tr, 99, 11, 'Cost Performance Index by Type', { col: [THEME.purple] });
  ch_(sh, 'BAR', 'AZ1:BA' + tr, 83, 3, 'Safety Incidents by Type', { col: [THEME.danger] });
  ch_(sh, 'PIE', 'BW1:BX' + (m.poN + 1), 83, 11, 'PO Status Pipeline', { hole: 0.55, leg: 'right', col: THEME.palette });
}

/* ================================================================ */
/*  TABLE 1: NEW & RECENT PROJECTS (rows 19-26)                     */
/* ================================================================ */
function buildNewProjectsTable_(sh) {
  var hd = ['#', 'Project Name', 'Type', 'Client', 'Start Date', 'Value (AED)', 'Status', 'Age'];
  var sp = [[3, 3], [4, 6], [7, 8], [9, 10], [11, 12], [13, 14], [15, 15], [16, 17]];
  th_(sh, 19, sp, hd);
  var P = rd_('Projects', 16), CL = rd_('Clients', 12);
  var cn = {}; CL.forEach(function (r) { cn[String(r[0]).trim()] = String(r[1]).trim(); });
  var now = new Date();
  P.sort(function (a, b) { var da = a[9] instanceof Date ? a[9].getTime() : 0, db2 = b[9] instanceof Date ? b[9].getTime() : 0; return db2 - da; });
  var fP = getFiltered_(P);
  for (var i = 0; i < Math.min(fP.length, 7); i++) {
    var dr = 20 + i, r = fP[i], bg = (i % 2 === 0) ? THEME.cardBg : '#F1F5F9';
    var st = String(r[11]).trim(), age = r[9] instanceof Date ? db_(r[9], now) : 0;
    if (age <= 30) bg = THEME.gBg; if (st === 'Delayed') bg = THEME.rBg;
    var rk = sh.getRange(dr, 3); rk.setValue(i + 1); sc_(rk, bg, 'center'); rk.setFontWeight('bold').setFontColor(THEME.accent);
    var nm = sh.getRange(dr, 4, 1, 3); nm.merge().setValue(String(r[1]).trim().substring(0, 28)); sc_(nm, bg, 'left');
    var tp = sh.getRange(dr, 7, 1, 2); tp.merge().setValue(String(r[2]).trim()); sc_(tp, bg, 'center');
    var cl = sh.getRange(dr, 9, 1, 2); cl.merge().setValue((cn[String(r[3]).trim()] || '—').substring(0, 18)); sc_(cl, bg, 'center');
    var sd = sh.getRange(dr, 11, 1, 2); sd.merge().setValue(r[9] instanceof Date ? Utilities.formatDate(r[9], 'Asia/Dubai', 'dd-MMM-yy') : '—'); sc_(sd, bg, 'center');
    var vl = sh.getRange(dr, 13, 1, 2); vl.merge().setValue(fm_(Number(r[8]) || 0) + 'M'); sc_(vl, bg, 'center');
    var ss2 = sh.getRange(dr, 15); ss2.setValue(st); sc_(ss2, bg, 'center'); ss2.setFontWeight('bold');
    if (st === 'In Progress') ss2.setFontColor(THEME.success); else if (st === 'Delayed') ss2.setFontColor(THEME.danger); else if (st === 'Completed') ss2.setFontColor(THEME.info); else ss2.setFontColor(THEME.warning);
    var ag = sh.getRange(dr, 16, 1, 2); ag.merge().setValue(age <= 30 ? 'NEW (' + age + 'd)' : age <= 60 ? age + 'd' : age + 'd'); sc_(ag, bg, 'center'); ag.setFontWeight('bold');
    if (age <= 30) ag.setFontColor(THEME.success).setBackground(THEME.gBg); else if (age <= 90) ag.setFontColor(THEME.info); else ag.setFontColor(THEME.subText);
  }
}

/* ================================================================ */
/*  TABLE 2: PHASE PROGRESS TRACKER (rows 27-33)                    */
/* ================================================================ */
function buildPhaseProgressTable_(sh) {
  var hd = ['#', 'Project', 'Phase', 'Planned End', 'Actual End', 'Progress %', 'Status'];
  var sp = [[3, 3], [4, 6], [7, 8], [9, 10], [11, 12], [13, 14], [15, 17]];
  th_(sh, 27, sp, hd);
  var PH = rd_('Project_Phases', 9), P = rd_('Projects', 16);
  var pn = {}; P.forEach(function (r) { pn[String(r[0]).trim()] = String(r[1]).trim(); });
  // Show phases that are behind schedule or critical
  var now = new Date(), items = [];
  PH.forEach(function (r) {
    var pe = r[4], ae = r[6], prog = Number(r[8]) || 0, st = String(r[7]).trim(), pid = String(r[1]).trim();
    var daysLate = 0;
    if (pe instanceof Date && !(ae instanceof Date)) { daysLate = db_(pe, now); if (daysLate > 0 && prog < 100) items.push({ r: r, pid: pid, late: daysLate, prog: prog, st: st }); }
    else if (pe instanceof Date && ae instanceof Date) { daysLate = db_(pe, ae); items.push({ r: r, pid: pid, late: daysLate, prog: prog, st: st }); }
  });
  items.sort(function (a, b) { return b.late - a.late; });
  for (var i = 0; i < Math.min(items.length, 6); i++) {
    var dr = 28 + i, it = items[i], r = it.r, bg = (i % 2 === 0) ? THEME.cardBg : '#F1F5F9';
    if (it.late > 30) bg = THEME.rBg; else if (it.late > 0) bg = THEME.yBg;
    var rk = sh.getRange(dr, 3); rk.setValue(i + 1); sc_(rk, bg, 'center'); rk.setFontWeight('bold').setFontColor(THEME.accent);
    var nm = sh.getRange(dr, 4, 1, 3); nm.merge().setValue((pn[it.pid] || '—').substring(0, 25)); sc_(nm, bg, 'left');
    var ph = sh.getRange(dr, 7, 1, 2); ph.merge().setValue(String(r[2]).trim()); sc_(ph, bg, 'center');
    var pe2 = sh.getRange(dr, 9, 1, 2); pe2.merge().setValue(r[4] instanceof Date ? Utilities.formatDate(r[4], 'Asia/Dubai', 'dd-MMM-yy') : '—'); sc_(pe2, bg, 'center');
    var ae2 = sh.getRange(dr, 11, 1, 2); ae2.merge().setValue(r[6] instanceof Date ? Utilities.formatDate(r[6], 'Asia/Dubai', 'dd-MMM-yy') : 'Ongoing'); sc_(ae2, bg, 'center');
    var pg = sh.getRange(dr, 13, 1, 2); pg.merge().setValue(it.prog + '%'); sc_(pg, bg, 'center'); pg.setFontWeight('bold');
    if (it.prog >= 80) pg.setFontColor(THEME.success); else if (it.prog >= 50) pg.setFontColor(THEME.warning); else pg.setFontColor(THEME.danger);
    var sl = sh.getRange(dr, 15, 1, 3); sl.merge(); sc_(sl, bg, 'center'); sl.setFontWeight('bold');
    if (it.late > 30) { sl.setValue('CRITICAL (' + it.late + 'd late)'); sl.setFontColor(THEME.danger).setBackground(THEME.rBg); }
    else if (it.late > 0) { sl.setValue('BEHIND (' + it.late + 'd)'); sl.setFontColor(THEME.warning).setBackground(THEME.yBg); }
    else { sl.setValue(it.st); sl.setFontColor(THEME.success).setBackground(THEME.gBg); }
  }
}

/* ================================================================ */
/*  TABLE 3: CONTRACTOR RECOMMENDATIONS (rows 42-49)                */
/* ================================================================ */
function buildContractorTable_(sh) {
  var hd = ['Rank', 'Contractor', 'Specialty', 'Rating', 'Contracts', 'Score', 'On-Time', 'Prequal'];
  var sp = [[3, 3], [4, 6], [7, 8], [9, 9], [10, 11], [12, 13], [14, 15], [16, 17]];
  th_(sh, 42, sp, hd);
  var CON = rd_('Contractors', 12), CT = rd_('Contracts', 15);
  var rM = { 'A': 4, 'B': 3, 'C': 2, 'D': 1 };
  var cC = {}; CT.forEach(function (r) { var c = String(r[3]).trim(); if (c) cC[c] = (cC[c] || 0) + 1; });
  // Calculate on-time delivery from contracts
  var onTime = {}, ctTotal = {};
  CT.forEach(function (r) {
    var cid = String(r[3]).trim(), pe = r[8], ae = r[8], st = String(r[14]).trim();
    if (cid) { ctTotal[cid] = (ctTotal[cid] || 0) + 1; if (st === 'Completed' || st === 'Active') onTime[cid] = (onTime[cid] || 0) + 1; }
  });
  CON.sort(function (a, b) {
    var ra = (rM[String(a[10]).trim()] || 0) * 25 + (cC[String(a[0]).trim()] || 0) * 5 + (String(a[11]) === 'True' || a[11] === true ? 10 : 0);
    var rb = (rM[String(b[10]).trim()] || 0) * 25 + (cC[String(b[0]).trim()] || 0) * 5 + (String(b[11]) === 'True' || b[11] === true ? 10 : 0);
    return rb - ra;
  });
  for (var i = 0; i < Math.min(CON.length, 7); i++) {
    var dr = 43 + i, r = CON[i], bg = (i % 2 === 0) ? THEME.cardBg : '#F1F5F9', cid = String(r[0]).trim();
    var rating = String(r[10]).trim(), prq = String(r[11]).trim(), score = (rM[rating] || 0) * 25 + (cC[cid] || 0) * 5 + ((prq === 'True' || r[11] === true) ? 10 : 0);
    var ot = ctTotal[cid] ? Math.round((onTime[cid] || 0) / ctTotal[cid] * 100) : 0;
    var rk = sh.getRange(dr, 3); rk.setValue(i + 1); sc_(rk, bg, 'center'); rk.setFontWeight('bold').setFontColor(THEME.accent);
    var nm = sh.getRange(dr, 4, 1, 3); nm.merge().setValue(String(r[1]).trim().substring(0, 25)); sc_(nm, bg, 'left');
    var sp2 = sh.getRange(dr, 7, 1, 2); sp2.merge().setValue(String(r[2]).trim()); sc_(sp2, bg, 'center');
    var rt = sh.getRange(dr, 9); rt.setValue(rating); sc_(rt, bg, 'center'); rt.setFontWeight('bold');
    if (rating === 'A') rt.setFontColor(THEME.success).setBackground(THEME.gBg);
    else if (rating === 'B') rt.setFontColor(THEME.info).setBackground(THEME.bBg);
    else if (rating === 'C') rt.setFontColor(THEME.warning).setBackground(THEME.yBg);
    else rt.setFontColor(THEME.danger).setBackground(THEME.rBg);
    var cc = sh.getRange(dr, 10, 1, 2); cc.merge().setValue(cC[cid] || 0); sc_(cc, bg, 'center');
    var sc2 = sh.getRange(dr, 12, 1, 2); sc2.merge().setValue(score); sc_(sc2, bg, 'center'); sc2.setFontWeight('bold').setFontColor(THEME.accent);
    var otR = sh.getRange(dr, 14, 1, 2); otR.merge().setValue(ot + '%'); sc_(otR, bg, 'center');
    if (ot >= 80) otR.setFontColor(THEME.success); else if (ot >= 50) otR.setFontColor(THEME.warning); else otR.setFontColor(THEME.danger);
    var pq = sh.getRange(dr, 16, 1, 2); pq.merge().setValue(prq === 'True' || r[11] === true ? 'YES' : 'NO'); sc_(pq, bg, 'center'); pq.setFontWeight('bold');
    if (prq === 'True' || r[11] === true) pq.setFontColor(THEME.success).setBackground(THEME.gBg); else pq.setFontColor(THEME.danger).setBackground(THEME.rBg);
  }
}

/* ================================================================ */
/*  TABLE 4: BUDGET HEALTH MONITOR (rows 54-61)                     */
/* ================================================================ */
function buildBudgetTable_(sh) {
  var hd = ['#', 'Project', 'Contract (M)', 'Planned (M)', 'Actual (M)', 'Variance %', 'VOs (M)', 'Health'];
  var sp = [[3, 3], [4, 6], [7, 8], [9, 10], [11, 12], [13, 14], [15, 15], [16, 17]];
  th_(sh, 54, sp, hd);
  var P = rd_('Projects', 16), WP = rd_('Work_Packages', 9), VO2 = rd_('Variation_Orders', 11), CT = rd_('Contracts', 15);
  var wpPl = {}, wpAc = {};
  WP.forEach(function (r) { var pid = String(r[1]).trim(); wpPl[pid] = (wpPl[pid] || 0) + (Number(r[6]) || 0); wpAc[pid] = (wpAc[pid] || 0) + (Number(r[8]) || 0); });
  var voVal = {}; VO2.forEach(function (r) { var pid = String(r[1]).trim(), v = Number(r[6]) || 0; voVal[pid] = (voVal[pid] || 0) + v; });
  var ctVal = {}; CT.forEach(function (r) { var pid = String(r[1]).trim(), v = Number(r[5]) || 0; ctVal[pid] = (ctVal[pid] || 0) + v; });
  var fP = getFiltered_(P);
  for (var i = 0; i < Math.min(fP.length, 7); i++) {
    var dr = 55 + i, r = fP[i], bg = (i % 2 === 0) ? THEME.cardBg : '#F1F5F9', pid = String(r[0]).trim();
    var cv = ctVal[pid] || Number(r[8]) || 0, pl = wpPl[pid] || 0, ac = wpAc[pid] || 0, vo = voVal[pid] || 0;
    var variance = pl > 0 ? Math.round((ac / pl - 1) * 1000) / 10 : 0;
    var health, hBg, hFg;
    if (variance <= -5) { health = 'UNDER'; hBg = THEME.gBg; hFg = THEME.gFg; }
    else if (variance <= 5) { health = 'ON TRACK'; hBg = THEME.bBg; hFg = THEME.bFg; }
    else if (variance <= 15) { health = 'WARNING'; hBg = THEME.yBg; hFg = THEME.yFg; }
    else { health = 'OVER'; hBg = THEME.rBg; hFg = THEME.rFg; }
    var rk = sh.getRange(dr, 3); rk.setValue(i + 1); sc_(rk, bg, 'center'); rk.setFontWeight('bold').setFontColor(THEME.accent);
    var nm = sh.getRange(dr, 4, 1, 3); nm.merge().setValue(String(r[1]).trim().substring(0, 25)); sc_(nm, bg, 'left');
    var cv2 = sh.getRange(dr, 7, 1, 2); cv2.merge().setValue(fm_(cv) + 'M'); sc_(cv2, bg, 'center');
    var pl2 = sh.getRange(dr, 9, 1, 2); pl2.merge().setValue(fm_(pl) + 'M'); sc_(pl2, bg, 'center');
    var ac2 = sh.getRange(dr, 11, 1, 2); ac2.merge().setValue(fm_(ac) + 'M'); sc_(ac2, bg, 'center');
    var vr = sh.getRange(dr, 13, 1, 2); vr.merge().setValue(variance + '%'); sc_(vr, bg, 'center'); vr.setFontWeight('bold');
    if (variance > 10) vr.setFontColor(THEME.danger); else if (variance > 0) vr.setFontColor(THEME.warning); else vr.setFontColor(THEME.success);
    var vo2 = sh.getRange(dr, 15); vo2.setValue(fm_(vo) + 'M'); sc_(vo2, bg, 'center');
    var hl = sh.getRange(dr, 16, 1, 2); hl.merge().setValue(health); sc_(hl, hBg, 'center'); hl.setFontWeight('bold').setFontColor(hFg);
  }
}

/* ================================================================ */
/*  TABLE 5: CASH FLOW & PAYMENT TRACKING (rows 62-68)              */
/* ================================================================ */
function buildCashFlowTable_(sh) {
  var hd = ['#', 'Project', 'Gross (M)', 'Net Certified', 'Retention', 'Deductions', 'Status'];
  var sp = [[3, 3], [4, 6], [7, 8], [9, 10], [11, 12], [13, 14], [15, 17]];
  th_(sh, 62, sp, hd);
  var PA = rd_('Payment_Applications', 13), P = rd_('Projects', 16);
  var pn = {}; P.forEach(function (r) { pn[String(r[0]).trim()] = String(r[1]).trim(); });
  // Aggregate by project
  var agg = {}; PA.forEach(function (r) {
    var pid = String(r[1]).trim(), gv = Number(r[6]) || 0, nc = Number(r[10]) || 0, rt = Number(r[8]) || 0, dd = Number(r[9]) || 0, st = String(r[11]).trim();
    if (!agg[pid]) agg[pid] = { gv: 0, nc: 0, rt: 0, dd: 0, last: st };
    agg[pid].gv += gv; agg[pid].nc += nc; agg[pid].rt += rt; agg[pid].dd += dd; agg[pid].last = st;
  });
  var sorted = Object.keys(agg).sort(function (a, b) { return agg[b].gv - agg[a].gv; });
  for (var i = 0; i < Math.min(sorted.length, 6); i++) {
    var dr = 63 + i, pid = sorted[i], d = agg[pid], bg = (i % 2 === 0) ? THEME.cardBg : '#F1F5F9';
    var rk = sh.getRange(dr, 3); rk.setValue(i + 1); sc_(rk, bg, 'center'); rk.setFontWeight('bold').setFontColor(THEME.accent);
    var nm = sh.getRange(dr, 4, 1, 3); nm.merge().setValue((pn[pid] || pid).substring(0, 25)); sc_(nm, bg, 'left');
    var gv = sh.getRange(dr, 7, 1, 2); gv.merge().setValue(fm_(d.gv) + 'M'); sc_(gv, bg, 'center');
    var nc = sh.getRange(dr, 9, 1, 2); nc.merge().setValue(fm_(d.nc) + 'M'); sc_(nc, bg, 'center'); nc.setFontWeight('bold').setFontColor(THEME.success);
    var rt = sh.getRange(dr, 11, 1, 2); rt.merge().setValue(fm_(d.rt) + 'M'); sc_(rt, bg, 'center');
    if (d.rt > 500000) rt.setFontColor(THEME.warning);
    var dd = sh.getRange(dr, 13, 1, 2); dd.merge().setValue(fm_(d.dd) + 'M'); sc_(dd, bg, 'center');
    if (d.dd > 200000) dd.setFontColor(THEME.danger);
    var ss2 = sh.getRange(dr, 15, 1, 3); ss2.merge().setValue(d.last); sc_(ss2, bg, 'center'); ss2.setFontWeight('bold');
    if (d.last === 'Paid') ss2.setFontColor(THEME.success).setBackground(THEME.gBg);
    else if (d.last === 'Certified') ss2.setFontColor(THEME.info).setBackground(THEME.bBg);
    else ss2.setFontColor(THEME.warning).setBackground(THEME.yBg);
  }
}

/* ================================================================ */
/*  TABLE 6: EQUIPMENT RENTAL ALERTS (rows 73-79)                   */
/* ================================================================ */
function buildEquipmentTable_(sh) {
  var hd = ['#', 'Equipment Type', 'Model', 'Ownership', 'Daily Rate', 'Status', 'Project', 'Cost Alert'];
  var sp = [[3, 3], [4, 5], [6, 7], [8, 9], [10, 11], [12, 13], [14, 15], [16, 17]];
  th_(sh, 73, sp, hd);
  var EQ = rd_('Equipment', 10);
  var rented = EQ.filter(function (r) { return String(r[4]).trim() === 'Rented'; });
  rented.sort(function (a, b) { return (Number(b[5]) || 0) - (Number(a[5]) || 0); });
  for (var i = 0; i < Math.min(rented.length, 6); i++) {
    var dr = 74 + i, r = rented[i], bg = (i % 2 === 0) ? THEME.cardBg : '#F1F5F9';
    var rate = Number(r[5]) || 0, stat = String(r[6]).trim(), proj = String(r[7]).trim();
    var alert, aBg, aFg;
    if (rate >= 1500) { alert = 'HIGH COST'; aBg = THEME.rBg; aFg = THEME.rFg; }
    else if (rate >= 800) { alert = 'MEDIUM'; aBg = THEME.yBg; aFg = THEME.yFg; }
    else { alert = 'LOW'; aBg = THEME.gBg; aFg = THEME.gFg; }
    var rk = sh.getRange(dr, 3); rk.setValue(i + 1); sc_(rk, bg, 'center'); rk.setFontWeight('bold').setFontColor(THEME.accent);
    var tp = sh.getRange(dr, 4, 1, 2); tp.merge().setValue(String(r[1]).trim()); sc_(tp, bg, 'left');
    var md = sh.getRange(dr, 6, 1, 2); md.merge().setValue(String(r[2]).trim()); sc_(md, bg, 'center');
    var ow = sh.getRange(dr, 8, 1, 2); ow.merge().setValue('Rented'); sc_(ow, bg, 'center'); ow.setFontColor(THEME.warning);
    var rt = sh.getRange(dr, 10, 1, 2); rt.merge().setValue(rate.toLocaleString() + ' AED'); sc_(rt, bg, 'center'); rt.setFontWeight('bold');
    if (rate >= 1500) rt.setFontColor(THEME.danger);
    var st = sh.getRange(dr, 12, 1, 2); st.merge().setValue(stat); sc_(st, bg, 'center');
    if (stat === 'In Use') st.setFontColor(THEME.success); else st.setFontColor(THEME.subText);
    var pj = sh.getRange(dr, 14, 1, 2); pj.merge().setValue(proj.substring(0, 12)); sc_(pj, bg, 'center');
    var al = sh.getRange(dr, 16, 1, 2); al.merge().setValue(alert); sc_(al, aBg, 'center'); al.setFontWeight('bold').setFontColor(aFg);
  }
}

/* ================================================================ */
/*  TABLE 7: VO IMPACT TRACKER (rows 84-89)                         */
/* ================================================================ */
function buildVOTable_(sh) {
  var hd = ['#', 'Project', 'VO Description', 'Value (AED)', 'Days', 'Status', 'Impact'];
  var sp = [[3, 3], [4, 5], [6, 9], [10, 11], [12, 13], [14, 15], [16, 17]];
  th_(sh, 84, sp, hd);
  var VO2 = rd_('Variation_Orders', 11), P = rd_('Projects', 16);
  var pn = {}; P.forEach(function (r) { pn[String(r[0]).trim()] = String(r[1]).trim(); });
  VO2.sort(function (a, b) { return (Number(b[6]) || 0) - (Number(a[6]) || 0); });
  for (var i = 0; i < Math.min(VO2.length, 5); i++) {
    var dr = 85 + i, r = VO2[i], bg = (i % 2 === 0) ? THEME.cardBg : '#F1F5F9';
    var val = Number(r[6]) || 0, days = Number(r[7]) || 0, stat = String(r[8]).trim();
    var impact, iBg, iFg;
    if (val > 1000000 || days > 60) { impact = 'CRITICAL'; iBg = THEME.rBg; iFg = THEME.rFg; }
    else if (val > 500000 || days > 30) { impact = 'HIGH'; iBg = THEME.yBg; iFg = THEME.yFg; }
    else { impact = 'NORMAL'; iBg = THEME.gBg; iFg = THEME.gFg; }
    var rk = sh.getRange(dr, 3); rk.setValue(i + 1); sc_(rk, bg, 'center'); rk.setFontWeight('bold').setFontColor(THEME.accent);
    var pj = sh.getRange(dr, 4, 1, 2); pj.merge().setValue((pn[String(r[1]).trim()] || '—').substring(0, 18)); sc_(pj, bg, 'left');
    var ds = sh.getRange(dr, 6, 1, 4); ds.merge().setValue(String(r[3]).trim().substring(0, 40)); sc_(ds, bg, 'left');
    var vl = sh.getRange(dr, 10, 1, 2); vl.merge().setValue(val.toLocaleString()); sc_(vl, bg, 'center'); vl.setFontWeight('bold');
    if (val > 1000000) vl.setFontColor(THEME.danger);
    var dy = sh.getRange(dr, 12, 1, 2); dy.merge().setValue(days + ' days'); sc_(dy, bg, 'center');
    if (days > 60) dy.setFontColor(THEME.danger); else if (days > 30) dy.setFontColor(THEME.warning); else dy.setFontColor(THEME.success);
    var ss2 = sh.getRange(dr, 14, 1, 2); ss2.merge().setValue(stat); sc_(ss2, bg, 'center'); ss2.setFontWeight('bold');
    if (stat === 'Approved') ss2.setFontColor(THEME.success); else if (stat === 'Pending') ss2.setFontColor(THEME.warning); else ss2.setFontColor(THEME.danger);
    var im = sh.getRange(dr, 16, 1, 2); im.merge().setValue(impact); sc_(im, iBg, 'center'); im.setFontWeight('bold').setFontColor(iFg);
  }
}

/* ================================================================ */
/*  TABLE 8: LAD RISK CALCULATOR (rows 90-95)                       */
/* ================================================================ */
function buildLADRiskTable_(sh) {
  var hd = ['#', 'Project', 'LAD/Day (AED)', 'Days Delayed', 'Total LAD Risk', 'Contract', 'Risk'];
  var sp = [[3, 3], [4, 6], [7, 8], [9, 10], [11, 12], [13, 14], [15, 17]];
  th_(sh, 90, sp, hd);
  var CT = rd_('Contracts', 15), P = rd_('Projects', 16);
  var pn = {}, pStat = {};
  P.forEach(function (r) { pn[String(r[0]).trim()] = String(r[1]).trim(); pStat[String(r[0]).trim()] = String(r[11]).trim(); });
  var items = [];
  CT.forEach(function (r) {
    var pid = String(r[1]).trim(), lad = Number(r[12]) || 0, cv = Number(r[5]) || 0;
    var pe = r[8], now = new Date(), delayed = 0;
    if (pe instanceof Date && pStat[pid] !== 'Completed') { delayed = Math.max(0, db_(pe, now)); }
    if (lad > 0) items.push({ pid: pid, lad: lad, delayed: delayed, risk: lad * delayed, cv: cv });
  });
  items.sort(function (a, b) { return b.risk - a.risk; });
  for (var i = 0; i < Math.min(items.length, 5); i++) {
    var dr = 91 + i, it = items[i], bg = (i % 2 === 0) ? THEME.cardBg : '#F1F5F9';
    var riskLvl, rBg, rFg;
    if (it.risk > 500000) { riskLvl = 'CRITICAL'; rBg = THEME.rBg; rFg = THEME.rFg; }
    else if (it.risk > 100000) { riskLvl = 'HIGH'; rBg = THEME.yBg; rFg = THEME.yFg; }
    else { riskLvl = 'LOW'; rBg = THEME.gBg; rFg = THEME.gFg; }
    var rk = sh.getRange(dr, 3); rk.setValue(i + 1); sc_(rk, bg, 'center'); rk.setFontWeight('bold').setFontColor(THEME.accent);
    var nm = sh.getRange(dr, 4, 1, 3); nm.merge().setValue((pn[it.pid] || it.pid).substring(0, 25)); sc_(nm, bg, 'left');
    var ld = sh.getRange(dr, 7, 1, 2); ld.merge().setValue(it.lad.toLocaleString()); sc_(ld, bg, 'center');
    var dl = sh.getRange(dr, 9, 1, 2); dl.merge().setValue(it.delayed + ' days'); sc_(dl, bg, 'center');
    if (it.delayed > 60) dl.setFontColor(THEME.danger).setFontWeight('bold'); else if (it.delayed > 0) dl.setFontColor(THEME.warning);
    var tr = sh.getRange(dr, 11, 1, 2); tr.merge().setValue(fm_(it.risk) + 'M'); sc_(tr, bg, 'center'); tr.setFontWeight('bold');
    if (it.risk > 500000) tr.setFontColor(THEME.danger);
    var cv = sh.getRange(dr, 13, 1, 2); cv.merge().setValue(fm_(it.cv) + 'M'); sc_(cv, bg, 'center');
    var rl = sh.getRange(dr, 15, 1, 3); rl.merge().setValue(riskLvl); sc_(rl, rBg, 'center'); rl.setFontWeight('bold').setFontColor(rFg);
  }
}

/* ================================================================ */
/*  TABLE 9: SAFETY INCIDENTS (rows 100-106)                        */
/* ================================================================ */
function buildSafetyTable_(sh) {
  var hd = ['#', 'Project', 'Type', 'Severity', 'Date', 'LTI Days', 'Status'];
  var sp = [[3, 3], [4, 6], [7, 9], [10, 11], [12, 13], [14, 15], [16, 17]];
  th_(sh, 100, sp, hd);
  var SI = rd_('Safety_Incidents', 12), P = rd_('Projects', 16);
  var pn = {}; P.forEach(function (r) { pn[String(r[0]).trim()] = String(r[1]).trim(); });
  SI.sort(function (a, b) { var da = a[2] instanceof Date ? a[2].getTime() : 0, db2 = b[2] instanceof Date ? b[2].getTime() : 0; return db2 - da; });
  var fSI = SI, fpid = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DN);
  for (var i = 0; i < Math.min(SI.length, 6); i++) {
    var dr = 101 + i, r = SI[i], bg = (i % 2 === 0) ? THEME.cardBg : '#F1F5F9';
    var sev = String(r[3]).trim(), stat = String(r[11]).trim(), lti = Number(r[8]) || 0;
    if (sev === 'Critical' || sev === 'High') bg = THEME.rBg;
    var rk = sh.getRange(dr, 3); rk.setValue(i + 1); sc_(rk, bg, 'center'); rk.setFontWeight('bold').setFontColor(THEME.accent);
    var nm = sh.getRange(dr, 4, 1, 3); nm.merge().setValue((pn[String(r[1]).trim()] || '—').substring(0, 25)); sc_(nm, bg, 'left');
    var tp = sh.getRange(dr, 7, 1, 3); tp.merge().setValue(String(r[3]).trim()); sc_(tp, bg, 'center');
    var sv = sh.getRange(dr, 10, 1, 2); sv.merge().setValue(sev); sc_(sv, bg, 'center'); sv.setFontWeight('bold');
    if (sev === 'Critical' || sev === 'High') sv.setFontColor(THEME.danger);
    else if (sev === 'Medium') sv.setFontColor(THEME.warning); else sv.setFontColor(THEME.success);
    var dt = sh.getRange(dr, 12, 1, 2); dt.merge().setValue(r[2] instanceof Date ? Utilities.formatDate(r[2], 'Asia/Dubai', 'dd-MMM-yy') : '—'); sc_(dt, bg, 'center');
    var lt = sh.getRange(dr, 14, 1, 2); lt.merge().setValue(lti); sc_(lt, bg, 'center');
    if (lti > 0) lt.setFontColor(THEME.danger).setFontWeight('bold');
    var ss2 = sh.getRange(dr, 16, 1, 2); ss2.merge().setValue(stat); sc_(ss2, bg, 'center'); ss2.setFontWeight('bold');
    if (stat === 'Closed') ss2.setFontColor(THEME.success); else ss2.setFontColor(THEME.warning);
  }
}

/* ================================================================ */
/*  TABLE 10: INSPECTION DEFECTS (rows 107-112)                     */
/* ================================================================ */
function buildInspectionTable_(sh) {
  var hd = ['#', 'Project', 'Type', 'Date', 'Result', 'Defects', 'Action Required'];
  var sp = [[3, 3], [4, 6], [7, 8], [9, 10], [11, 12], [13, 14], [15, 17]];
  th_(sh, 107, sp, hd);
  var INS = rd_('Inspections', 11), P = rd_('Projects', 16);
  var pn = {}; P.forEach(function (r) { pn[String(r[0]).trim()] = String(r[1]).trim(); });
  // Show failed inspections first
  var failed = INS.filter(function (r) { return String(r[6]).trim() === 'Failed'; });
  failed.sort(function (a, b) { var da = a[3] instanceof Date ? a[3].getTime() : 0, db2 = b[3] instanceof Date ? b[3].getTime() : 0; return db2 - da; });
  for (var i = 0; i < Math.min(failed.length, 5); i++) {
    var dr = 108 + i, r = failed[i], bg = (i % 2 === 0) ? THEME.rBg : '#FEF2F2';
    var rk = sh.getRange(dr, 3); rk.setValue(i + 1); sc_(rk, bg, 'center'); rk.setFontWeight('bold').setFontColor(THEME.accent);
    var nm = sh.getRange(dr, 4, 1, 3); nm.merge().setValue((pn[String(r[1]).trim()] || '—').substring(0, 25)); sc_(nm, bg, 'left');
    var tp = sh.getRange(dr, 7, 1, 2); tp.merge().setValue(String(r[2]).trim()); sc_(tp, bg, 'center');
    var dt = sh.getRange(dr, 9, 1, 2); dt.merge().setValue(r[3] instanceof Date ? Utilities.formatDate(r[3], 'Asia/Dubai', 'dd-MMM-yy') : '—'); sc_(dt, bg, 'center');
    var rs = sh.getRange(dr, 11, 1, 2); rs.merge().setValue('FAILED'); sc_(rs, bg, 'center'); rs.setFontWeight('bold').setFontColor(THEME.danger);
    var df = sh.getRange(dr, 13, 1, 2); df.merge().setValue(Number(r[7]) || 0); sc_(df, bg, 'center'); df.setFontWeight('bold');
    if ((Number(r[7]) || 0) > 3) df.setFontColor(THEME.danger);
    var ac = sh.getRange(dr, 15, 1, 3); ac.merge().setValue(String(r[8]).trim().substring(0, 30) || 'Pending'); sc_(ac, bg, 'left');
    if (String(r[8]).trim()) ac.setFontColor(THEME.info); else ac.setFontColor(THEME.danger);
  }
}

/* ================================================================ */
/*  TABLE 11: PERMITS EXPIRY TRACKER (rows 115-121)                 */
/* ================================================================ */
function buildPermitsTable_(sh) {
  var hd = ['#', 'Permit Type', 'Project', 'Issue Date', 'Expiry Date', 'Status', 'Days Left'];
  var sp = [[3, 3], [4, 6], [7, 9], [10, 11], [12, 13], [14, 15], [16, 17]];
  th_(sh, 115, sp, hd);
  var PM = rd_('Permits_Approvals', 10), P = rd_('Projects', 16);
  var pn = {}; P.forEach(function (r) { pn[String(r[0]).trim()] = String(r[1]).trim(); });
  var now = new Date();
  PM.sort(function (a, b) { var da = a[6] instanceof Date ? a[6].getTime() : Infinity, db2 = b[6] instanceof Date ? b[6].getTime() : Infinity; return da - db2; });
  for (var i = 0; i < Math.min(PM.length, 6); i++) {
    var dr = 116 + i, r = PM[i], bg = (i % 2 === 0) ? THEME.cardBg : '#F1F5F9';
    var exp = r[6], daysLeft = exp instanceof Date ? db_(now, exp) : 999, stat = String(r[7]).trim();
    if (daysLeft < 0) bg = THEME.rBg; else if (daysLeft < 30) bg = THEME.yBg;
    var rk = sh.getRange(dr, 3); rk.setValue(i + 1); sc_(rk, bg, 'center'); rk.setFontWeight('bold').setFontColor(THEME.accent);
    var pt = sh.getRange(dr, 4, 1, 3); pt.merge().setValue(String(r[2]).trim()); sc_(pt, bg, 'left');
    var pj = sh.getRange(dr, 7, 1, 3); pj.merge().setValue((pn[String(r[1]).trim()] || '—').substring(0, 20)); sc_(pj, bg, 'center');
    var id2 = sh.getRange(dr, 10, 1, 2); id2.merge().setValue(r[5] instanceof Date ? Utilities.formatDate(r[5], 'Asia/Dubai', 'dd-MMM-yy') : '—'); sc_(id2, bg, 'center');
    var ed = sh.getRange(dr, 12, 1, 2); ed.merge().setValue(exp instanceof Date ? Utilities.formatDate(exp, 'Asia/Dubai', 'dd-MMM-yy') : '—'); sc_(ed, bg, 'center');
    var ss2 = sh.getRange(dr, 14, 1, 2); ss2.merge().setValue(stat); sc_(ss2, bg, 'center'); ss2.setFontWeight('bold');
    if (stat === 'Approved' || stat === 'Active') ss2.setFontColor(THEME.success); else if (stat === 'Pending') ss2.setFontColor(THEME.warning); else ss2.setFontColor(THEME.danger);
    var dl2 = sh.getRange(dr, 16, 1, 2); dl2.merge().setValue(daysLeft < 0 ? 'EXPIRED' : daysLeft + ' days'); sc_(dl2, bg, 'center'); dl2.setFontWeight('bold');
    if (daysLeft < 0) dl2.setFontColor(THEME.danger).setBackground(THEME.rBg);
    else if (daysLeft < 30) dl2.setFontColor(THEME.warning).setBackground(THEME.yBg);
    else dl2.setFontColor(THEME.success).setBackground(THEME.gBg);
  }
}

/* ================================================================ */
/*  TABLE 12: DOCUMENT SUBMISSION STATUS (rows 122-127)             */
/* ================================================================ */
function buildDocStatusTable_(sh) {
  var hd = ['#', 'Project', 'Document Title', 'Type', 'Submitted', 'Status'];
  var sp = [[3, 3], [4, 6], [7, 9], [10, 11], [12, 13], [14, 17]];
  th_(sh, 122, sp, hd);
  var DOC = rd_('Project_Documents', 11), P = rd_('Projects', 16);
  var pn = {}; P.forEach(function (r) { pn[String(r[0]).trim()] = String(r[1]).trim(); });
  // Show pending / rejected docs first
  var pending = DOC.filter(function (r) { var s = String(r[8]).trim(); return s === 'Pending' || s === 'Rejected' || s === 'Under Review'; });
  pending.sort(function (a, b) { var sa = String(a[8]).trim() === 'Rejected' ? 0 : 1, sb = String(b[8]).trim() === 'Rejected' ? 0 : 1; return sa - sb; });
  for (var i = 0; i < Math.min(pending.length, 5); i++) {
    var dr = 123 + i, r = pending[i], bg = (i % 2 === 0) ? THEME.cardBg : '#F1F5F9', stat = String(r[8]).trim();
    if (stat === 'Rejected') bg = THEME.rBg;
    var rk = sh.getRange(dr, 3); rk.setValue(i + 1); sc_(rk, bg, 'center'); rk.setFontWeight('bold').setFontColor(THEME.accent);
    var nm = sh.getRange(dr, 4, 1, 3); nm.merge().setValue((pn[String(r[1]).trim()] || '—').substring(0, 22)); sc_(nm, bg, 'left');
    var tt = sh.getRange(dr, 7, 1, 3); tt.merge().setValue(String(r[3]).trim().substring(0, 30)); sc_(tt, bg, 'left');
    var tp = sh.getRange(dr, 10, 1, 2); tp.merge().setValue(String(r[2]).trim()); sc_(tp, bg, 'center');
    var sd = sh.getRange(dr, 12, 1, 2); sd.merge().setValue(r[5] instanceof Date ? Utilities.formatDate(r[5], 'Asia/Dubai', 'dd-MMM-yy') : '—'); sc_(sd, bg, 'center');
    var ss2 = sh.getRange(dr, 14, 1, 4); ss2.merge().setValue(stat); sc_(ss2, bg, 'center'); ss2.setFontWeight('bold');
    if (stat === 'Approved') ss2.setFontColor(THEME.success).setBackground(THEME.gBg);
    else if (stat === 'Rejected') ss2.setFontColor(THEME.danger).setBackground(THEME.rBg);
    else if (stat === 'Under Review') ss2.setFontColor(THEME.info).setBackground(THEME.bBg);
    else ss2.setFontColor(THEME.warning).setBackground(THEME.yBg);
  }
}

/* ================================================================ */
/*  TABLE 13: RETENTION MONEY TRACKER (rows 128-133)                */
/* ================================================================ */
function buildRetentionTable_(sh) {
  var hd = ['#', 'Project', 'Retention %', 'Total Retention', 'Contract', 'DLP (mo)', 'Release'];
  var sp = [[3, 3], [4, 6], [7, 8], [9, 10], [11, 12], [13, 14], [15, 17]];
  th_(sh, 128, sp, hd);
  var CT = rd_('Contracts', 15), P = rd_('Projects', 16), PA = rd_('Payment_Applications', 13);
  var pn = {}; P.forEach(function (r) { pn[String(r[0]).trim()] = String(r[1]).trim(); });
  // Aggregate retention by project
  var retAgg = {}; PA.forEach(function (r) { var pid = String(r[1]).trim(), rt = Number(r[8]) || 0; retAgg[pid] = (retAgg[pid] || 0) + rt; });
  var items = [];
  CT.forEach(function (r) {
    var pid = String(r[1]).trim(), retPct = Number(r[9]) || 0, cv = Number(r[5]) || 0, dlp = Number(r[13]) || 0;
    var totalRet = retAgg[pid] || 0;
    if (retPct > 0 || totalRet > 0) items.push({ pid: pid, retPct: retPct, totalRet: totalRet, cv: cv, dlp: dlp });
  });
  items.sort(function (a, b) { return b.totalRet - a.totalRet; });
  for (var i = 0; i < Math.min(items.length, 5); i++) {
    var dr = 129 + i, it = items[i], bg = (i % 2 === 0) ? THEME.cardBg : '#F1F5F9';
    var release = it.dlp > 12 ? 'Pending' : 'Due Soon';
    var rk = sh.getRange(dr, 3); rk.setValue(i + 1); sc_(rk, bg, 'center'); rk.setFontWeight('bold').setFontColor(THEME.accent);
    var nm = sh.getRange(dr, 4, 1, 3); nm.merge().setValue((pn[it.pid] || it.pid).substring(0, 25)); sc_(nm, bg, 'left');
    var rp = sh.getRange(dr, 7, 1, 2); rp.merge().setValue(it.retPct + '%'); sc_(rp, bg, 'center');
    var tr = sh.getRange(dr, 9, 1, 2); tr.merge().setValue(fm_(it.totalRet) + 'M'); sc_(tr, bg, 'center'); tr.setFontWeight('bold');
    if (it.totalRet > 1000000) tr.setFontColor(THEME.warning);
    var cv = sh.getRange(dr, 11, 1, 2); cv.merge().setValue(fm_(it.cv) + 'M'); sc_(cv, bg, 'center');
    var dp = sh.getRange(dr, 13, 1, 2); dp.merge().setValue(it.dlp + 'm'); sc_(dp, bg, 'center');
    var rl = sh.getRange(dr, 15, 1, 3); rl.merge().setValue(release); sc_(rl, bg, 'center'); rl.setFontWeight('bold');
    if (release === 'Due Soon') rl.setFontColor(THEME.warning).setBackground(THEME.yBg);
    else rl.setFontColor(THEME.info).setBackground(THEME.bBg);
  }
}

/* ================================================================ */
/*  TABLE 14: PO PROCUREMENT PIPELINE (rows 136-141)                */
/* ================================================================ */
function buildPOPipelineTable_(sh) {
  var hd = ['#', 'Project', 'Description', 'Amount (AED)', 'Delivery', 'Status'];
  var sp = [[3, 3], [4, 6], [7, 9], [10, 12], [13, 14], [15, 17]];
  th_(sh, 136, sp, hd);
  var PO = rd_('Purchase_Orders', 10), P = rd_('Projects', 16);
  var pn = {}; P.forEach(function (r) { pn[String(r[0]).trim()] = String(r[1]).trim(); });
  // Show pending / in-progress POs
  var pending = PO.filter(function (r) { var s = String(r[8]).trim(); return s === 'Pending' || s === 'In Progress' || s === 'Ordered'; });
  pending.sort(function (a, b) { return (Number(b[7]) || 0) - (Number(a[7]) || 0); });
  for (var i = 0; i < Math.min(pending.length, 5); i++) {
    var dr = 137 + i, r = pending[i], bg = (i % 2 === 0) ? THEME.cardBg : '#F1F5F9', stat = String(r[8]).trim();
    var rk = sh.getRange(dr, 3); rk.setValue(i + 1); sc_(rk, bg, 'center'); rk.setFontWeight('bold').setFontColor(THEME.accent);
    var nm = sh.getRange(dr, 4, 1, 3); nm.merge().setValue((pn[String(r[1]).trim()] || '—').substring(0, 22)); sc_(nm, bg, 'left');
    var ds = sh.getRange(dr, 7, 1, 3); ds.merge().setValue(String(r[4]).trim().substring(0, 30)); sc_(ds, bg, 'left');
    var am = sh.getRange(dr, 10, 1, 3); am.merge().setValue((Number(r[7]) || 0).toLocaleString()); sc_(am, bg, 'center'); am.setFontWeight('bold');
    if ((Number(r[7]) || 0) > 500000) am.setFontColor(THEME.danger);
    var dl = sh.getRange(dr, 13, 1, 2); dl.merge().setValue(r[6] instanceof Date ? Utilities.formatDate(r[6], 'Asia/Dubai', 'dd-MMM-yy') : '—'); sc_(dl, bg, 'center');
    var ss2 = sh.getRange(dr, 15, 1, 3); ss2.merge().setValue(stat); sc_(ss2, bg, 'center'); ss2.setFontWeight('bold');
    if (stat === 'Pending') ss2.setFontColor(THEME.warning).setBackground(THEME.yBg);
    else if (stat === 'In Progress' || stat === 'Ordered') ss2.setFontColor(THEME.info).setBackground(THEME.bBg);
    else ss2.setFontColor(THEME.success).setBackground(THEME.gBg);
  }
}

/* ================================================================ */
/*  TABLE 15: DECISION MATRIX (rows 142-152)                        */
/* ================================================================ */
function buildDecisionMatrix_(sh) {
  var hd = ['Project Type', 'Duration(mo)', 'Budget Var %', 'VOs', 'Safety', 'Pass %', 'CPI', 'Risk'];
  var sp = [[3, 5], [6, 7], [8, 9], [10, 10], [11, 12], [13, 13], [14, 15], [16, 17]];
  th_(sh, 142, sp, hd);
  var P = rd_('Projects', 16), WP = rd_('Work_Packages', 9), VO2 = rd_('Variation_Orders', 11);
  var SI = rd_('Safety_Incidents', 12), INS = rd_('Inspections', 11), PA = rd_('Payment_Applications', 13), CT = rd_('Contracts', 15);
  var types = uq_(P, 2), pMap2 = {}, tPC = {};
  P.forEach(function (r) { var tp = String(r[2]).trim(), pid = String(r[0]).trim(); pMap2[pid] = tp; tPC[tp] = (tPC[tp] || 0) + 1; });
  var durS = {}, durC = {}; P.forEach(function (r) { var tp = String(r[2]).trim(), s = r[9], e = r[10]; if (tp && s instanceof Date && e instanceof Date) { durS[tp] = (durS[tp] || 0) + Math.round((e - s) / 864e5 / 30); durC[tp] = (durC[tp] || 0) + 1; } });
  var wpPl = {}, wpAc = {}; WP.forEach(function (r) { var tp = pMap2[String(r[1]).trim()]; if (tp) { wpPl[tp] = (wpPl[tp] || 0) + (Number(r[6]) || 0); wpAc[tp] = (wpAc[tp] || 0) + (Number(r[8]) || 0); } });
  var voC = {}; VO2.forEach(function (r) { var tp = pMap2[String(r[1]).trim()]; if (tp) voC[tp] = (voC[tp] || 0) + 1; });
  var sfC = {}; SI.forEach(function (r) { var tp = pMap2[String(r[1]).trim()]; if (tp) sfC[tp] = (sfC[tp] || 0) + 1; });
  var inP = {}, inT2 = {}; INS.forEach(function (r) { var tp = pMap2[String(r[1]).trim()], rs = String(r[6]).trim(); if (tp && rs) { inT2[tp] = (inT2[tp] || 0) + 1; if (rs === 'Passed') inP[tp] = (inP[tp] || 0) + 1; } });
  var cpN = {}, cpD = {}; PA.forEach(function (r) { var tp = pMap2[String(r[1]).trim()], v = Number(r[10]) || 0, s = String(r[11]).trim(); if (tp && (s === 'Certified' || s === 'Paid')) cpN[tp] = (cpN[tp] || 0) + v; });
  CT.forEach(function (r) { var tp = pMap2[String(r[1]).trim()], v = Number(r[5]) || 0; if (tp && v > 0) cpD[tp] = (cpD[tp] || 0) + v; });
  for (var i = 0; i < Math.min(types.length, 10); i++) {
    var dr = 143 + i, tp = types[i], bg = (i % 2 === 0) ? THEME.cardBg : '#F1F5F9', pc = tPC[tp] || 1;
    var avgDur = durC[tp] ? Math.round(durS[tp] / durC[tp] * 10) / 10 : 0;
    var budVar = wpPl[tp] ? Math.round(((wpAc[tp] || 0) / wpPl[tp] - 1) * 1000) / 10 : 0;
    var avgVO = Math.round((voC[tp] || 0) / pc * 10) / 10;
    var safScore = Math.round((sfC[tp] || 0) / pc * 10) / 10;
    var passRate = inT2[tp] ? Math.round((inP[tp] || 0) / inT2[tp] * 1000) / 10 : 0;
    var cpi = cpD[tp] ? Math.round((cpN[tp] || 0) / cpD[tp] * 100) / 100 : 0;
    var risk = (budVar > 15 || safScore > 2 || cpi < 0.7) ? 'HIGH' : (budVar < 5 && safScore < 1 && cpi > 0.9) ? 'LOW' : 'MEDIUM';
    var pt = sh.getRange(dr, 3, 1, 3); pt.merge().setValue(tp); sc_(pt, bg, 'left'); pt.setFontWeight('bold');
    var ad = sh.getRange(dr, 6, 1, 2); ad.merge().setValue(avgDur); sc_(ad, bg, 'center');
    var bv = sh.getRange(dr, 8, 1, 2); bv.merge().setValue(budVar + '%'); sc_(bv, bg, 'center');
    if (budVar > 10) bv.setFontColor(THEME.danger).setFontWeight('bold'); else if (budVar > 0) bv.setFontColor(THEME.warning); else bv.setFontColor(THEME.success);
    var vo = sh.getRange(dr, 10); vo.setValue(avgVO); sc_(vo, bg, 'center');
    var sf = sh.getRange(dr, 11, 1, 2); sf.merge().setValue(safScore); sc_(sf, bg, 'center');
    if (safScore > 2) sf.setFontColor(THEME.danger).setFontWeight('bold');
    var pr = sh.getRange(dr, 13); pr.setValue(passRate + '%'); sc_(pr, bg, 'center');
    if (passRate < 60) pr.setFontColor(THEME.danger); else if (passRate < 80) pr.setFontColor(THEME.warning); else pr.setFontColor(THEME.success);
    var cp = sh.getRange(dr, 14, 1, 2); cp.merge().setValue(cpi); sc_(cp, bg, 'center'); cp.setFontWeight('bold');
    if (cpi < 0.8) cp.setFontColor(THEME.danger); else if (cpi < 1.0) cp.setFontColor(THEME.warning); else cp.setFontColor(THEME.success);
    var rl = sh.getRange(dr, 16, 1, 2); rl.merge().setValue(risk); sc_(rl, bg, 'center'); rl.setFontWeight('bold');
    if (risk === 'LOW') rl.setBackground(THEME.gBg).setFontColor(THEME.gFg);
    else if (risk === 'HIGH') rl.setBackground(THEME.rBg).setFontColor(THEME.rFg);
    else rl.setBackground(THEME.yBg).setFontColor(THEME.yFg);
  }
  // Footer
  sh.getRange(153, 3, 1, 15).merge()
    .setValue('25+ Operational Intelligence Panels  |  17 Data Modules  |  All Values in AED  |  KPIs react to filters  |  \u26A0 = Alert Conditions')
    .setFontFamily(THEME.font).setFontSize(8).setFontColor(THEME.subText)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(THEME.bg).setFontStyle('italic');
}

/* ================================================================ */
/*  FILTER HELPERS                                                  */
/* ================================================================ */
function getFilterVals_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DN);
  if (!sh) return { fp: 'All Projects', ft: 'All Types', fs: 'All Statuses' };
  return { fp: String(sh.getRange('D4').getValue()).trim(), ft: String(sh.getRange('H4').getValue()).trim(), fs: String(sh.getRange('L4').getValue()).trim() };
}
function getFiltered_(projects) {
  var f = getFilterVals_();
  return projects.filter(function (r) {
    var pid = String(r[0]).trim(), tp = String(r[2]).trim(), st = String(r[11]).trim();
    if (f.fp !== 'All Projects' && pid !== f.fp) return false;
    if (f.ft !== 'All Types' && tp !== f.ft) return false;
    if (f.fs !== 'All Statuses' && st !== f.fs) return false;
    return true;
  });
}
/* Returns a map of project_ids that pass the current filter — used for secondary sheets */
function getFilteredPids_() {
  var P = rd_('Projects', 16), fP = getFiltered_(P), pids = {};
  var f = getFilterVals_();
  if (f.fp === 'All Projects' && f.ft === 'All Types' && f.fs === 'All Statuses') {
    P.forEach(function (r) { pids[String(r[0]).trim()] = true; });
  } else {
    fP.forEach(function (r) { pids[String(r[0]).trim()] = true; });
  }
  return pids;
}

/* ================================================================ */
/*  onEdit — FILTER REACTIVITY + CHECKBOX RESET                     */
/* ================================================================ */
function onEdit(e) {
  if (!e) return;
  var sh = e.range.getSheet(); if (sh.getName() !== DN) return;
  var r = e.range.getRow(), c = e.range.getColumn();

  // RESET checkbox: P4 = column 16
  if (r === 4 && c === 16) {
    var val = sh.getRange('P4').getValue();
    if (val === true) {
      sh.getRange('D4').setValue('All Projects');
      sh.getRange('H4').setValue('All Types');
      sh.getRange('L4').setValue('All Statuses');
      sh.getRange('P4').setValue(false);
      SpreadsheetApp.flush();
      rebuildTables_(sh);
    }
    return;
  }

  // Filter dropdowns: D4(col4), H4(col8), L4(col12)
  if (r === 4 && (c === 4 || c === 5 || c === 8 || c === 9 || c === 12 || c === 13)) {
    SpreadsheetApp.flush();
    rebuildTables_(sh);
  }
}

/* ================================================================ */
/*  REBUILD ALL — called by onEdit for full filter reactivity       */
/* ================================================================ */
function rebuildTables_(sh) {
  // KPIs (all 12 react to filters + pain points banner)
  computeKPIs_(sh);
  // New Projects (rows 20-26)
  sh.getRange(20, 3, 7, 15).clearContent().clearFormat().setBackground(THEME.bg);
  buildNewProjectsTable_(sh);
  // Budget (rows 55-61)
  sh.getRange(55, 3, 7, 15).clearContent().clearFormat().setBackground(THEME.bg);
  buildBudgetTable_(sh);
  // Cash Flow (rows 63-68)
  sh.getRange(63, 3, 6, 15).clearContent().clearFormat().setBackground(THEME.bg);
  buildCashFlowTable_(sh);
  SpreadsheetApp.flush();
}
