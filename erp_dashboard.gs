/**
 * ============================================================
 * DUBAI ERP — DASHBOARD SHARED CORE  (erp_dashboard.gs)
 * ============================================================
 * Config · Data · Sheet setup · All drawing primitives
 * Page builders live in:  db_overview · db_financial · db_projects
 *                         db_workforce · db_operations · db_hse
 * ============================================================
 */

// ─── CONFIG ──────────────────────────────────────────────────
var DB = {
  FS   : '_DB_Filters',
  FROW : 6,
  FCOLS: { proj: 7, type: 13, stat: 19 },

  PAGES: [
    { k:'overview',   i:'📊', l:'Overview',    s:'📊 Overview'    },
    { k:'financial',  i:'💰', l:'Financial',   s:'💰 Financial'   },
    { k:'projects',   i:'🏗️',  l:'Projects',    s:'🏗️ Projects'    },
    { k:'workforce',  i:'👷', l:'Workforce',   s:'👷 Workforce'   },
    { k:'operations', i:'⚙️',  l:'Operations',  s:'⚙️ Operations'  },
    { k:'hse',        i:'🛡️',  l:'HSE & Safety',s:'🛡️ HSE & Safety'}
  ],

  C: {
    PAGE : '#f0f2f5',                      // light page bg
    PANEL: '#ffffff',                      // white card / panel
    HDR  : '#1e293b', HDR_T: '#ffffff',    // navy header bar
    IND  : '#6366f1', IND_L: '#eef2ff',    // indigo accent / light tint
    IND_M: '#6366f1', NAV_T: '#64748b',    // indigo pill, nav inactive

    // KPI accent fg / soft pastel icon bg
    G : '#16a34a', GL: '#dcfce7',  // emerald
    B : '#2563eb', BL: '#dbeafe',  // blue
    A : '#d97706', AL: '#fef3c7',  // amber
    R : '#dc2626', RL: '#fee2e2',  // red
    P : '#9333ea', PL: '#f3e8ff',  // purple
    T : '#0891b2', TL: '#cffafe',  // teal
    I : '#4f46e5', IL: '#eef2ff',  // indigo

    TH  : '#f1f5f9', TH_T: '#1e293b',  // light table header
    TR1 : '#ffffff', TR2 : '#f8fafc',   // white / off-white rows
    MU  : '#64748b', BD  : '#e2e8f0',   // muted text, border

    // Extra for light theme
    TXT : '#0f172a',  // primary text
    TXT2: '#334155',  // secondary text
    SEC_BG: '#f8fafc', SEC_T: '#1e293b', // section header
    NAV_BG: '#ffffff', NAV_ACT: '#6366f1' // nav bar
  },

  FONT: 'Inter, Arial'
};

// ─── ENTRY POINTS ────────────────────────────────────────────

/** First-run: ONLY creates sheets + hides gridlines. Fast — no data, no formatting. */
function buildDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open your Google Sheet first, then run from the Dubai ERP menu.');
  ss.toast('Creating sheets…', '🏗️', -1);
  _dbSetupFS(ss);
  DB.PAGES.forEach(function(p) {
    var sh = ss.getSheetByName(p.s) || ss.insertSheet(p.s);
    sh.setHiddenGridlines(true);
    sh.setTabColor(DB.C.IND);
  });
  SpreadsheetApp.flush();
  ss.toast('✅ Sheets created — now build each page from the menu', '📊', 8);
}

/** Shared loader — called by every individual page builder */
function _dbPrepare(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet found. Run this from the Google Sheets menu.');
  _dbSetupFS(ss);
  var all  = _dbLoad(ss);
  var filt = _dbGetFilt(ss);
  var fd   = _dbFilt(all, filt);
  var gids = _dbGids(ss);
  var cols = _dbColMap(ss);
  return { ss:ss, all:all, filt:filt, fd:fd, gids:gids, cols:cols };
}

// ─── INDIVIDUAL PAGE BUILDERS ─────────────────────────────────

function buildOverview() {
  var d = _dbPrepare();
  d.ss.toast('Building Overview…', '📊', -1);
  _pgOverview(d.ss, d.fd, d.all, d.filt, d.gids, d.cols);
  d.ss.toast('✅ Overview done', '📊', 5);
}

function buildFinancial() {
  var d = _dbPrepare();
  d.ss.toast('Building Financial…', '💰', -1);
  _pgFinancial(d.ss, d.fd, d.all, d.filt, d.gids, d.cols);
  d.ss.toast('✅ Financial done', '💰', 5);
}

function buildProjects() {
  var d = _dbPrepare();
  d.ss.toast('Building Projects…', '🏗️', -1);
  _pgProjects(d.ss, d.fd, d.all, d.filt, d.gids, d.cols);
  d.ss.toast('✅ Projects done', '🏗️', 5);
}

function buildWorkforce() {
  var d = _dbPrepare();
  d.ss.toast('Building Workforce…', '👷', -1);
  _pgWorkforce(d.ss, d.fd, d.all, d.filt, d.gids, d.cols);
  d.ss.toast('✅ Workforce done', '👷', 5);
}

function buildOperations() {
  var d = _dbPrepare();
  d.ss.toast('Building Operations…', '⚙️', -1);
  _pgOperations(d.ss, d.fd, d.all, d.filt, d.gids, d.cols);
  d.ss.toast('✅ Operations done', '⚙️', 5);
}

function buildHSE() {
  var d = _dbPrepare();
  d.ss.toast('Building HSE & Safety…', '🛡️', -1);
  _pgHSE(d.ss, d.fd, d.all, d.filt, d.gids, d.cols);
  d.ss.toast('✅ HSE & Safety done', '🛡️', 5);
}

/** Refresh all 6 pages in sequence (use only if time allows) */
function refreshDashboard() {
  var d = _dbPrepare();
  d.ss.toast('Refreshing all pages…', '🔄', -1);
  _pgOverview  (d.ss, d.fd, d.all, d.filt, d.gids, d.cols);
  _pgFinancial (d.ss, d.fd, d.all, d.filt, d.gids, d.cols);
  _pgProjects  (d.ss, d.fd, d.all, d.filt, d.gids, d.cols);
  _pgWorkforce (d.ss, d.fd, d.all, d.filt, d.gids, d.cols);
  _pgOperations(d.ss, d.fd, d.all, d.filt, d.gids, d.cols);
  _pgHSE       (d.ss, d.fd, d.all, d.filt, d.gids, d.cols);
  d.ss.toast('✅ All pages refreshed  ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm'), '📊', 6);
}

// ─── DATA LAYER ──────────────────────────────────────────────

function _dbLoad(ss) {
  function rows(name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() <= 1) return [];
    var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    var h = v[0], out = [];
    for (var r = 1; r < v.length; r++) {
      if (!String(v[r][0]).trim()) continue;
      var o = {}; for (var c = 0; c < h.length; c++) o[h[c]] = v[r][c];
      out.push(o);
    }
    return out;
  }
  return {
    prj: rows('Projects'),  con: rows('Contracts'),
    pay: rows('Payment_Applications'), vo: rows('Variation_Orders'),
    po:  rows('Purchase_Orders'),  ph: rows('Project_Phases'),
    emp: rows('Employees'),  ctr: rows('Contractors'),
    sup: rows('Suppliers'),  eq:  rows('Equipment'),
    inc: rows('Safety_Incidents'), ins: rows('Inspections'),
    pmt: rows('Permits_Approvals'), dsr: rows('Daily_Site_Reports'),
    doc: rows('Project_Documents'), cli: rows('Clients'),
    wp:  rows('Work_Packages')
  };
}

function _dbGetFilt(ss) {
  var fs = ss.getSheetByName(DB.FS);
  if (!fs) return { proj:'All', type:'All', stat:'All' };
  var v = fs.getRange(2, 2, 1, 3).getValues()[0];
  return {
    proj: String(v[0]||'').trim() || 'All',
    type: String(v[1]||'').trim() || 'All',
    stat: String(v[2]||'').trim() || 'All'
  };
}

function _dbFilt(all, f) {
  var pp = all.prj.filter(function(p) {
    if (f.proj !== 'All' && p.project_id !== f.proj && p.project_name !== f.proj) return false;
    if (f.type !== 'All' && p.project_type !== f.type) return false;
    if (f.stat !== 'All' && p.status       !== f.stat) return false;
    return true;
  });
  var ids = {}; pp.forEach(function(p) { ids[p.project_id] = 1; });
  var full = pp.length === all.prj.length;
  function fk(arr, k) { return full ? arr : arr.filter(function(r){ return ids[r[k||'project_id']]; }); }
  return {
    prj: pp, con: fk(all.con), pay: fk(all.pay),
    vo: fk(all.vo), po: fk(all.po), ph: fk(all.ph),
    inc: fk(all.inc), ins: fk(all.ins), pmt: fk(all.pmt),
    dsr: fk(all.dsr), doc: fk(all.doc), wp: fk(all.wp),
    emp: all.emp, ctr: all.ctr, sup: all.sup, eq: all.eq, cli: all.cli
  };
}

function _dbGids(ss) {
  var m = {};
  DB.PAGES.forEach(function(p) {
    var sh = ss.getSheetByName(p.s);
    if (sh) m[p.k] = sh.getSheetId();
  });
  return m;
}

// ─── COLUMN MAP ──────────────────────────────────────────────

/** Returns {sheetName: {colName: colLetter}} for all 17 data sheets.
 *  Called once per _dbPrepare() so all page builders share the same map. */
function _dbColMap(ss) {
  var SHEETS = [
    'Projects','Contracts','Employees','Contractors','Suppliers',
    'Equipment','Safety_Incidents','Inspections','Permits_Approvals',
    'Payment_Applications','Variation_Orders','Purchase_Orders',
    'Daily_Site_Reports','Project_Phases','Work_Packages','Clients'
  ];
  var m = {};
  SHEETS.forEach(function(sn) {
    var sh = ss.getSheetByName(sn);
    m[sn] = {};
    if (!sh || sh.getLastRow() < 1) return;
    var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    hdrs.forEach(function(h, i) { if (h) m[sn][String(h)] = colL(i + 1); });
  });
  return m;
}

/** Safe range ref — "Sheet!Col2:Col" or "A1:A1" (returns 0) if field missing */
function _cr(cols, sn, f) {
  var l = (cols[sn] || {})[f];
  return l ? sn + '!' + l + '2:' + l : 'A1:A1';
}

/** Write formula-based chart data to hidden rows; returns the bound Range.
 *  labels  — array of row-label strings
 *  fmls    — array of formula strings (one per label)
 *  Returns the range spanning header + all data rows (2 cols wide). */
function _dbFData(sh, startRow, hdr0, hdr1, labels, fmls) {
  sh.getRange(startRow, 1).setValue(hdr0);
  sh.getRange(startRow, 2).setValue(hdr1);
  for (var i = 0; i < labels.length; i++) {
    sh.getRange(startRow + 1 + i, 1).setValue(labels[i]);
    sh.getRange(startRow + 1 + i, 2).setFormula(fmls[i]);
  }
  return sh.getRange(startRow, 1, labels.length + 1, 2);
}

// ─── SHEET SETUP ─────────────────────────────────────────────

function _dbSetupFS(ss) {
  var fs = ss.getSheetByName(DB.FS) || ss.insertSheet(DB.FS);
  fs.clearContents(); fs.clearFormats();
  fs.setTabColor('#6366f1');
  fs.getRange(1,1,1,6).setValues([['','Project','Type','Status','','Matched_PIDs']])
    .setBackground(DB.C.HDR).setFontColor('#ffffff').setFontWeight('bold');
  fs.getRange(2,2).setValue('All');
  fs.getRange(2,3).setValue('All');
  fs.getRange(2,4).setValue('All');

  // Build matched project ID list via FILTER (col F)
  var cols = _dbColMap(ss);
  var pnC = (cols['Projects']||{})['project_name'] || 'C';
  var ptC = (cols['Projects']||{})['project_type'] || 'D';
  var psC = (cols['Projects']||{})['status']       || 'G';
  fs.getRange(2,6).setFormula(
    '=IFERROR(FILTER(Projects!A2:A,'
    + 'IF(B2="All",TRUE,Projects!' + pnC + '2:' + pnC + '=B2),'
    + 'IF(C2="All",TRUE,Projects!' + ptC + '2:' + ptC + '=C2),'
    + 'IF(D2="All",TRUE,Projects!' + psC + '2:' + psC + '=D2)),"")');

  var types = ['All','Commercial Office Building','High-Rise Residential Tower',
    'Hotel & Hospitality','Industrial Warehouse','Infrastructure - Bridges',
    'Infrastructure - Roads','Marine & Coastal Works','Mixed-Use Development',
    'Retail Mall','School/Educational','Utility Infrastructure','Villa Development'];
  var stats = ['All','Awarded','In Progress','On Hold','Delayed','Completed','Cancelled'];
  fs.getRange(2,3).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(types,true).setAllowInvalid(false).build());
  fs.getRange(2,4).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(stats,true).setAllowInvalid(false).build());
  fs.autoResizeColumns(1,6);
  // DO NOT hide — formulas reference this sheet
}

function _dbSetupPage(ss, key) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var p = DB.PAGES.filter(function(x){ return x.k === key; })[0];
  if (!p) return null;
  var sh = ss.getSheetByName(p.s) || ss.insertSheet(p.s);
  sh.clearContents(); sh.clearFormats();
  sh.getCharts().forEach(function(c){ sh.removeChart(c); });
  sh.setHiddenGridlines(true);
  sh.setTabColor(DB.C.IND);
  // Columns — batch call
  sh.setColumnWidth(1, 16);
  sh.setColumnWidths(2, 24, 85);
  sh.setColumnWidth(26, 16);
  // Key rows
  sh.setRowHeight(1, 4);
  sh.setRowHeight(2, 44);
  sh.setRowHeight(3, 3);
  sh.setRowHeight(4, 36);
  sh.setRowHeight(5, 4);
  sh.setRowHeight(6, 36);
  sh.setRowHeight(7, 6);
  return sh;
}

// ─── BACKGROUND CANVAS ───────────────────────────────────────

/** Paint entire visible area with PAGE_BG — must be called first */
function _dbPaintBg(sh) {
  sh.getRange(1, 1, 400, 27).setBackground(DB.C.PAGE)
    .setFontFamily(DB.FONT).setFontColor(DB.C.TXT);
}

// ─── HEADER  (row 2) ─────────────────────────────────────────

function _dbHdr(sh, pageIco, pageTitle) {
  var C = DB.C;
  // Full row navy
  sh.getRange(2, 1, 1, 27).setBackground(C.HDR);
  // Brand — left half
  sh.getRange(2, 2, 1, 12).merge()
    .setValue('  🏗️  DUBAI ERP  ·  Project Management')
    .setBackground(C.HDR).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(13).setFontFamily(DB.FONT)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  // Page title — right half
  sh.getRange(2, 14, 1, 12).merge()
    .setValue(pageIco + '  ' + pageTitle + '   ')
    .setBackground(C.HDR).setFontColor('#a5b4fc')
    .setFontWeight('bold').setFontSize(13).setFontFamily(DB.FONT)
    .setVerticalAlignment('middle').setHorizontalAlignment('right');
  // Accent line
  sh.getRange(2, 2, 1, 24)
    .setBorder(null,null,true,null,null,null, C.IND, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

// ─── NAV BAR  (row 4) ────────────────────────────────────────

function _dbNav(sh, gids, activeKey) {
  var C = DB.C;
  // Base: white nav bar with subtle bottom border
  sh.getRange(4, 1, 1, 27).setBackground(C.NAV_BG)
    .setBorder(null,null,true,null,null,null, C.BD, SpreadsheetApp.BorderStyle.SOLID);
  // 6 tabs × 4 cols = 24 cols (B–Y)
  DB.PAGES.forEach(function(p, i) {
    var col = 2 + i * 4;
    var gid = gids[p.k] || 0;
    var lbl = '  ' + p.i + '  ' + p.l + '  ';
    var isActive = p.k === activeKey;
    var rng = sh.getRange(4, col, 1, 4).merge();
    rng.setFormula('=HYPERLINK("#gid=' + gid + '","' + lbl.replace(/"/g,"'") + '")')
      .setBackground(isActive ? C.NAV_ACT : C.NAV_BG)
      .setFontColor(isActive ? '#ffffff' : C.NAV_T)
      .setFontWeight(isActive ? 'bold' : 'normal')
      .setFontSize(10).setFontFamily(DB.FONT)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    if (isActive) {
      rng.setBorder(null,null,true,null,null,null, C.IND, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  });
}

// ─── FILTER BAR  (row 6) ─────────────────────────────────────

function _dbFiltBar(sh, filt, all) {
  var C = DB.C, row = DB.FROW, FC = DB.FCOLS;

  // Panel background — soft indigo tint
  sh.getRange(row, 2, 1, 24).setBackground(C.IND_L)
    .setBorder(true,true,true,true,false,false, '#c7d2fe', SpreadsheetApp.BorderStyle.SOLID);
  // Left accent bar
  sh.getRange(row, 2, 1, 1)
    .setBorder(null,true,null,null,null,null, C.IND, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  function lbl(col, txt, bold) {
    sh.getRange(row, col).setValue(txt)
      .setBackground(C.IND_L).setFontColor(bold ? '#4338ca' : C.MU)
      .setFontWeight(bold ? 'bold' : 'normal').setFontSize(9)
      .setFontFamily(DB.FONT).setVerticalAlignment('middle');
  }

  lbl(3,  '🔍  FILTERS', true);
  lbl(5,  'Project', true);
  lbl(10, 'Type', true);
  lbl(16, 'Status', true);

  // Dropdown cells — white pill with indigo border
  function pill(col, val) {
    sh.getRange(row, col, 1, 2).merge()
      .setValue(val)
      .setBackground('#ffffff').setFontColor('#4338ca')
      .setFontWeight('bold').setFontSize(9).setFontFamily(DB.FONT)
      .setVerticalAlignment('middle').setHorizontalAlignment('center')
      .setBorder(true,true,true,true,null,null, '#c7d2fe', SpreadsheetApp.BorderStyle.SOLID);
  }
  pill(FC.proj, filt.proj);
  pill(FC.type, filt.type);
  pill(FC.stat, filt.stat);

  // Refresh hint
  lbl(22, '🔄 Refresh via menu', false);

  // Data validation (project names from live data)
  var prjList = ['All'].concat(
    all.prj.map(function(p){ return p.project_name||''; })
      .filter(function(v,i,a){ return v && a.indexOf(v)===i; }).sort().slice(0,200));
  var types = ['All','Commercial Office Building','High-Rise Residential Tower',
    'Hotel & Hospitality','Industrial Warehouse','Infrastructure - Bridges',
    'Infrastructure - Roads','Marine & Coastal Works','Mixed-Use Development',
    'Retail Mall','School/Educational','Utility Infrastructure','Villa Development'];
  var stats = ['All','Awarded','In Progress','On Hold','Delayed','Completed','Cancelled'];
  sh.getRange(row, FC.proj, 1, 2).merge()
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(prjList,true).setAllowInvalid(true).build());
  sh.getRange(row, FC.type, 1, 2).merge()
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(types,true).setAllowInvalid(false).build());
  sh.getRange(row, FC.stat, 1, 2).merge()
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(stats,true).setAllowInvalid(false).build());
}

// ─── KPI CARD ────────────────────────────────────────────────

/**
 * KPI card — 5 cols wide, 4 data rows + 1 gap = 5 rows total
 * Draws at (R, col).  Row heights must be pre-set via _dbKpiRows().
 * ck = 'G'|'B'|'A'|'R'|'P'|'T'|'I'
 */
function _dbKpi(sh, R, col, ico, val, lbl, sub, ck) {
  var C  = DB.C;
  var fg = C[ck]  || C.I;
  var bg = C[ck+'L'] || C.IL;
  var W  = 5;

  // ① Accent stripe (thin colored top bar)
  sh.getRange(R, col, 1, W).setBackground(fg);

  // ② Card body: white panel with soft border
  sh.getRange(R+1, col, 3, W).setBackground(C.PANEL)
    .setBorder(false, true, true, true, false, false, C.BD, SpreadsheetApp.BorderStyle.SOLID);

  // ③ Icon cell (merged vertically — pastel pill background)
  sh.getRange(R+1, col, 3, 1).merge()
    .setValue(ico).setBackground(bg)
    .setFontSize(20).setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ④ Label — dark muted uppercase
  sh.getRange(R+1, col+1, 1, W-1).merge()
    .setValue('  ' + lbl.toUpperCase())
    .setBackground(C.PANEL).setFontColor(C.MU)
    .setFontSize(8).setFontWeight('bold').setFontFamily(DB.FONT)
    .setVerticalAlignment('bottom');

  // ⑤ Value — dark text, supports formula strings
  var vRng = sh.getRange(R+2, col+1, 1, W-1).merge()
    .setBackground(C.PANEL).setFontColor(C.TXT)
    .setFontSize(22).setFontWeight('bold').setFontFamily(DB.FONT)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  if (typeof val === 'string' && val.charAt(0) === '=') {
    vRng.setFormula(val);
  } else {
    vRng.setValue('  ' + String(val));
  }

  // ⑥ Sub-label — muted gray, supports formula strings
  var sRng = sh.getRange(R+3, col+1, 1, W-1).merge()
    .setBackground(C.PANEL).setFontColor(C.MU)
    .setFontSize(8).setFontFamily(DB.FONT).setVerticalAlignment('top');
  if (typeof sub === 'string' && sub.charAt(0) === '=') {
    sRng.setFormula(sub);
  } else {
    sRng.setValue('  ' + sub);
  }
}

/** Set KPI row-group heights for the accent+card+gap structure */
function _dbKpiRows(sh, R) {
  sh.setRowHeight(R,   3);
  sh.setRowHeight(R+1, 18);
  sh.setRowHeight(R+2, 34);
  sh.setRowHeight(R+3, 14);
  sh.setRowHeight(R+4, 8);  // gap
}

// ─── SECTION HEADER ──────────────────────────────────────────

function _dbSec(sh, row, col, span, label, cnt) {
  var C = DB.C;
  var badge = cnt !== undefined ? '   ●  ' + cnt + ' records' : '';
  sh.getRange(row, col, 1, span).merge()
    .setValue('  ' + label + badge)
    .setBackground(C.SEC_BG).setFontColor(C.SEC_T)
    .setFontWeight('bold').setFontSize(10).setFontFamily(DB.FONT)
    .setVerticalAlignment('middle');
  sh.getRange(row, col, 1, 1)
    .setBorder(null,true,null,null,null,null, C.IND, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(row, col, 1, span)
    .setBorder(null,null,true,null,null,null, C.BD, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 28);
  return row + 1;
}

// ─── TABLE HEADER ────────────────────────────────────────────

function _dbTH(sh, row, col, hdrs) {
  var C = DB.C;
  for (var i = 0; i < hdrs.length; i++) {
    sh.getRange(row, col+i)
      .setValue(hdrs[i]).setBackground(C.TH).setFontColor(C.TH_T)
      .setFontWeight('bold').setFontSize(9).setFontFamily(DB.FONT)
      .setVerticalAlignment('middle').setWrap(false);
  }
  sh.getRange(row, col, 1, hdrs.length)
    .setBorder(false,false,true,false,false,false, C.BD, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 30);
  return row + 1;
}

// ─── TABLE ROW ───────────────────────────────────────────────

function _dbTR(sh, row, col, vals, alt) {
  var C = DB.C, bg = alt ? C.TR2 : C.TR1;
  for (var i = 0; i < vals.length; i++) {
    var rng = sh.getRange(row, col+i);
    rng.setValue(vals[i]).setBackground(bg)
      .setFontSize(9).setFontFamily(DB.FONT)
      .setFontColor(i === 0 ? C.IND : C.TXT2)
      .setFontWeight(i === 0 ? 'bold' : 'normal')
      .setVerticalAlignment('middle');
  }
  sh.getRange(row, col, 1, vals.length)
    .setBorder(false,false,true,false,false,false, '#f1f5f9', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(row, 22);
  return row + 1;
}

// ─── STATUS BADGE ────────────────────────────────────────────

function _dbBadge(sh, row, col, val) {
  var s = ERP.COLORS.STATUS[val];
  if (!s) return;
  sh.getRange(row, col).setBackground(s.bg).setFontColor(s.fg).setFontWeight('bold');
}

// ─── SPACER ROW ──────────────────────────────────────────────

function _dbSp(sh, row, h) {
  sh.getRange(row, 2, 1, 24).setBackground(DB.C.PAGE);
  sh.setRowHeight(row, h || 8);
  return row + 1;
}

// ─── CHART HELPERS ───────────────────────────────────────────

/** Apply shared light-theme styling to a chart builder */
function _dbChartStyle(bld, title, w, h, colors) {
  bld.setOption('title', title)
    .setOption('width',  w || 440)
    .setOption('height', h || 250)
    .setOption('fontName', 'Inter')
    .setOption('backgroundColor',  { fill:'#ffffff' })
    .setOption('titleTextStyle',   { color:'#1e293b', fontSize:11, bold:true })
    .setOption('vAxis',  { textStyle:{ color:'#64748b', fontSize:8 }, gridlines:{ color:'#f1f5f9', count:5 }, baselineColor:'#e2e8f0', minorGridlines:{ count:0 } })
    .setOption('hAxis',  { textStyle:{ color:'#64748b', fontSize:8 }, gridlines:{ color:'#f1f5f9' } })
    .setOption('legend', { textStyle:{ color:'#475569', fontSize:9 }, position:'bottom' })
    .setOption('chartArea', { backgroundColor:'#ffffff', left:50, top:30, width:'75%', height:'65%' })
    .setOption('bar',    { groupWidth:'55%' })
    .setOption('pieHole', 0.5)
    .setOption('pieSliceTextStyle', { color:'#475569', fontSize:9 })
    .setOption('tooltip', { textStyle:{ fontSize:10 } });
  if (colors && colors.length) bld.setOption('colors', colors);
  return bld;
}

/** Chart from formula range already written in the sheet (preferred — live data) */
function _dbChartF(sh, type, title, rng, anchorRow, anchorCol, w, h, colors) {
  var bld = sh.newChart()
    .setChartType(type).addRange(rng).setNumHeaders(1)
    .setPosition(anchorRow, anchorCol, 0, 0);
  sh.insertChart(_dbChartStyle(bld, title, w, h, colors).build());
}

/** Chart from inline data array (fallback — writes values to dataRow first) */
function _dbChart(sh, type, title, data, dataRow, anchorRow, anchorCol, w, h, colors) {
  if (!data || data.length < 2) return;
  sh.getRange(dataRow, 1, data.length, data[0].length).setValues(data);
  var rng = sh.getRange(dataRow, 1, data.length, data[0].length);
  var bld = sh.newChart()
    .setChartType(type).addRange(rng).setNumHeaders(1)
    .setPosition(anchorRow, anchorCol, 0, 0);
  sh.insertChart(_dbChartStyle(bld, title, w, h, colors).build());
}

// ─── FORMULA STRING HELPERS (inject live formulas into KPI cells) ────────────

/** Returns a formula string that formats a number as "AED 1.23M" */
function _aedF(f) {
  var v = '(' + f + ')';
  return '=IFERROR(IF(' + v + '>=1E9,"AED "&TEXT(' + v + '/1E9,"0.00")&"B",'
       + 'IF(' + v + '>=1E6,"AED "&TEXT(' + v + '/1E6,"0.00")&"M",'
       + 'IF(' + v + '>=1E3,"AED "&TEXT(' + v + '/1E3,"0.0")&"K",'
       +        '"AED "&TEXT(' + v + ',"#,##0")))),"AED 0")';
}

/** Returns a formula string that formats a number as "42.5%" */
function _pctF(f) {
  return '=IFERROR(TEXT((' + f + '),"0.0")&"%","0.0%")';
}

/** Returns a formula string that formats a number with thousands separator */
function _numF(f) {
  return '=IFERROR(TEXT((' + f + '),"#,##0"),"—")';
}

// ─── FILTER FORMULA HELPERS ──────────────────────────────────

/**
 * Returns COUNTIFS filter criteria suffix for the Projects sheet.
 * Adds conditions for project_name, project_type, status
 * referencing _DB_Filters!B2, C2, D2 dynamically.
 * Usage: _numF('COUNTIFS('+PS+',"In Progress"'+_fp(cols)+')')
 */
function _fp(cols) {
  var pn = _cr(cols,'Projects','project_name');
  var pt = _cr(cols,'Projects','project_type');
  var ps = _cr(cols,'Projects','status');
  return ','+pn+',IF(_DB_Filters!B2="All","*",_DB_Filters!B2)'
       + ','+pt+',IF(_DB_Filters!C2="All","*",_DB_Filters!C2)'
       + ','+ps+',IF(_DB_Filters!D2="All","*",_DB_Filters!D2)';
}

/**
 * Returns a COUNTIF match expression for related sheets.
 * Checks if the row's project_id is in the filtered list (_DB_Filters!$F:$F).
 * For use inside SUMPRODUCT:  SUMPRODUCT(criteria * values * _fm(cols,'SheetName'))
 */
function _fm(cols, sn) {
  var pid = _cr(cols, sn, 'project_id');
  return '(COUNTIF(_DB_Filters!$F:$F,' + pid + ')>0)';
}

// ─── FORMATTERS ──────────────────────────────────────────────

function _aed(n) {
  n = parseFloat(n)||0;
  if (n >= 1e9) return 'AED ' + (n/1e9).toFixed(2) + 'B';
  if (n >= 1e6) return 'AED ' + (n/1e6).toFixed(2) + 'M';
  if (n >= 1e3) return 'AED ' + (n/1e3).toFixed(1) + 'K';
  return 'AED ' + Math.round(n);
}
function _pct(n)  { return (parseFloat(n)||0).toFixed(1) + '%'; }
function _num(n)  { return String(Math.round(parseFloat(n)||0)); }
function _days(d) {
  if (!d) return '—';
  try { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'dd/MM/yy'); }
  catch(e) { return '—'; }
}
function _pctComp(p) {
  var v = parseFloat(p.completion_pct)||0;
  return v <= 1 ? v*100 : v;
}

// ─── AGGREGATORS ─────────────────────────────────────────────

function _sum(a, f)  { return a.reduce(function(s,r){ return s+(parseFloat(r[f])||0); }, 0); }
function _avg(a, f)  { return a.length ? _sum(a,f)/a.length : 0; }
function _cnt(a, f, v) { return v===undefined ? a.length : a.filter(function(r){ return r[f]===v; }).length; }
function _grpCnt(a, f) {
  var m={};
  a.forEach(function(r){ var k=r[f]||'Other'; m[k]=(m[k]||0)+1; });
  return m;
}
function _grpSum(a, kf, vf) {
  var m={};
  a.forEach(function(r){ var k=r[kf]||'Other'; m[k]=(m[k]||0)+(parseFloat(r[vf])||0); });
  return m;
}
function _top(obj, n, asc) {
  return Object.keys(obj).map(function(k){ return [k, obj[k]]; })
    .sort(function(a,b){ return asc ? a[1]-b[1] : b[1]-a[1]; }).slice(0, n||10);
}

// ─── MENU & ONEDIT HOOK ──────────────────────────────────────

function addDashboardMenu(menu) {
  return menu
    .addSeparator()
    .addItem('🏗️  Step 1 — Create All Sheets  (run once)',  'buildDashboard')
    .addSeparator()
    .addItem('📊  Build Overview Page',    'buildOverview')
    .addItem('💰  Build Financial Page',   'buildFinancial')
    .addItem('🏗️  Build Projects Page',    'buildProjects')
    .addItem('👷  Build Workforce Page',   'buildWorkforce')
    .addItem('⚙️   Build Operations Page', 'buildOperations')
    .addItem('🛡️  Build HSE & Safety Page','buildHSE')
    .addSeparator()
    .addItem('🔄  Refresh All Pages',      'refreshDashboard');
}

function _dbOnEdit(e) {
  try {
    if (!e || !e.range) return false;
    var sh    = e.range.getSheet();
    var sName = sh.getName();
    var isDash = DB.PAGES.some(function(p){ return p.s === sName; });
    var isFltSh = sName === DB.FS;
    if (!isDash && !isFltSh) return false;

    var row = e.range.getRow(), col = e.range.getColumn();
    var FC = DB.FCOLS;
    var isFRow = isDash && row === DB.FROW && (col === FC.proj || col === FC.type || col === FC.stat);
    var isFSh  = isFltSh && row === 2 && col >= 2 && col <= 4;
    if (!isFRow && !isFSh) return false;

    // Sync from dashboard page → _DB_Filters (lightweight — no rebuild)
    if (isDash && isFRow) {
      var fs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DB.FS);
      if (fs) {
        var proj = sh.getRange(DB.FROW, FC.proj).getValue() || 'All';
        var type = sh.getRange(DB.FROW, FC.type).getValue() || 'All';
        var stat = sh.getRange(DB.FROW, FC.stat).getValue() || 'All';
        fs.getRange(2,2).setValue(proj);
        fs.getRange(2,3).setValue(type);
        fs.getRange(2,4).setValue(stat);
      }
    }
    // Formulas auto-update; toast tells user tables need manual refresh
    SpreadsheetApp.getActiveSpreadsheet()
      .toast('KPIs & charts updated. Use menu → Refresh to update tables.', '🔄', 5);
    return true;
  } catch(err) {
    SpreadsheetApp.getActiveSpreadsheet()
      .toast('Dashboard: ' + err.message, '⚠️', 6);
    return true;
  }
}

