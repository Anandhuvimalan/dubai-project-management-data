/**
 * db_operations.gs — Operations & Procurement
 * KPIs use formula-based filters via _fm() for project-linked sheets
 * Charts: formula-based
 */
function _pgOperations(ss, fd, all, filt, gids, cols) {
  var sh = _dbSetupPage(ss, 'operations');
  _dbPaintBg(sh);
  _dbHdr(sh, '⚙️', 'Operations & Procurement');
  _dbNav(sh, gids, 'operations');
  _dbFiltBar(sh, filt, all);

  var EqS  = _cr(cols,'Equipment','status');
  var EqO  = _cr(cols,'Equipment','ownership');
  var EqR  = _cr(cols,'Equipment','daily_rate_aed');
  var POs  = _cr(cols,'Purchase_Orders','status');
  var POv  = _cr(cols,'Purchase_Orders','amount_aed');
  var FMpo = _fm(cols,'Purchase_Orders');

  /* ── KPI Row 1 ─────────────────────────────────────────── */
  _dbKpiRows(sh, 8);
  _dbKpi(sh, 8,  2, '🚜',
    _numF('COUNTA(Equipment!A2:A)'),
    'Total Equipment',
    '=IFERROR(TEXT(COUNTIF('+EqS+',"In Use"),"0")&" in use","—")',
    'B');
  _dbKpi(sh, 8,  8, '🔧',
    _numF('COUNTIF('+EqS+',"In Use")'),
    'Units In Use',
    'Currently deployed',
    'G');
  _dbKpi(sh, 8, 14, '🔩',
    _numF('COUNTIF('+EqS+',"Under Maintenance")'),
    'Under Maintenance',
    'Offline for service',
    'A');
  _dbKpi(sh, 8, 20, '💸',
    _aedF('SUMIFS('+EqR+','+EqO+',"Rented",'+EqS+',"In Use")'),
    'Daily Rental Cost',
    '=IFERROR(TEXT(COUNTIFS('+EqO+',"Rented",'+EqS+',"In Use"),"0")&" rentals","—")',
    'R');

  /* ── KPI Row 2 ─────────────────────────────────────────── */
  _dbKpiRows(sh, 13);
  _dbKpi(sh, 13,  2, '📦',
    _numF('SUMPRODUCT(('+POs+'="Draft")*'+FMpo+')+SUMPRODUCT(('+POs+'="Issued")*'+FMpo+')'),
    'Pending POs',
    'Draft + Issued (filtered)',
    'A');
  _dbKpi(sh, 13,  8, '💰',
    _aedF('SUMPRODUCT('+POv+'*'+FMpo+')'),
    'Total PO Value',
    '=IFERROR(TEXT(SUMPRODUCT('+FMpo+'*1),"#,##0")&" orders","—")',
    'B');
  _dbKpi(sh, 13, 14, '✅',
    _aedF('SUMPRODUCT(('+POs+'="Issued")*'+POv+'*'+FMpo+')+SUMPRODUCT(('+POs+'="Delivered")*'+POv+'*'+FMpo+')'),
    'Issued / Delivered',
    'In delivery pipeline',
    'G');
  _dbKpi(sh, 13, 20, '🔄',
    _pctF('IFERROR(COUNTIF('+EqO+',"Rented")/COUNTA(Equipment!A2:A)*100,0)'),
    'Rented Equipment %',
    '=IFERROR(TEXT(COUNTIF('+EqO+',"Rented"),"0")&" rented","—")',
    'T');

  _dbSp(sh, 18, 10);
  var R = 19;
  var C = DB.C;

  /* ── Charts ─────────────────────────────────────────────── */
  var CD  = 400;
  var CD2 = 410;
  var eqSt = ['Available','In Use','Under Maintenance','Decommissioned','In Transit'];
  var r1 = _dbFData(sh, CD, 'Status', 'Count', eqSt,
    eqSt.map(function(s){ return '=COUNTIF('+EqS+',"'+s+'")'; }));
  var poSt = ['Draft','Issued','Delivered','Cancelled','Pending Approval'];
  var r2 = _dbFData(sh, CD2, 'Status', 'AED', poSt,
    poSt.map(function(s){ return '=IFERROR(SUMPRODUCT(('+POs+'="'+s+'")*'+POv+'*'+FMpo+'),0)'; }));

  sh.getRange(R, 2, 1, 12).merge().setValue('  📊  EQUIPMENT FLEET')
    .setBackground(C.SEC_BG).setFontColor(C.SEC_T).setFontWeight('bold').setFontSize(10).setFontFamily(DB.FONT).setVerticalAlignment('middle');
  sh.getRange(R, 2, 1, 1).setBorder(null,true,null,null,null,null, C.IND, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(R, 14, 1, 12).merge().setValue('  📊  PO VALUE BY STATUS')
    .setBackground(C.SEC_BG).setFontColor(C.SEC_T).setFontWeight('bold').setFontSize(10).setFontFamily(DB.FONT).setVerticalAlignment('middle');
  sh.getRange(R, 14, 1, 1).setBorder(null,true,null,null,null,null, C.IND, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(R, 28); R++;

  _dbChartF(sh, Charts.ChartType.PIE, 'Equipment Fleet', r1, R, 2, 440, 250,
    ['#16a34a','#2563eb','#d97706','#94a3b8','#9333ea']);
  _dbChartF(sh, Charts.ChartType.COLUMN, 'PO Value by Status', r2, R, 14, 440, 250,
    ['#94a3b8','#2563eb','#16a34a','#dc2626','#d97706']);
  R += 14;
  _dbSp(sh, R++, 10);

  /* ── Equipment Rental Alerts ───────────────────────────── */
  var eqRented = all.eq.filter(function(r){ return String(r.ownership).toLowerCase()==='rented'; });
  var topEq = eqRented.slice().sort(function(a,b){ return (parseFloat(b.daily_rate_aed)||0)-(parseFloat(a.daily_rate_aed)||0); }).slice(0,10);

  R = _dbSec(sh, R, 2, 24, '  🚜  EQUIPMENT RENTAL ALERTS', topEq.length);
  R = _dbTH(sh, R, 2, ['Equipment','Model','Status','Daily Rate','Project','Alert']);
  topEq.forEach(function(eq, i){
    var rate = parseFloat(eq.daily_rate_aed)||0;
    R = _dbTR(sh, R, 2, [
      eq.equipment_type||'—', eq.model||'—', eq.status||'—',
      _aed(rate)+'/day', eq.current_project||'—',
      rate>5000?'🔴 HIGH':rate>2000?'🟡 MED':'🟢 LOW'
    ], i%2===1);
  });

  try { sh.hideRows(400, 20); } catch(e) {}
  SpreadsheetApp.flush();
}
