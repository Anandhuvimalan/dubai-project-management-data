/**
 * db_overview.gs — Portfolio Overview
 * All KPIs use COUNTIFS/SUMPRODUCT with dynamic filter refs to _DB_Filters
 * Charts use formula-based data with filter criteria
 */
function _pgOverview(ss, fd, all, filt, gids, cols) {
  var sh = _dbSetupPage(ss, 'overview');
  _dbPaintBg(sh);
  _dbHdr(sh, '📊', 'Portfolio Overview');
  _dbNav(sh, gids, 'overview');
  _dbFiltBar(sh, filt, all);

  var PS  = _cr(cols,'Projects','status');
  var PCp = _cr(cols,'Projects','completion_pct');
  var CV  = _cr(cols,'Contracts','contract_value_aed');
  var IcS = _cr(cols,'Safety_Incidents','status');
  var InR = _cr(cols,'Inspections','result');
  var PmS = _cr(cols,'Permits_Approvals','status');
  var PmE = _cr(cols,'Permits_Approvals','expiry_date');
  var EqS = _cr(cols,'Equipment','status');
  var FP  = _fp(cols);  // Projects filter criteria
  var FMc = _fm(cols,'Contracts');
  var FMi = _fm(cols,'Safety_Incidents');
  var FMn = _fm(cols,'Inspections');
  var FMp = _fm(cols,'Permits_Approvals');

  /* ── KPI Row 1 — filtered formulas ─────────────────────── */
  _dbKpiRows(sh, 8);
  _dbKpi(sh, 8,  2, '🏗️',
    _numF('COUNTIFS('+PS+',"In Progress"'+FP+')'),
    'Active Projects',
    '=IFERROR(TEXT(COUNTIFS('+PS+',"<>"&""'+FP+'),"#,##0")&" in portfolio","—")',
    'G');
  _dbKpi(sh, 8,  8, '⚠️',
    _numF('COUNTIFS('+PS+',"Delayed"'+FP+')+COUNTIFS('+PS+',"On Hold"'+FP+')'),
    'Delayed / On Hold',
    'Requires attention',
    'R');
  _dbKpi(sh, 8, 14, '💰',
    _aedF('SUMPRODUCT('+CV+'*'+FMc+')'),
    'Total Contract Value',
    '=IFERROR(TEXT(SUMPRODUCT('+FMc+'*1),"#,##0")&" contracts","—")',
    'B');
  _dbKpi(sh, 8, 20, '📈',
    _pctF('IFERROR(AVERAGEIFS('+PCp+','+PS+',"<>"&""'+FP+'),0)'),
    'Portfolio Completion',
    'Avg across filtered projects',
    'P');

  /* ── KPI Row 2 ─────────────────────────────────────────── */
  _dbKpiRows(sh, 13);
  _dbKpi(sh, 13,  2, '🚨',
    _numF('SUMPRODUCT(('+IcS+'="Open")*'+FMi+')+SUMPRODUCT(('+IcS+'="Under Investigation")*'+FMi+')'),
    'Open Safety Incidents',
    'HSE cases in progress',
    'R');
  _dbKpi(sh, 13,  8, '✅',
    _pctF('IFERROR(SUMPRODUCT(('+InR+'="Pass")*'+FMn+')/SUMPRODUCT('+FMn+'*1)*100,0)'),
    'Inspection Pass Rate',
    '=IFERROR(TEXT(SUMPRODUCT('+FMn+'*1),"#,##0")&" total","—")',
    'G');
  _dbKpi(sh, 13, 14, '📋',
    _numF('SUMPRODUCT(('+PmE+'>TODAY())*('+PmE+'<(TODAY()+30))*('+PmS+'<>"Expired")*'+FMp+')'),
    'Expiring Permits',
    'Renewal within 30 days',
    'A');
  _dbKpi(sh, 13, 20, '⚙️',
    _pctF('IFERROR(COUNTIF('+EqS+',"In Use")/COUNTA(Equipment!A2:A)*100,0)'),
    'Equipment Utilization',
    '=IFERROR(TEXT(COUNTIF('+EqS+',"In Use"),"0")&" / "&TEXT(COUNTA(Equipment!A2:A),"0"),"—")',
    'T');

  _dbSp(sh, 18, 10);
  var R = 19;

  /* ── Formula-based chart data (rows 350+) ──────────────── */
  var CD = 350;
  var C = DB.C;

  // Chart 1: Project Status Distribution (donut)
  var statuses = ['Awarded','In Progress','On Hold','Delayed','Completed','Cancelled'];
  var r1 = _dbFData(sh, CD, 'Status', 'Count', statuses,
    statuses.map(function(s){
      return '=COUNTIFS('+PS+',"'+s+'"'+FP+')';
    }));

  // Chart 2: Project Type Mix (donut)
  var CD2 = 362;
  var types = ['Commercial Office Building','High-Rise Residential Tower',
    'Hotel & Hospitality','Industrial Warehouse','Infrastructure - Bridges',
    'Infrastructure - Roads','Marine & Coastal Works','Mixed-Use Development',
    'Retail Mall','School/Educational','Utility Infrastructure','Villa Development'];
  var r2 = _dbFData(sh, CD2, 'Type', 'Count', types,
    types.map(function(t){
      return '=COUNTIFS('+_cr(cols,'Projects','project_type')+',"'+t+'"'+FP+')';
    }));

  // Section + Chart 1 (left)
  sh.getRange(R, 2, 1, 12).merge()
    .setValue('  📊  PROJECT STATUS').setBackground(C.SEC_BG).setFontColor(C.SEC_T)
    .setFontWeight('bold').setFontSize(10).setFontFamily(DB.FONT).setVerticalAlignment('middle');
  sh.getRange(R, 2, 1, 1).setBorder(null,true,null,null,null,null, C.IND, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  // Section + Chart 2 (right)
  sh.getRange(R, 14, 1, 12).merge()
    .setValue('  📊  PROJECT TYPE MIX').setBackground(C.SEC_BG).setFontColor(C.SEC_T)
    .setFontWeight('bold').setFontSize(10).setFontFamily(DB.FONT).setVerticalAlignment('middle');
  sh.getRange(R, 14, 1, 1).setBorder(null,true,null,null,null,null, C.IND, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(R, 28);
  R++;

  _dbChartF(sh, Charts.ChartType.PIE, 'Project Status',
    r1, R, 2, 440, 250,
    ['#9333ea','#16a34a','#d97706','#dc2626','#0891b2','#94a3b8']);
  _dbChartF(sh, Charts.ChartType.PIE, 'Project Type Mix',
    r2, R, 14, 440, 250,
    ['#2563eb','#9333ea','#16a34a','#d97706','#0891b2','#dc2626',
     '#4f46e5','#ec4899','#84cc16','#f97316','#06b6d4','#64748b']);
  R += 14; // reserve rows for charts

  _dbSp(sh, R++, 10);

  /* ── Project Health Matrix ──────────────────────────────── */
  var cerMap = {};
  fd.pay.forEach(function(r){
    if (r.status==='Certified'||r.status==='Paid')
      cerMap[r.project_id] = (cerMap[r.project_id]||0) + (parseFloat(r.gross_value_aed)||0);
  });
  var voMap  = _grpCnt(fd.vo,  'project_id');
  var incMap = _grpCnt(fd.inc, 'project_id');
  var top10  = fd.prj.slice().sort(function(a,b){
    return (parseFloat(b.contract_value_aed)||0) - (parseFloat(a.contract_value_aed)||0);
  }).slice(0, 10);

  R = _dbSec(sh, R, 2, 24, '  🏥  PROJECT HEALTH MATRIX  —  Top 10', top10.length);
  R = _dbTH(sh, R, 2, ['Project','Status','Completion','Contract Value','Certified','VOs','Incidents']);
  top10.forEach(function(p, i){
    var pid = p.project_id;
    R = _dbTR(sh, R, 2, [
      p.project_name || pid,
      p.status || '—',
      _pct(_pctComp(p)),
      _aed(p.contract_value_aed),
      _aed(cerMap[pid]||0),
      voMap[pid]||0,
      incMap[pid]||0
    ], i%2===1);
    _dbBadge(sh, R-1, 3, p.status);
  });

  try { sh.hideRows(350, 60); } catch(e) {}
  SpreadsheetApp.flush();
}
