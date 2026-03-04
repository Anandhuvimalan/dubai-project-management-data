/**
 * db_projects.gs — Projects & Contracts
 * All KPIs use COUNTIFS with _fp() filter criteria referencing _DB_Filters
 * Charts: formula-based with filter criteria
 */
function _pgProjects(ss, fd, all, filt, gids, cols) {
  var sh = _dbSetupPage(ss, 'projects');
  _dbPaintBg(sh);
  _dbHdr(sh, '🏗️', 'Projects & Contracts');
  _dbNav(sh, gids, 'projects');
  _dbFiltBar(sh, filt, all);

  var PS  = _cr(cols,'Projects','status');
  var PCp = _cr(cols,'Projects','completion_pct');
  var PRk = _cr(cols,'Projects','risk_score');
  var VS  = _cr(cols,'Variation_Orders','status');
  var VV  = _cr(cols,'Variation_Orders','vo_value_aed');
  var FP  = _fp(cols);
  var FMvo= _fm(cols,'Variation_Orders');

  /* ── KPI Row 1 ─────────────────────────────────────────── */
  _dbKpiRows(sh, 8);
  _dbKpi(sh, 8,  2, '📋',
    _numF('COUNTIFS('+PS+',"<>"&""'+FP+')'),
    'Total Projects',
    'In filtered portfolio',
    'B');
  _dbKpi(sh, 8,  8, '🔄',
    _numF('COUNTIFS('+PS+',"In Progress"'+FP+')'),
    'In Progress',
    'Actively executing',
    'G');
  _dbKpi(sh, 8, 14, '⏸️',
    _numF('COUNTIFS('+PS+',"On Hold"'+FP+')+COUNTIFS('+PS+',"Delayed"'+FP+')'),
    'On Hold / Delayed',
    'Needs attention',
    'A');
  _dbKpi(sh, 8, 20, '✅',
    _numF('COUNTIFS('+PS+',"Completed"'+FP+')'),
    'Completed',
    'Projects closed out',
    'T');

  /* ── KPI Row 2 ─────────────────────────────────────────── */
  _dbKpiRows(sh, 13);
  _dbKpi(sh, 13,  2, '🔄',
    _numF('SUMPRODUCT(('+VS+'<>"Approved")*('+VS+'<>"Rejected")*('+VS+'<>"")*'+FMvo+')'),
    'Open Variation Orders',
    'Pending approval',
    'A');
  _dbKpi(sh, 13,  8, '💰',
    _aedF('SUMPRODUCT(('+VS+'="Approved")*'+VV+'*'+FMvo+')'),
    'Approved VO Value',
    'Cost impact to date',
    'R');
  _dbKpi(sh, 13, 14, '📊',
    _numF('IFERROR(AVERAGEIFS('+PRk+','+PS+',"<>"&""'+FP+'),0)'),
    'Avg Risk Score',
    '1 = Low  ·  5 = Critical',
    'P');
  _dbKpi(sh, 13, 20, '📈',
    _pctF('IFERROR(AVERAGEIFS('+PCp+','+PS+',"<>"&""'+FP+'),0)'),
    'Avg Completion',
    'Across filtered projects',
    'I');

  _dbSp(sh, 18, 10);
  var R = 19;
  var C = DB.C;

  /* ── Charts (side by side) ──────────────────────────────── */
  var CD = 350;
  var statuses = ['Awarded','In Progress','On Hold','Delayed','Completed','Cancelled'];
  var r1 = _dbFData(sh, CD, 'Status', 'Count', statuses,
    statuses.map(function(s){ return '=COUNTIFS('+PS+',"'+s+'"'+FP+')'; }));
  var CD2 = 360;
  var voStatuses = ['Draft','Submitted','Under Negotiation','Approved','Rejected'];
  var r2 = _dbFData(sh, CD2, 'VO Status', 'Count', voStatuses,
    voStatuses.map(function(s){ return '=SUMPRODUCT(('+VS+'="'+s+'")*'+FMvo+')'; }));

  sh.getRange(R, 2, 1, 12).merge().setValue('  📊  PROJECT STATUS')
    .setBackground(C.SEC_BG).setFontColor(C.SEC_T).setFontWeight('bold').setFontSize(10).setFontFamily(DB.FONT).setVerticalAlignment('middle');
  sh.getRange(R, 2, 1, 1).setBorder(null,true,null,null,null,null, C.IND, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(R, 14, 1, 12).merge().setValue('  📊  VO STATUS')
    .setBackground(C.SEC_BG).setFontColor(C.SEC_T).setFontWeight('bold').setFontSize(10).setFontFamily(DB.FONT).setVerticalAlignment('middle');
  sh.getRange(R, 14, 1, 1).setBorder(null,true,null,null,null,null, C.IND, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(R, 28); R++;

  _dbChartF(sh, Charts.ChartType.COLUMN, 'Project Status', r1, R, 2, 440, 250,
    ['#9333ea','#16a34a','#d97706','#dc2626','#0891b2','#94a3b8']);
  _dbChartF(sh, Charts.ChartType.COLUMN, 'VO Status', r2, R, 14, 440, 250,
    ['#94a3b8','#2563eb','#d97706','#16a34a','#dc2626']);
  R += 14;
  _dbSp(sh, R++, 10);

  /* ── Phase Progress Tracker ────────────────────────────── */
  var prjMap = {};
  fd.prj.forEach(function(p){ prjMap[p.project_id] = p.project_name||p.project_id; });
  var today  = new Date();
  var latePh = fd.ph.filter(function(ph){
    if (!ph.planned_end) return false;
    return (ph.status==='In Progress'||ph.status==='Not Started') && new Date(ph.planned_end) < today;
  }).sort(function(a,b){ return new Date(a.planned_end)-new Date(b.planned_end); }).slice(0,10);

  R = _dbSec(sh, R, 2, 24, '  📅  OVERDUE PHASES', latePh.length);
  R = _dbTH(sh, R, 2, ['Phase','Project','Planned End','Status','Days Behind','Alert']);
  latePh.forEach(function(ph, i){
    var diff  = Math.round((today - new Date(ph.planned_end)) / 864e5);
    R = _dbTR(sh, R, 2, [
      ph.phase_name||'—', prjMap[ph.project_id]||ph.project_id,
      _days(ph.planned_end), ph.status||'—', diff+' days',
      diff > 30 ? '🔴 CRITICAL' : '🟡 BEHIND'
    ], i%2===1);
    _dbBadge(sh, R-1, 5, ph.status);
  });

  try { sh.hideRows(350, 20); } catch(e) {}
  SpreadsheetApp.flush();
}
