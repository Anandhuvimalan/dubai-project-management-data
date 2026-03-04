/**
 * db_hse.gs — HSE & Safety Compliance
 * ALL KPIs formula-based with _fm() project filter
 * Charts: formula-based with filter criteria
 */
function _pgHSE(ss, fd, all, filt, gids, cols) {
  var sh = _dbSetupPage(ss, 'hse');
  _dbPaintBg(sh);
  _dbHdr(sh, '🛡️', 'HSE & Safety Compliance');
  _dbNav(sh, gids, 'hse');
  _dbFiltBar(sh, filt, all);

  var IcS  = _cr(cols,'Safety_Incidents','status');
  var IcSv = _cr(cols,'Safety_Incidents','severity');
  var IcLt = _cr(cols,'Safety_Incidents','lti_days');
  var InR  = _cr(cols,'Inspections','result');
  var PmS  = _cr(cols,'Permits_Approvals','status');
  var PmEx = _cr(cols,'Permits_Approvals','expiry_date');
  var FMi  = _fm(cols,'Safety_Incidents');
  var FMn  = _fm(cols,'Inspections');
  var FMp  = _fm(cols,'Permits_Approvals');

  /* ── KPI Row 1 — formula-based with project filter ─────── */
  _dbKpiRows(sh, 8);
  _dbKpi(sh, 8,  2, '🚨',
    _numF('SUMPRODUCT(('+IcS+'="Open")*'+FMi+')+SUMPRODUCT(('+IcS+'="Under Investigation")*'+FMi+')'),
    'Open Incidents',
    'Under investigation / open',
    'R');
  _dbKpi(sh, 8,  8, '💀',
    _numF('SUMPRODUCT(('+IcSv+'="High")*'+FMi+')'),
    'High Severity Cases',
    'Critical HSE incidents',
    'R');
  _dbKpi(sh, 8, 14, '🏥',
    _numF('SUMPRODUCT('+IcLt+'*'+FMi+')'),
    'Total LTI Days',
    'Lost-time injury days',
    'A');
  _dbKpi(sh, 8, 20, '✅',
    _pctF('IFERROR(SUMPRODUCT(('+InR+'="Pass")*'+FMn+')/SUMPRODUCT('+FMn+'*1)*100,0)'),
    'Inspection Pass Rate',
    '=IFERROR(TEXT(SUMPRODUCT('+FMn+'*1),"#,##0")&" total","—")',
    'G');

  /* ── KPI Row 2 ─────────────────────────────────────────── */
  _dbKpiRows(sh, 13);
  _dbKpi(sh, 13,  2, '🔍',
    _numF('SUMPRODUCT('+FMn+'*1)'),
    'Total Inspections',
    'All logged (filtered)',
    'B');
  _dbKpi(sh, 13,  8, '❌',
    _numF('SUMPRODUCT(('+InR+'="Fail")*'+FMn+')'),
    'Failed Inspections',
    'Corrective action needed',
    'R');
  _dbKpi(sh, 13, 14, '📋',
    _numF('SUMPRODUCT(('+PmS+'="Active")*'+FMp+')'),
    'Active Permits',
    'Valid and in force',
    'G');
  _dbKpi(sh, 13, 20, '⚠️',
    _numF('SUMPRODUCT(('+PmEx+'>TODAY())*('+PmEx+'<(TODAY()+30))*('+PmS+'<>"Expired")*'+FMp+')'),
    'Expiring Permits',
    'Within 30 days',
    'A');

  _dbSp(sh, 18, 10);
  var R = 19;
  var C = DB.C;

  /* ── Charts (side by side) ──────────────────────────────── */
  var CD  = 400;
  var CD2 = 410;
  var severities = ['Low','Medium','High','Critical'];
  var r1 = _dbFData(sh, CD, 'Severity', 'Count', severities,
    severities.map(function(s){ return '=SUMPRODUCT(('+IcSv+'="'+s+'")*'+FMi+')'; }));
  var results = ['Pass','Fail','Conditional Pass','Pending'];
  var r2 = _dbFData(sh, CD2, 'Result', 'Count', results,
    results.map(function(r){ return '=SUMPRODUCT(('+InR+'="'+r+'")*'+FMn+')'; }));

  sh.getRange(R, 2, 1, 12).merge().setValue('  📊  INCIDENTS BY SEVERITY')
    .setBackground(C.SEC_BG).setFontColor(C.SEC_T).setFontWeight('bold').setFontSize(10).setFontFamily(DB.FONT).setVerticalAlignment('middle');
  sh.getRange(R, 2, 1, 1).setBorder(null,true,null,null,null,null, C.IND, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(R, 14, 1, 12).merge().setValue('  📊  INSPECTION RESULTS')
    .setBackground(C.SEC_BG).setFontColor(C.SEC_T).setFontWeight('bold').setFontSize(10).setFontFamily(DB.FONT).setVerticalAlignment('middle');
  sh.getRange(R, 14, 1, 1).setBorder(null,true,null,null,null,null, C.IND, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(R, 28); R++;

  _dbChartF(sh, Charts.ChartType.PIE, 'Incidents by Severity', r1, R, 2, 440, 250,
    ['#16a34a','#d97706','#dc2626','#7f1d1d']);
  _dbChartF(sh, Charts.ChartType.COLUMN, 'Inspection Results', r2, R, 14, 440, 250,
    ['#16a34a','#dc2626','#d97706','#2563eb']);
  R += 14;
  _dbSp(sh, R++, 10);

  /* ── Safety Incidents Table ────────────────────────────── */
  var prjMap = {};
  fd.prj.forEach(function(p){ prjMap[p.project_id] = p.project_name || p.project_id; });
  var recInc = fd.inc.slice()
    .sort(function(a,b){ return new Date(b.incident_date) - new Date(a.incident_date); })
    .slice(0, 10);

  R = _dbSec(sh, R, 2, 24, '  🚨  RECENT SAFETY INCIDENTS', fd.inc.length);
  R = _dbTH(sh, R, 2, ['Date','Project','Type','Severity','LTI Days','Status']);
  recInc.forEach(function(inc, i){
    R = _dbTR(sh, R, 2, [
      _days(inc.incident_date), prjMap[inc.project_id]||inc.project_id,
      inc.incident_type||'—', inc.severity||'—', inc.lti_days||0, inc.status||'—'
    ], i%2===1);
    _dbBadge(sh, R-1, 5, inc.severity);
    _dbBadge(sh, R-1, 7, inc.status);
  });

  _dbSp(sh, R++, 8);

  /* ── Permit Expiry Tracker ─────────────────────────────── */
  var today = new Date();
  var pmtSorted = fd.pmt.slice().sort(function(a,b){
    if (!a.expiry_date) return 1; if (!b.expiry_date) return -1;
    return new Date(a.expiry_date) - new Date(b.expiry_date);
  }).slice(0, 10);

  R = _dbSec(sh, R, 2, 24, '  📋  PERMITS EXPIRY TRACKER', fd.pmt.length);
  R = _dbTH(sh, R, 2, ['Permit Type','Project','Reference','Expiry','Days Left','Flag']);
  pmtSorted.forEach(function(pm, i){
    var diff = pm.expiry_date ? Math.round((new Date(pm.expiry_date)-today)/864e5) : null;
    var flag = diff===null?'—':diff<0?'🔴 EXPIRED':diff<30?'🟡 EXPIRING':'🟢 VALID';
    R = _dbTR(sh, R, 2, [
      pm.permit_type||'—', prjMap[pm.project_id]||pm.project_id,
      pm.reference_number||'—', _days(pm.expiry_date),
      diff!==null?diff+' days':'—', flag
    ], i%2===1);
  });

  try { sh.hideRows(400, 30); } catch(e) {}
  SpreadsheetApp.flush();
}
