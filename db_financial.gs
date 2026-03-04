/**
 * db_financial.gs — Financial Performance
 * All KPIs use COUNTIFS/SUMPRODUCT with dynamic filter refs to _DB_Filters
 * Charts: formula-based with filter criteria
 */
function _pgFinancial(ss, fd, all, filt, gids, cols) {
  var sh = _dbSetupPage(ss, 'financial');
  _dbPaintBg(sh);
  _dbHdr(sh, '💰', 'Financial Performance');
  _dbNav(sh, gids, 'financial');
  _dbFiltBar(sh, filt, all);

  var PA  = _cr(cols,'Payment_Applications','status');
  var PAg = _cr(cols,'Payment_Applications','gross_value_aed');
  var PAr = _cr(cols,'Payment_Applications','retention_aed');
  var CV  = _cr(cols,'Contracts','contract_value_aed');
  var VS  = _cr(cols,'Variation_Orders','status');
  var VV  = _cr(cols,'Variation_Orders','vo_value_aed');
  var VR  = _cr(cols,'Variation_Orders','reason');
  var POs = _cr(cols,'Purchase_Orders','status');
  var POv = _cr(cols,'Purchase_Orders','amount_aed');
  var PBv = _cr(cols,'Projects','budget_variance_pct');
  var PCo = _cr(cols,'Projects','total_cost');
  var PCV = _cr(cols,'Projects','contract_value_aed');
  var FP  = _fp(cols);
  var FMpa= _fm(cols,'Payment_Applications');
  var FMc = _fm(cols,'Contracts');
  var FMvo= _fm(cols,'Variation_Orders');
  var FMpo= _fm(cols,'Purchase_Orders');

  /* ── KPI Row 1 ─────────────────────────────────────────── */
  _dbKpiRows(sh, 8);
  _dbKpi(sh, 8,  2, '📄',
    _aedF('SUMPRODUCT('+CV+'*'+FMc+')'),
    'Total Contract Value',
    '=IFERROR(TEXT(SUMPRODUCT('+FMc+'*1),"0")&" contracts","—")',
    'B');
  _dbKpi(sh, 8,  8, '✅',
    _aedF('SUMPRODUCT(('+PA+'="Certified")*'+PAg+'*'+FMpa+')+SUMPRODUCT(('+PA+'="Paid")*'+PAg+'*'+FMpa+')'),
    'Certified Payments',
    '=IFERROR(TEXT(SUMPRODUCT(('+PA+'="Certified")*'+FMpa+')+SUMPRODUCT(('+PA+'="Paid")*'+FMpa+'),"0")&" apps","—")',
    'G');
  _dbKpi(sh, 8, 14, '🔄',
    _aedF('SUMPRODUCT(('+VS+'="Approved")*'+VV+'*'+FMvo+')'),
    'Approved VO Value',
    '=IFERROR(TEXT(SUMPRODUCT(('+VS+'="Approved")*'+FMvo+'),"0")&" VOs","—")',
    'A');
  _dbKpi(sh, 8, 20, '🛒',
    _aedF('SUMPRODUCT('+POv+'*'+FMpo+')'),
    'Total PO Value',
    '=IFERROR(TEXT(SUMPRODUCT('+FMpo+'*1),"0")&" orders","—")',
    'P');

  /* ── KPI Row 2 ─────────────────────────────────────────── */
  _dbKpiRows(sh, 13);
  _dbKpi(sh, 13,  2, '⏳',
    _aedF('SUMPRODUCT(('+PA+'="Submitted")*'+PAg+'*'+FMpa+')'),
    'Pending Payments',
    '=IFERROR(TEXT(SUMPRODUCT(('+PA+'="Submitted")*'+FMpa+'),"0")&" submitted","—")',
    'A');
  _dbKpi(sh, 13,  8, '📊',
    _pctF('IFERROR(SUMPRODUCT('+PCo+'*('+PCo+'<>"")*(COUNTIFS('+_cr(cols,'Projects','status')+',"<>"&""'+FP+')>0))/SUMPRODUCT('+PCV+'*('+PCV+'<>"")*(COUNTIFS('+_cr(cols,'Projects','status')+',"<>"&""'+FP+')>0))*100,0)'),
    'Budget Utilization',
    'Cost vs contract value',
    'T');
  _dbKpi(sh, 13, 14, '📉',
    _pctF('IFERROR(AVERAGEIFS('+PBv+','+PBv+',">0"'+FP+'),0)'),
    'Avg Budget Overrun',
    '=IFERROR(TEXT(COUNTIFS('+PBv+',">0"'+FP+'),"0")&" over budget","—")',
    'R');
  _dbKpi(sh, 13, 20, '🏦',
    _aedF('SUMPRODUCT(('+PA+'="Certified")*'+PAr+'*'+FMpa+')+SUMPRODUCT(('+PA+'="Submitted")*'+PAr+'*'+FMpa+')'),
    'Retention Held',
    'Pending DLP release',
    'I');

  _dbSp(sh, 18, 10);
  var R = 19;
  var C = DB.C;

  /* ── Charts (row 19-33) ──────────────────────────────── */
  var CD = 350;
  // Chart 1: Payment Status Pipeline
  var payStatuses = ['Submitted','Certified','Paid','Rejected','Under Review','Invoiced'];
  var r1 = _dbFData(sh, CD, 'Status', 'AED', payStatuses,
    payStatuses.map(function(s){
      return '=IFERROR(SUMPRODUCT(('+PA+'="'+s+'")*'+PAg+'*'+FMpa+'),0)';
    }));
  // Chart 2: VO Value by Reason
  var CD2 = 360;
  var voReasons = ['Client Request','Design Change','Site Condition',
    'Scope Increase','Scope Reduction','Unforeseen Conditions'];
  var r2 = _dbFData(sh, CD2, 'Reason', 'AED', voReasons,
    voReasons.map(function(r){
      return '=IFERROR(SUMPRODUCT(('+VR+'="'+r+'")*'+VV+'*'+FMvo+'),0)';
    }));

  sh.getRange(R, 2, 1, 12).merge().setValue('  📊  PAYMENT PIPELINE')
    .setBackground(C.SEC_BG).setFontColor(C.SEC_T).setFontWeight('bold').setFontSize(10).setFontFamily(DB.FONT).setVerticalAlignment('middle');
  sh.getRange(R, 2, 1, 1).setBorder(null,true,null,null,null,null, C.IND, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(R, 14, 1, 12).merge().setValue('  📊  VO VALUE BY REASON')
    .setBackground(C.SEC_BG).setFontColor(C.SEC_T).setFontWeight('bold').setFontSize(10).setFontFamily(DB.FONT).setVerticalAlignment('middle');
  sh.getRange(R, 14, 1, 1).setBorder(null,true,null,null,null,null, C.IND, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(R, 28); R++;

  _dbChartF(sh, Charts.ChartType.BAR, 'Payment Pipeline (AED)', r1, R, 2, 440, 250,
    ['#2563eb','#16a34a','#0891b2','#dc2626','#d97706','#9333ea']);
  _dbChartF(sh, Charts.ChartType.BAR, 'VO Value by Reason (AED)', r2, R, 14, 440, 250,
    ['#d97706','#dc2626','#9333ea','#2563eb','#16a34a','#0891b2']);
  R += 14;
  _dbSp(sh, R++, 10);

  /* ── Budget Health Monitor ──────────────────────────────── */
  var topProj = fd.prj.slice().sort(function(a,b){
    return (parseFloat(b.contract_value_aed)||0) - (parseFloat(a.contract_value_aed)||0);
  }).slice(0, 10);

  R = _dbSec(sh, R, 2, 24, '  💰  BUDGET HEALTH MONITOR', topProj.length);
  R = _dbTH(sh, R, 2, ['Project','Contract Value','Total Cost','Variance %','Flag']);
  topProj.forEach(function(p, i){
    var bv   = parseFloat(p.budget_variance_pct)||0;
    var flag = bv > 10 ? '🔴 OVER' : bv > 0 ? '🟡 WARNING' : '🟢 OK';
    R = _dbTR(sh, R, 2, [
      p.project_name || p.project_id,
      _aed(p.contract_value_aed),
      _aed(p.total_cost),
      _pct(bv),
      flag
    ], i%2===1);
  });

  try { sh.hideRows(350, 30); } catch(e) {}
  SpreadsheetApp.flush();
}
