/**
 * db_workforce.gs — Workforce & Contractors
 * KPIs use COUNTIFS with _fp() or raw COUNTIF for non-project sheets
 * Charts: formula-based
 */
function _pgWorkforce(ss, fd, all, filt, gids, cols) {
  var sh = _dbSetupPage(ss, 'workforce');
  _dbPaintBg(sh);
  _dbHdr(sh, '👷', 'Workforce & Contractors');
  _dbNav(sh, gids, 'workforce');
  _dbFiltBar(sh, filt, all);

  var ES  = _cr(cols,'Employees','status');
  var EV  = _cr(cols,'Employees','visa_status');
  var CS  = _cr(cols,'Contractors','status');
  var CR  = _cr(cols,'Contractors','rating');
  var CPq = _cr(cols,'Contractors','prequalified');
  var SS  = _cr(cols,'Suppliers','status');
  var DD  = _cr(cols,'Daily_Site_Reports','manpower_direct');
  var DS  = _cr(cols,'Daily_Site_Reports','manpower_subcon');
  var FMd = _fm(cols,'Daily_Site_Reports');

  /* ── KPI Row 1 ─────────────────────────────────────────── */
  _dbKpiRows(sh, 8);
  _dbKpi(sh, 8,  2, '👤',
    _numF('COUNTIF('+ES+',"Active")'),
    'Active Employees',
    '=IFERROR(TEXT(COUNTA(Employees!A2:A),"#,##0")&" on payroll","—")',
    'B');
  _dbKpi(sh, 8,  8, '🏢',
    _numF('COUNTIF('+CS+',"Active")'),
    'Active Contractors',
    '=IFERROR(TEXT(COUNTA(Contractors!A2:A),"#,##0")&" in database","—")',
    'G');
  _dbKpi(sh, 8, 14, '👷',
    _numF('IFERROR(SUMPRODUCT('+DD+'*'+FMd+')/SUMPRODUCT('+FMd+'*1)+SUMPRODUCT('+DS+'*'+FMd+')/SUMPRODUCT('+FMd+'*1),0)'),
    'Avg Daily Manpower',
    'Direct + Subcon (filtered)',
    'T');
  _dbKpi(sh, 8, 20, '⭐',
    '=IFERROR(TEXT((COUNTIF('+CR+',"A")*4+COUNTIF('+CR+',"B")*3+COUNTIF('+CR+',"C")*2+COUNTIF('+CR+',"D")*1)/COUNTA(Contractors!A2:A),"0.0"),"—")',
    'Avg Contractor Rating',
    'A=4 · B=3 · C=2 · D=1',
    'P');

  /* ── KPI Row 2 ─────────────────────────────────────────── */
  _dbKpiRows(sh, 13);
  _dbKpi(sh, 13,  2, '🚨',
    _numF('COUNTIF('+EV+',"Expired")'),
    'Expired Visas',
    'Employees needing renewal',
    'R');
  _dbKpi(sh, 13,  8, '✅',
    _numF('COUNTIF('+CPq+',TRUE)+COUNTIF('+CPq+',"TRUE")+COUNTIF('+CPq+',"true")'),
    'Prequalified Ctrs',
    'Trade license verified',
    'G');
  _dbKpi(sh, 13, 14, '🚚',
    _numF('COUNTIF('+SS+',"Active")'),
    'Active Suppliers',
    '=IFERROR(TEXT(COUNTA(Suppliers!A2:A),"#,##0")&" registered","—")',
    'B');
  _dbKpi(sh, 13, 20, '🥇',
    _numF('COUNTIF('+CR+',"A")'),
    'Grade A Contractors',
    'Highest performance tier',
    'A');

  _dbSp(sh, 18, 10);
  var R = 19;
  var C = DB.C;

  /* ── Charts (side by side) ──────────────────────────────── */
  var CD  = 400;
  var CD2 = 410;
  var ratings = ['A','B','C','D'];
  var r1 = _dbFData(sh, CD, 'Rating', 'Count', ratings,
    ratings.map(function(g){ return '=COUNTIF('+CR+',"'+g+'")'; }));
  var empSt = ['Active','On Leave','Resigned','Inactive'];
  var r2 = _dbFData(sh, CD2, 'Status', 'Count', empSt,
    empSt.map(function(s){ return '=COUNTIF('+ES+',"'+s+'")'; }));

  sh.getRange(R, 2, 1, 12).merge().setValue('  📊  CONTRACTOR RATINGS')
    .setBackground(C.SEC_BG).setFontColor(C.SEC_T).setFontWeight('bold').setFontSize(10).setFontFamily(DB.FONT).setVerticalAlignment('middle');
  sh.getRange(R, 2, 1, 1).setBorder(null,true,null,null,null,null, C.IND, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(R, 14, 1, 12).merge().setValue('  📊  EMPLOYEE STATUS')
    .setBackground(C.SEC_BG).setFontColor(C.SEC_T).setFontWeight('bold').setFontSize(10).setFontFamily(DB.FONT).setVerticalAlignment('middle');
  sh.getRange(R, 14, 1, 1).setBorder(null,true,null,null,null,null, C.IND, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(R, 28); R++;

  _dbChartF(sh, Charts.ChartType.BAR, 'Contractor Ratings', r1, R, 2, 440, 250,
    ['#16a34a','#2563eb','#d97706','#dc2626']);
  _dbChartF(sh, Charts.ChartType.PIE, 'Employee Status', r2, R, 14, 440, 250,
    ['#16a34a','#d97706','#94a3b8','#e2e3e5']);
  R += 14;
  _dbSp(sh, R++, 10);

  /* ── Contractor Scorecard ──────────────────────────────── */
  var scored = all.ctr.filter(function(r){ return r.status==='Active'; })
    .map(function(c){ return { c:c, n: c.rating==='A'?4:c.rating==='B'?3:c.rating==='C'?2:1 }; })
    .sort(function(a,b){ return b.n - a.n; }).slice(0, 10);

  R = _dbSec(sh, R, 2, 24, '  🏆  CONTRACTOR SCORECARD', scored.length);
  R = _dbTH(sh, R, 2, ['Contractor','Specialty','Rating','Pre-Qual','Status','Grade']);
  scored.forEach(function(item, i){
    var c = item.c;
    var pq = (c.prequalified===true||String(c.prequalified).toLowerCase()==='true') ? '✅' : '❌';
    R = _dbTR(sh, R, 2, [
      c.company_name||c.contractor_id, c.specialty||'—', c.rating||'—',
      pq, c.status||'—',
      c.rating==='A'?'🥇 A':c.rating==='B'?'🥈 B':c.rating==='C'?'🥉 C':'⬜ D'
    ], i%2===1);
    _dbBadge(sh, R-1, 6, c.status);
  });

  try { sh.hideRows(400, 20); } catch(e) {}
  SpreadsheetApp.flush();
}
