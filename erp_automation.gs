/**
 * ============================================================
 * DUBAI ERP - FORMULA-BASED AUTOMATION ENGINE
 * ============================================================
 * All automation via spreadsheet formulas — no triggers needed.
 * Creates _Lookups sheet, injects formulas, sets dropdowns,
 * and applies professional styling.
 *
 * Run setupERPSystem() once after importing data.
 * ============================================================
 */

// ============================================================
// CONFIGURATION
// ============================================================
var ERP = {
  ADMIN: 'anandhu7833@gmail.com',
  EXTRA_ROWS: 100, // Pre-fill formulas for new data entry
  LOOKUP_SHEET: '_Lookups',

  SHEETS: {
    'Clients':              {pk:'client_id',     prefix:'CL-',   pad:3, statusCol:'status'},
    'Employees':            {pk:'employee_id',   prefix:'EMP-',  pad:4, statusCol:'status'},
    'Contractors':          {pk:'contractor_id', prefix:'CON-',  pad:3, statusCol:'status'},
    'Suppliers':            {pk:'supplier_id',   prefix:'SUP-',  pad:3, statusCol:'status'},
    'Equipment':            {pk:'equipment_id',  prefix:'EQ-',   pad:4, statusCol:'status'},
    'Projects':             {pk:'project_id',    prefix:'PRJ-',  pad:3, statusCol:'status'},
    'Contracts':            {pk:'contract_id',   prefix:'MC-',   pad:3, statusCol:'status'},
    'Project_Phases':       {pk:'phase_id',      prefix:'PH-',   pad:4, statusCol:'status'},
    'Work_Packages':        {pk:'package_id',    prefix:'WP-',   pad:4, statusCol:null},
    'Permits_Approvals':    {pk:'permit_id',     prefix:'PM-',   pad:3, statusCol:'status'},
    'Inspections':          {pk:'inspection_id', prefix:'INS-',  pad:4, statusCol:'result'},
    'Safety_Incidents':     {pk:'incident_id',   prefix:'INC-',  pad:4, statusCol:'status'},
    'Payment_Applications': {pk:'ipc_id',        prefix:'IPC-',  pad:4, statusCol:'status'},
    'Variation_Orders':     {pk:'vo_id',         prefix:'VO-',   pad:3, statusCol:'status'},
    'Purchase_Orders':      {pk:'po_id',         prefix:'LPO-',  pad:4, statusCol:'status'},
    'Daily_Site_Reports':   {pk:'report_id',     prefix:'DSR-',  pad:4, statusCol:null},
    'Project_Documents':    {pk:'document_id',   prefix:'DOC-',  pad:5, statusCol:'status'}
  },

  COLORS: {
    HEADER_BG:'#2d3748', HEADER_FG:'#f7fafc',
    ALT_ROW:'#f7fafc', WHITE:'#ffffff',
    SYSTEM_BG:'#edf2f7', PRIVATE_BG:'#fef5e7',
    STATUS: {
      'Completed':{bg:'#d4edda',fg:'#155724'}, 'Active':{bg:'#d4edda',fg:'#155724'},
      'Paid':{bg:'#d4edda',fg:'#155724'}, 'Approved':{bg:'#d4edda',fg:'#155724'},
      'Pass':{bg:'#d4edda',fg:'#155724'}, 'Available':{bg:'#d4edda',fg:'#155724'},
      'Closed':{bg:'#d4edda',fg:'#155724'},
      'In Progress':{bg:'#cce5ff',fg:'#004085'}, 'Certified':{bg:'#cce5ff',fg:'#004085'},
      'Under Review':{bg:'#cce5ff',fg:'#004085'}, 'Issued':{bg:'#cce5ff',fg:'#004085'},
      'Delivered':{bg:'#cce5ff',fg:'#004085'}, 'In Use':{bg:'#cce5ff',fg:'#004085'},
      'Under Investigation':{bg:'#cce5ff',fg:'#004085'},
      'On Hold':{bg:'#fff3cd',fg:'#856404'}, 'Submitted':{bg:'#fff3cd',fg:'#856404'},
      'Pending':{bg:'#fff3cd',fg:'#856404'}, 'Draft':{bg:'#fff3cd',fg:'#856404'},
      'Open':{bg:'#fff3cd',fg:'#856404'}, 'Under Negotiation':{bg:'#fff3cd',fg:'#856404'},
      'For Review':{bg:'#fff3cd',fg:'#856404'}, 'Under Maintenance':{bg:'#fff3cd',fg:'#856404'},
      'Conditional Pass':{bg:'#fff3cd',fg:'#856404'},
      'Delayed':{bg:'#f8d7da',fg:'#721c24'}, 'Overdue':{bg:'#f8d7da',fg:'#721c24'},
      'Rejected':{bg:'#f8d7da',fg:'#721c24'}, 'Fail':{bg:'#f8d7da',fg:'#721c24'},
      'Expired':{bg:'#f8d7da',fg:'#721c24'}, 'Blacklisted':{bg:'#f8d7da',fg:'#721c24'},
      'High':{bg:'#f8d7da',fg:'#721c24'}, 'Red':{bg:'#f8d7da',fg:'#721c24'},
      'Cancelled':{bg:'#e2e3e5',fg:'#383d41'}, 'Inactive':{bg:'#e2e3e5',fg:'#383d41'},
      'Not in Contact':{bg:'#e2e3e5',fg:'#383d41'}, 'Decommissioned':{bg:'#e2e3e5',fg:'#383d41'},
      'Superseded':{bg:'#e2e3e5',fg:'#383d41'}, 'Terminated':{bg:'#e2e3e5',fg:'#383d41'},
      'Awarded':{bg:'#e8daef',fg:'#4a235a'}, 'Not Started':{bg:'#e8daef',fg:'#4a235a'},
      'Medium':{bg:'#fff3cd',fg:'#856404'}, 'Amber':{bg:'#fff3cd',fg:'#856404'},
      'Low':{bg:'#d4edda',fg:'#155724'}, 'Green':{bg:'#d4edda',fg:'#155724'},
      'Invoiced':{bg:'#d6eaf8',fg:'#1b4f72'}
    }
  },

  DROPDOWNS: {
    'Clients': {
      'client_type':['Private Developer','Government','Semi-Government','Contractor'],
      'status':['Active','Inactive','Not in Contact','Blacklisted'],
      'payment_terms':['Net 30','Net 45','Net 60','Net 90']
    },
    'Employees': {
      'department':['Engineering','Construction','HSE','Finance','HR','Procurement','QA/QC','Document Control','Admin','IT','PMO','Commercial'],
      'status':['Active','On Leave','Resigned','Terminated','Inactive'],
      'visa_status':['Valid','Expiring Soon','Expired','In Process','Cancelled','Employment Visa','Golden Visa','Mission Visa']
    },
    'Contractors': {
      'specialty':['Concrete Works','Demolition','Electrical','Facade & Cladding','Fire Fighting','Fit-Out','HVAC','Landscaping','MEP','MEP Works','Painting','Piling & Foundations','Plumbing','Road Works','Structural Steel','Waterproofing'],
      'status':['Active','Suspended','Blacklisted'],
      'rating':['A','B','C','D']
    },
    'Suppliers': {
      'category':['Cement & Aggregates','Construction Chemicals','Facade Materials','Finishing Materials','Formwork & Scaffolding','MEP Materials','Ready-Mix Concrete','Safety Equipment','Steel Reinforcement','Structural Steel'],
      'status':['Active','Suspended','Inactive']
    },
    'Equipment': {
      'equipment_type':['Backhoe Loader','Batching Plant','Boom Lift','Compactor','Concrete Pump','Dump Truck','Excavator','Forklift','Generator','Mobile Crane','Scaffolding Set','Tower Crane','Welding Machine','Wheel Loader'],
      'status':['Available','In Use','Under Maintenance','Decommissioned']
    },
    'Projects': {
      'project_type':['Commercial Office Building','High-Rise Residential Tower','Hotel & Hospitality','Industrial Warehouse','Infrastructure - Bridges','Infrastructure - Roads','Marine & Coastal Works','Mixed-Use Development','Retail Mall','School/Educational','Utility Infrastructure','Villa Development'],
      'status':['Awarded','In Progress','On Hold','Delayed','Completed','Cancelled'],
      'health_status':['Green','Amber','Red']
    },
    'Contracts': {
      'contract_type':['Main Contract','Subcontract'],
      'status':['Active','Completed','Terminated','Rejected','Closed']
    },
    'Project_Phases': {
      'phase_name':['Mobilization','Enabling Works','Piling','Substructure','Superstructure','MEP Rough-In','Facade','Internal Finishes','External Works','Testing & Commissioning','Excavation & Shoring','Finishes','Handover'],
      'status':['Not Started','In Progress','Completed','Delayed']
    },
    'Permits_Approvals': {
      'permit_type':['Dubai Municipality - Building Permit','Dubai Municipality - NOC','Civil Defense NOC','DEWA Connection','RTA Access Permit','Crane Permit','DDA Approval','Environmental Permit','Etisalat NOC','Trakhees Permit'],
      'status':['Pending','Submitted','Approved','Expired','Rejected','Renewal Required']
    },
    'Inspections': {
      'inspection_type':['Concrete Pour','Rebar Placement','Formwork','Structural','Waterproofing','MEP Rough-In','Fire Safety','Final Handover'],
      'result':['Pass','Fail','Conditional Pass','Pending','Passed','Failed','Conditional']
    },
    'Safety_Incidents': {
      'incident_type':['Fall from Height','Struck by Object','Electrical','Heat Stroke','Equipment Malfunction','Caught In/Between','Slip/Trip','Chemical Exposure','Fire','Vehicle Accident','Environmental Spill','Fire Incident','First Aid','Lost Time Injury','Medical Treatment','Near Miss','Property Damage'],
      'severity':['Low','Medium','High'],
      'status':['Open','Under Investigation','Closed']
    },
    'Payment_Applications': {
      'status':['Submitted','Under Review','Certified','Paid','Rejected','Overdue']
    },
    'Variation_Orders': {
      'reason':['Client Request','Design Change','Site Condition','Authority Requirement','Scope Addition','Material Substitution'],
      'status':['Draft','Submitted','Under Negotiation','Approved','Rejected']
    },
    'Purchase_Orders': {
      'po_type':['Material','Subcontract'],
      'status':['Draft','Issued','Delivered','Invoiced','Paid']
    },
    'Daily_Site_Reports': {
      'weather':['Clear','Sunny','Hot','Humid','Windy','Sandstorm','Rain']
    },
    'Project_Documents': {
      'document_type':['Shop Drawing','Material Submittal','RFI','Method Statement','Inspection Request','NCR','Progress Report','Meeting Minutes','Site Instruction','Test Report','As-Built Drawing','Variation Order'],
      'status':['Draft','For Review','Approved','Superseded'],
      'discipline':['Civil','Structural','Architectural','MEP','HSE']
    }
  },

  PROTECTED_COLS: ['created_at','created_by','record_version'],
  PRIVATE_COLS: ['salary_aed','contract_value_aed','gross_value_aed','net_certified_aed','amount_aed','credit_limit_aed']
};

// ============================================================
// MAIN SETUP — Run once after importing data
// ============================================================
function setupERPSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ss.toast('Setting up Formula-Based ERP...', '⚙️', -1);

  // 0. Ensure all sheets have enough rows for formulas
  ss.toast('0/7 Expanding sheets...', '⚙️', -1);
  expandAllSheets(ss);

  // 1. Build _Lookups sheet
  ss.toast('1/7 Building lookup tables...', '⚙️', -1);
  buildLookupsSheet(ss);

  // 2. Apply base styling
  ss.toast('2/7 Styling all sheets...', '⚙️', -1);
  styleAllSheets(ss);

  // 3. Apply dropdown validations + FK dropdowns
  ss.toast('3/7 Adding dropdowns...', '⚙️', -1);
  applyAllDropdowns(ss);
  applyFKDropdowns(ss);

  // 4. Inject formulas (auto-ID, timestamps, defaults, calculations)
  ss.toast('4/7 Injecting formulas...', '⚙️', -1);
  injectAllFormulas(ss);

  // 5. Apply conditional formatting rules
  ss.toast('5/7 Conditional formatting...', '⚙️', -1);
  applyConditionalFormatting(ss);

  // 6. Color existing status cells
  ss.toast('6/7 Coloring existing data...', '⚙️', -1);
  colorAllStatusCells(ss);

  // 7. Done!
  ss.toast('7/7 Finalizing...', '⚙️', -1);
  SpreadsheetApp.flush();

  ss.toast('✅ ERP Setup Complete!', '⚙️ Done', 10);
  ui.alert('✅ Formula-Based ERP Ready!\n\n• _Lookups reference sheet\n• FK dropdowns (ID only)\n• Auto-ID + timestamps + default status for new rows\n• Calculated fields (duration, profit, health, etc.)\n• Conditional formatting\n• Pre-filled for ' + ERP.EXTRA_ROWS + ' new rows\n\nTip: Use menu > Resize Columns separately');
}

// ============================================================
// MODULE 1: _Lookups HELPER SHEET
// ============================================================
function buildLookupsSheet(ss) {
  var ls = ss.getSheetByName(ERP.LOOKUP_SHEET);
  if (ls) ls.clear(); else ls = ss.insertSheet(ERP.LOOKUP_SHEET);

  // Column layout: A=emp_display, B=emp_id | D=client_display, E=client_id | G=contractor_display, H=con_id | J=supplier_display, K=sup_id | M=project_display, N=proj_id
  var headers = [];
  headers[0] = 'Employee (Name - ID)'; headers[1] = 'emp_id';
  headers[3] = 'Client (Name - ID)'; headers[4] = 'client_id';
  headers[6] = 'Contractor (Name - ID)'; headers[7] = 'con_id';
  headers[9] = 'Supplier (Name - ID)'; headers[10] = 'sup_id';
  headers[12] = 'Project (Name - ID)'; headers[13] = 'proj_id';
  // Fill gaps with empty
  for (var i = 0; i < 14; i++) { if (!headers[i]) headers[i] = ''; }
  ls.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#2d3748').setFontColor('#f7fafc');

  // Employee list
  var empSheet = ss.getSheetByName('Employees');
  if (empSheet && empSheet.getLastRow() > 1) {
    var empData = empSheet.getDataRange().getValues();
    var empH = empData[0];
    var idIdx = empH.indexOf('employee_id'), nameIdx = empH.indexOf('full_name');
    var empRows = [];
    for (var i = 1; i < empData.length; i++) {
      empRows.push([empData[i][nameIdx] + ' (' + empData[i][idIdx] + ')', empData[i][idIdx]]);
    }
    if (empRows.length > 0) ls.getRange(2, 1, empRows.length, 2).setValues(empRows);
  }

  // Client list
  var cliSheet = ss.getSheetByName('Clients');
  if (cliSheet && cliSheet.getLastRow() > 1) {
    var cliData = cliSheet.getDataRange().getValues();
    var cliH = cliData[0];
    var ciIdx = cliH.indexOf('client_id'), cnIdx = cliH.indexOf('client_name');
    var cliRows = [];
    for (var i = 1; i < cliData.length; i++) {
      cliRows.push([cliData[i][cnIdx] + ' (' + cliData[i][ciIdx] + ')', cliData[i][ciIdx]]);
    }
    if (cliRows.length > 0) ls.getRange(2, 4, cliRows.length, 2).setValues(cliRows);
  }

  // Contractor list
  var conSheet = ss.getSheetByName('Contractors');
  if (conSheet && conSheet.getLastRow() > 1) {
    var conData = conSheet.getDataRange().getValues();
    var conH = conData[0];
    var coIdx = conH.indexOf('contractor_id'), coNIdx = conH.indexOf('company_name');
    var conRows = [];
    for (var i = 1; i < conData.length; i++) {
      conRows.push([conData[i][coNIdx] + ' (' + conData[i][coIdx] + ')', conData[i][coIdx]]);
    }
    if (conRows.length > 0) ls.getRange(2, 7, conRows.length, 2).setValues(conRows);
  }

  // Supplier list
  var supSheet = ss.getSheetByName('Suppliers');
  if (supSheet && supSheet.getLastRow() > 1) {
    var supData = supSheet.getDataRange().getValues();
    var supH = supData[0];
    var siIdx = supH.indexOf('supplier_id'), snIdx = supH.indexOf('company_name');
    var supRows = [];
    for (var i = 1; i < supData.length; i++) {
      supRows.push([supData[i][snIdx] + ' (' + supData[i][siIdx] + ')', supData[i][siIdx]]);
    }
    if (supRows.length > 0) ls.getRange(2, 10, supRows.length, 2).setValues(supRows);
  }

  // Project list
  var prjSheet = ss.getSheetByName('Projects');
  if (prjSheet && prjSheet.getLastRow() > 1) {
    var prjData = prjSheet.getDataRange().getValues();
    var prjH = prjData[0];
    var piIdx = prjH.indexOf('project_id'), pnIdx = prjH.indexOf('project_name');
    var prjRows = [];
    for (var i = 1; i < prjData.length; i++) {
      prjRows.push([prjData[i][pnIdx] + ' (' + prjData[i][piIdx] + ')', prjData[i][piIdx]]);
    }
    if (prjRows.length > 0) ls.getRange(2, 13, prjRows.length, 2).setValues(prjRows);
  }

  ls.autoResizeColumns(1, 14);
  ls.setFrozenRows(1);
  Logger.log('_Lookups sheet built');
}

// ============================================================
// MODULE 2: BASE STYLING
// ============================================================
function styleAllSheets(ss) {
  for (var name in ERP.SHEETS) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    styleSheet(sheet, name);
    SpreadsheetApp.flush();
  }
}

function styleSheet(sheet, sheetName) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastCol <= 0) return;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Font
  sheet.getRange(1, 1, lastRow, lastCol).setFontFamily('Inter').setFontSize(10).setVerticalAlignment('middle');

  // Header
  sheet.getRange(1, 1, 1, lastCol)
    .setBackground(ERP.COLORS.HEADER_BG).setFontColor(ERP.COLORS.HEADER_FG)
    .setFontWeight('bold').setFontSize(11).setWrap(true);
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 36);

  // Alternating rows
  for (var r = 2; r <= lastRow; r++) {
    sheet.getRange(r, 1, 1, lastCol).setBackground(r % 2 === 0 ? ERP.COLORS.ALT_ROW : ERP.COLORS.WHITE);
  }

  // System columns shading
  var cfg = ERP.SHEETS[sheetName];
  for (var c = 0; c < headers.length; c++) {
    var h = headers[c];
    if (h === cfg.pk || ERP.PROTECTED_COLS.indexOf(h) >= 0 || h === 'is_active' || h === 'record_version') {
      sheet.getRange(2, c + 1, lastRow - 1, 1).setBackground(ERP.COLORS.SYSTEM_BG).setFontColor('#718096');
    }
    if (ERP.PRIVATE_COLS.indexOf(h) >= 0) {
      sheet.getRange(2, c + 1, lastRow - 1, 1).setBackground(ERP.COLORS.PRIVATE_BG);
    }
  }
}

// ============================================================
// MODULE 3: DROPDOWN VALIDATION
// ============================================================
function applyAllDropdowns(ss) {
  for (var sheetName in ERP.DROPDOWNS) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var maxRow = sheet.getMaxRows();
    var dds = ERP.DROPDOWNS[sheetName];

    for (var col in dds) {
      var idx = headers.indexOf(col);
      if (idx < 0) continue;
      sheet.getRange(2, idx + 1, maxRow - 1, 1).clearDataValidations();
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(dds[col], true)
        .setAllowInvalid(true)
        .setHelpText('Select: ' + dds[col].join(', '))
        .build();
      sheet.getRange(2, idx + 1, maxRow - 1, 1).setDataValidation(rule);
    }
    SpreadsheetApp.flush();
  }
}

// ============================================================
// MODULE 4: FK DROPDOWNS (ID-only, with name in tooltip)
// ============================================================
function applyFKDropdowns(ss) {
  var ls = ss.getSheetByName(ERP.LOOKUP_SHEET);
  if (!ls) return;

  // Map: { sheetName: { fkColumnHeader: {idCol in _Lookups, displayCol for helpText} } }
  var FK_MAP = {
    'Projects': { 'client_id': {idCol:5}, 'project_manager_id': {idCol:2} },
    'Contracts': { 'project_id': {idCol:14}, 'contractor_id': {idCol:8} },
    'Project_Phases': { 'project_id': {idCol:14} },
    'Work_Packages': { 'project_id': {idCol:14}, 'contractor_id': {idCol:8} },
    'Permits_Approvals': { 'project_id': {idCol:14} },
    'Inspections': { 'project_id': {idCol:14}, 'inspector_id': {idCol:2} },
    'Safety_Incidents': { 'project_id': {idCol:14} },
    'Payment_Applications': { 'project_id': {idCol:14} },
    'Variation_Orders': { 'project_id': {idCol:14} },
    'Purchase_Orders': { 'project_id': {idCol:14} },
    'Daily_Site_Reports': { 'project_id': {idCol:14}, 'prepared_by': {idCol:2} },
    'Project_Documents': { 'project_id': {idCol:14} }
  };

  for (var sheetName in FK_MAP) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var maxRow = sheet.getMaxRows();
    var fks = FK_MAP[sheetName];

    for (var fkCol in fks) {
      var idx = headers.indexOf(fkCol);
      if (idx < 0) continue;
      var info = fks[fkCol];

      // Get ONLY IDs from _Lookups
      var lookupLastRow = ls.getLastRow();
      if (lookupLastRow <= 1) continue;
      var idValues = ls.getRange(2, info.idCol, lookupLastRow - 1, 1).getValues();
      var idList = [];
      for (var v = 0; v < idValues.length; v++) {
        var val = String(idValues[v][0]).trim();
        if (val !== '' && val !== 'undefined') idList.push(val);
      }

      if (idList.length > 0) {
        // Clear existing validation first to avoid red marks
        sheet.getRange(2, idx + 1, maxRow - 1, 1).clearDataValidations();
        var rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(idList, true)
          .setAllowInvalid(true)
          .setHelpText('Select ' + fkCol + ' — see _Lookups sheet for name reference')
          .build();
        sheet.getRange(2, idx + 1, maxRow - 1, 1).setDataValidation(rule);
      }
    }
    SpreadsheetApp.flush();
  }
}

// Ensure all managed sheets have enough rows for formulas
function expandAllSheets(ss) {
  for (var name in ERP.SHEETS) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) continue;
    var needed = sheet.getLastRow() + ERP.EXTRA_ROWS;
    if (sheet.getMaxRows() < needed) {
      sheet.insertRowsAfter(sheet.getMaxRows(), needed - sheet.getMaxRows());
    }
  }
}

// ============================================================
// MODULE 5: FORMULA INJECTION (BATCH-OPTIMIZED)
// ============================================================
function colL(n) { var l=''; while(n>0){var m=(n-1)%26;l=String.fromCharCode(65+m)+l;n=Math.floor((n-1)/26);}return l; }
function hMap(sheet) { var hd=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0],h={}; hd.forEach(function(n,i){h[n]=i+1;}); return h; }
function batchF(sheet,startRow,col,formulas) { if(!formulas.length)return; sheet.getRange(startRow,col,formulas.length,1).setFormulas(formulas.map(function(f){return[f];})); }
var TRIG = {'Clients':'client_name','Employees':'full_name','Contractors':'company_name','Suppliers':'company_name','Equipment':'equipment_type','Projects':'project_name','Contracts':'contract_type','Project_Phases':'phase_name','Work_Packages':'item_description','Permits_Approvals':'permit_type','Inspections':'inspection_type','Safety_Incidents':'incident_type','Payment_Applications':'ipc_number','Variation_Orders':'vo_number','Purchase_Orders':'po_type','Daily_Site_Reports':'report_date','Project_Documents':'document_type'};

function injectAllFormulas(ss) {
  injectProjectFormulas(ss);
  injectClientFormulas(ss);
  injectContractFormulas(ss);
  injectWorkPackageFormulas(ss);
  injectPermitFormulas(ss);
  injectPaymentFormulas(ss);
  injectAutoIDFormulas(ss);
  injectTimestampFormulas(ss);
  injectDefaultValueFormulas(ss);
}

function injectProjectFormulas(ss) {
  var sheet=ss.getSheetByName('Projects'); if(!sheet||sheet.getLastRow()<=1)return;
  var h=hMap(sheet), end=sheet.getLastRow()+ERP.EXTRA_ROWS;
  var sC=colL(h['start_date']),eC=colL(h['end_date']),vC=colL(h['contract_value_aed']),cC=colL(h['completion_pct']),coC=colL(h['total_cost']),rC=colL(h['total_revenue']),rkC=colL(h['risk_score']);
  var dur=[],ela=[],rem=[],bv=[],prf=[],hlth=[];
  for(var r=2;r<=end;r++){
    dur.push('=IF(AND('+sC+r+'<>"",'+eC+r+'<>""),'+eC+r+'-'+sC+r+',"")');
    ela.push('=IF(AND('+sC+r+'<>"",'+eC+r+'<>""),MAX(0,MIN(TODAY(),'+eC+r+')-'+sC+r+'),"")');
    rem.push('=IF('+eC+r+'<>"",MAX(0,'+eC+r+'-TODAY()),"")');
    bv.push('=IF(AND('+vC+r+'>0,'+cC+r+'>0),ROUND((('+coC+r+'/('+vC+r+'*'+cC+r+'/100))-1)*100,2),0)');
    prf.push('=IF('+rC+r+'<>"",'+rC+r+'-'+coC+r+',"")');
    hlth.push('=IF('+rkC+r+'="","",IF('+rkC+r+'<=3,"Green",IF('+rkC+r+'<=6,"Amber","Red")))');
  }
  if(h['duration_days'])batchF(sheet,2,h['duration_days'],dur);
  if(h['elapsed_days'])batchF(sheet,2,h['elapsed_days'],ela);
  if(h['remaining_days'])batchF(sheet,2,h['remaining_days'],rem);
  if(h['budget_variance_pct'])batchF(sheet,2,h['budget_variance_pct'],bv);
  if(h['profit'])batchF(sheet,2,h['profit'],prf);
  if(h['health_status'])batchF(sheet,2,h['health_status'],hlth);
  SpreadsheetApp.flush(); Logger.log('Project formulas injected (batch)');
}

function injectClientFormulas(ss) {
  var sheet=ss.getSheetByName('Clients'),prj=ss.getSheetByName('Projects');
  if(!sheet||!prj||sheet.getLastRow()<=1)return;
  var h=hMap(sheet),pH=prj.getRange(1,1,1,prj.getLastColumn()).getValues()[0];
  var cC=colL(h['client_id']),pCC=colL(pH.indexOf('client_id')+1),pRC=colL(pH.indexOf('total_revenue')+1);
  var pSC=colL(pH.indexOf('status')+1),pHC=colL(pH.indexOf('health_status')+1);
  var pE=prj.getLastRow()+ERP.EXTRA_ROWS,end=sheet.getLastRow()+ERP.EXTRA_ROWS;
  var rev=[],act=[],rsk=[];
  for(var r=2;r<=end;r++){
    var rng='$2:$'+pE;
    rev.push('=IF('+cC+r+'<>"",SUMIF(Projects!'+pCC+rng+','+cC+r+',Projects!'+pRC+rng+'),"")');
    act.push('=IF('+cC+r+'<>"",COUNTIFS(Projects!'+pCC+rng+','+cC+r+',Projects!'+pSC+rng+',"<>Completed",Projects!'+pSC+rng+',"<>Cancelled"),"")');
    rsk.push('=IF('+cC+r+'<>"",COUNTIFS(Projects!'+pCC+rng+','+cC+r+',Projects!'+pHC+rng+',"Red"),"")');
  }
  if(h['lifetime_revenue'])batchF(sheet,2,h['lifetime_revenue'],rev);
  if(h['active_projects'])batchF(sheet,2,h['active_projects'],act);
  if(h['risk_projects'])batchF(sheet,2,h['risk_projects'],rsk);
  SpreadsheetApp.flush();
}

function injectContractFormulas(ss) {
  var sheet=ss.getSheetByName('Contracts'); if(!sheet||sheet.getLastRow()<=1)return;
  var h=hMap(sheet); if(!h['remaining_value']||!h['payment_progress_pct'])return;
  var vC=colL(h['contract_value_aed']),pC=colL(h['payment_progress_pct']),end=sheet.getLastRow()+ERP.EXTRA_ROWS;
  var f=[]; for(var r=2;r<=end;r++) f.push('=IF('+vC+r+'<>"",'+vC+r+'*(1-'+pC+r+'/100),"")');
  batchF(sheet,2,h['remaining_value'],f); SpreadsheetApp.flush();
}

function injectWorkPackageFormulas(ss) {
  var sheet=ss.getSheetByName('Work_Packages'); if(!sheet||sheet.getLastRow()<=1)return;
  var h=hMap(sheet),qC=colL(h['quantity']),rC=colL(h['unit_rate_aed']),eC=colL(h['executed_qty']),end=sheet.getLastRow()+ERP.EXTRA_ROWS;
  var tot=[],exe=[]; for(var r=2;r<=end;r++){
    tot.push('=IF(AND('+qC+r+'<>"",'+rC+r+'<>""),'+qC+r+'*'+rC+r+',"")');
    exe.push('=IF(AND('+eC+r+'<>"",'+rC+r+'<>""),'+eC+r+'*'+rC+r+',"")');
  }
  if(h['total_amount_aed'])batchF(sheet,2,h['total_amount_aed'],tot);
  if(h['executed_amount_aed'])batchF(sheet,2,h['executed_amount_aed'],exe);
  SpreadsheetApp.flush();
}

function injectPermitFormulas(ss) {
  var sheet=ss.getSheetByName('Permits_Approvals'); if(!sheet||sheet.getLastRow()<=1)return;
  var h=hMap(sheet); if(!h['is_expired']||!h['expiry_date'])return;
  var xC=colL(h['expiry_date']),end=sheet.getLastRow()+ERP.EXTRA_ROWS;
  var f=[]; for(var r=2;r<=end;r++) f.push('=IF('+xC+r+'<>"",IF(TODAY()>'+xC+r+',TRUE,FALSE),"")');
  batchF(sheet,2,h['is_expired'],f); SpreadsheetApp.flush();
}

function injectPaymentFormulas(ss) {
  var sheet=ss.getSheetByName('Payment_Applications'); if(!sheet||sheet.getLastRow()<=1)return;
  var h=hMap(sheet); if(!h['net_certified_aed'])return;
  var gC=colL(h['gross_value_aed']),rC=colL(h['retention_aed']),dC=colL(h['deductions_aed']),end=sheet.getLastRow()+ERP.EXTRA_ROWS;
  var f=[]; for(var r=2;r<=end;r++) f.push('=IF('+gC+r+'<>"",'+gC+r+'-IF('+rC+r+'<>"",'+rC+r+',0)-IF('+dC+r+'<>"",'+dC+r+',0),"")');
  batchF(sheet,2,h['net_certified_aed'],f); SpreadsheetApp.flush();
}

function injectAutoIDFormulas(ss) {
  for(var sn in ERP.SHEETS){
    var sheet=ss.getSheetByName(sn); if(!sheet||sheet.getLastRow()<=1)continue;
    var cfg=ERP.SHEETS[sn],h=hMap(sheet),pkI=h[cfg.pk]; if(!pkI)continue;
    var tC=TRIG[sn]||Object.keys(h)[1],tI=h[tC]||2;
    var pkL=colL(pkI),tL=colL(tI),pad=new Array(cfg.pad+1).join('0');
    var dEnd=sheet.getLastRow(),end=dEnd+ERP.EXTRA_ROWS,f=[];
    for(var r=dEnd+1;r<=end;r++) f.push('=IF('+tL+r+'<>"","'+cfg.prefix+'"&TEXT(COUNTA('+pkL+'$2:'+pkL+(r-1)+')+1,"'+pad+'"),"")');
    if(f.length>0)batchF(sheet,dEnd+1,pkI,f);
    SpreadsheetApp.flush();
  }
}

function injectTimestampFormulas(ss) {
  for(var sn in ERP.SHEETS){
    var sheet=ss.getSheetByName(sn); if(!sheet||sheet.getLastRow()<=1)continue;
    var h=hMap(sheet),tC=TRIG[sn]||Object.keys(h)[1],tI=h[tC]||2,tL=colL(tI);
    var dEnd=sheet.getLastRow(),end=dEnd+ERP.EXTRA_ROWS,cnt=end-dEnd; if(cnt<=0)continue;
    var ca=[],ua=[],cb=[],rv=[],ia=[];
    for(var r=dEnd+1;r<=end;r++){
      ca.push('=IF('+tL+r+'<>"",NOW(),"")');
      ua.push('=IF('+tL+r+'<>"",NOW(),"")');
      cb.push('=IF('+tL+r+'<>"","'+ERP.ADMIN+'","")');
      rv.push('=IF('+tL+r+'<>"",1,"")');
      ia.push('=IF('+tL+r+'<>"",TRUE,"")');
    }
    if(h['created_at'])batchF(sheet,dEnd+1,h['created_at'],ca);
    if(h['last_updated_at'])batchF(sheet,dEnd+1,h['last_updated_at'],ua);
    if(h['created_by'])batchF(sheet,dEnd+1,h['created_by'],cb);
    if(h['record_version'])batchF(sheet,dEnd+1,h['record_version'],rv);
    if(h['is_active'])batchF(sheet,dEnd+1,h['is_active'],ia);
    SpreadsheetApp.flush();
  }
}

// --- Default values for new rows (status, tags, internal_notes, etc.) ---
function injectDefaultValueFormulas(ss) {
  // Default status per sheet when trigger column is filled
  var DEFAULTS = {
    'Clients':            {status:'Active', tags:'"—"', internal_notes:'"—"'},
    'Employees':          {status:'Active', visa_status:'Valid', tags:'"—"', internal_notes:'"—"'},
    'Contractors':        {status:'Active', prequalified:'FALSE', rating:'"C"', tags:'"—"', internal_notes:'"—"'},
    'Suppliers':          {status:'Active', approved:'TRUE', tags:'"—"', internal_notes:'"—"'},
    'Equipment':          {status:'Available', tags:'"—"', internal_notes:'"—"'},
    'Projects':           {status:'Awarded', completion_pct:'0', risk_score:'1', tags:'"—"', internal_notes:'"—"'},
    'Contracts':          {status:'Active', retention_pct:'10', performance_bond_pct:'10', advance_payment_pct:'0', defect_liability_months:'12', payment_progress_pct:'0', tags:'"—"', internal_notes:'"—"'},
    'Project_Phases':     {status:'Not Started', progress_pct:'0', tags:'"—"', internal_notes:'"—"'},
    'Work_Packages':      {executed_qty:'0', tags:'"—"', internal_notes:'"—"'},
    'Permits_Approvals':  {status:'Pending', tags:'"—"', internal_notes:'"—"', remarks:'"—"'},
    'Inspections':        {result:'Pending', defects_found:'0', tags:'"—"', internal_notes:'"—"', remarks:'"—"', corrective_action:'"N/A"'},
    'Safety_Incidents':   {severity:'Low', status:'Open', injured_persons:'0', lti_days:'0', tags:'"—"', internal_notes:'"—"'},
    'Payment_Applications': {status:'Submitted', deductions_aed:'0', tags:'"—"', internal_notes:'"—"'},
    'Variation_Orders':   {status:'Draft', tags:'"—"', internal_notes:'"—"'},
    'Purchase_Orders':    {status:'Draft', tags:'"—"', internal_notes:'"—"'},
    'Daily_Site_Reports': {weather:'Clear', tags:'"—"', internal_notes:'"—"', issues:'"None reported"'},
    'Project_Documents':  {status:'Draft', tags:'"—"', internal_notes:'"—"'}
  };

  for (var sn in DEFAULTS) {
    var sheet = ss.getSheetByName(sn); if (!sheet || sheet.getLastRow() <= 1) continue;
    var h = hMap(sheet);
    var tC = TRIG[sn] || Object.keys(h)[1], tI = h[tC] || 2, tL = colL(tI);
    var dEnd = sheet.getLastRow(), end = dEnd + ERP.EXTRA_ROWS;
    if (end <= dEnd) continue;
    var defs = DEFAULTS[sn];

    for (var col in defs) {
      if (!h[col]) continue;
      var val = defs[col];
      var isString = String(val).charAt(0) === '"';
      var formulas = [];
      for (var r = dEnd + 1; r <= end; r++) {
        if (isString) {
          formulas.push('=IF(' + tL + r + '<>"",' + val + ',"")');
        } else {
          formulas.push('=IF(' + tL + r + '<>"",' + val + ',"")');
        }
      }
      batchF(sheet, dEnd + 1, h[col], formulas);
    }
    SpreadsheetApp.flush();
  }
}

// ============================================================
// MODULE 7: AUTO-RESIZE (separate, not in setup to avoid timeout)
// ============================================================
function autoResizeAll(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  for (var name in ERP.SHEETS) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    try {
      var cols = Math.min(sheet.getLastColumn(), 20);
      sheet.autoResizeColumns(1, cols);
      for (var c = 1; c <= cols; c++) {
        if (sheet.getColumnWidth(c) > 250) sheet.setColumnWidth(c, 250);
        if (sheet.getColumnWidth(c) < 80) sheet.setColumnWidth(c, 80);
      }
    } catch(e) {}
    SpreadsheetApp.flush();
  }
}
// ============================================================
function applyConditionalFormatting(ss) {
  for (var sheetName in ERP.SHEETS) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    var cfg = ERP.SHEETS[sheetName];
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var maxRow = sheet.getMaxRows();

    // Find status/result/severity/health columns
    var colorCols = [];
    if (cfg.statusCol) colorCols.push(cfg.statusCol);
    ['health_status','severity','result'].forEach(function(c) {
      if (headers.indexOf(c) >= 0 && colorCols.indexOf(c) < 0) colorCols.push(c);
    });
    if (colorCols.length === 0) continue;

    sheet.clearConditionalFormatRules();
    var rules = [];

    for (var ci = 0; ci < colorCols.length; ci++) {
      var colIdx = headers.indexOf(colorCols[ci]);
      if (colIdx < 0) continue;
      var colLetter = '';
      var cn = colIdx + 1;
      while (cn > 0) { var m = (cn-1) % 26; colLetter = String.fromCharCode(65+m) + colLetter; cn = Math.floor((cn-1)/26); }

      var cellRange = sheet.getRange(2, colIdx + 1, maxRow - 1, 1);

      // Add a rule for each known status value
      for (var status in ERP.COLORS.STATUS) {
        var style = ERP.COLORS.STATUS[status];
        try {
          var rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(status)
            .setBackground(style.bg)
            .setFontColor(style.fg)
            .setBold(true)
            .setRanges([cellRange])
            .build();
          rules.push(rule);
        } catch(e) {}
      }
    }
    if (rules.length > 0) sheet.setConditionalFormatRules(rules);
    SpreadsheetApp.flush();
  }
}

// ============================================================
// MODULE 6B: COLOR EXISTING STATUS CELLS (one-time for imported data)
// ============================================================
function colorAllStatusCells(ss) {
  for (var sheetName in ERP.SHEETS) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var data = sheet.getDataRange().getValues();
    var cfg = ERP.SHEETS[sheetName];

    var colorCols = [];
    if (cfg.statusCol) colorCols.push(cfg.statusCol);
    ['health_status','severity','result'].forEach(function(c) {
      if (headers.indexOf(c) >= 0 && colorCols.indexOf(c) < 0) colorCols.push(c);
    });

    for (var ci = 0; ci < colorCols.length; ci++) {
      var colIdx = headers.indexOf(colorCols[ci]);
      if (colIdx < 0) continue;
      for (var r = 1; r < data.length; r++) {
        var val = String(data[r][colIdx]).trim();
        var style = ERP.COLORS.STATUS[val];
        if (style) {
          sheet.getRange(r + 1, colIdx + 1)
            .setBackground(style.bg).setFontColor(style.fg)
            .setFontWeight('bold').setHorizontalAlignment('center');
        }
      }
    }
    SpreadsheetApp.flush();
  }
}

// ============================================================
// MODULE 7: AUTO-RESIZE
// ============================================================
function autoResizeAll(ss) {
  for (var name in ERP.SHEETS) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    try {
      var cols = Math.min(sheet.getLastColumn(), 25);
      sheet.autoResizeColumns(1, cols);
      for (var c = 1; c <= cols; c++) {
        if (sheet.getColumnWidth(c) > 250) sheet.setColumnWidth(c, 250);
        if (sheet.getColumnWidth(c) < 80) sheet.setColumnWidth(c, 80);
      }
    } catch(e) {}
    SpreadsheetApp.flush();
  }
}

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏗️ Dubai ERP')
    .addItem('⚙️ Setup ERP System', 'setupERPSystem')
    .addSeparator()
    .addItem('📥 Import All Data', 'importAllData')
    .addSeparator()
    .addItem('🔄 Rebuild Lookups', 'rebuildLookups')
    .addItem('🎨 Reapply Formatting', 'reapplyFormatting')
    .addItem('🔄 Recolor Status Cells', 'recolorStatus')
    .addItem('📏 Resize All Columns', 'resizeColumnsMenu')
    .addToUi();
}

function rebuildLookups() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Rebuilding lookup tables...', '🔄', -1);
  buildLookupsSheet(ss);
  ss.toast('✅ Lookups rebuilt!', 'Done', 5);
}

function reapplyFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Reapplying formatting...', '🎨', -1);
  styleAllSheets(ss);
  applyAllDropdowns(ss);
  colorAllStatusCells(ss);
  autoResizeAll(ss);
  ss.toast('✅ Formatting applied!', 'Done', 5);
}

function recolorStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Recoloring status cells...', '🔄', -1);
  colorAllStatusCells(ss);
  ss.toast('✅ Status cells recolored!', 'Done', 5);
}

function resizeColumnsMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Resizing columns...', '📏', -1);
  autoResizeAll(ss);
  ss.toast('✅ Columns resized!', 'Done', 5);
}
