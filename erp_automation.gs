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
  EXTRA_ROWS: 200, // Pre-fill formulas for new data entry
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
      'department':['Engineering','Construction','HSE','Finance','HR','Procurement','QA/QC','Document Control','Admin','IT','PMO'],
      'status':['Active','On Leave','Resigned','Terminated'],
      'visa_status':['Valid','Expiring Soon','Expired','In Process','Cancelled']
    },
    'Contractors': {
      'specialty':['Concrete Works','Demolition','Electrical','Facade & Cladding','Fire Fighting','Fit-Out','HVAC','Landscaping','MEP','Painting','Piling & Foundations','Plumbing','Structural Steel','Waterproofing'],
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
      'status':['Active','Completed','Terminated','Rejected']
    },
    'Project_Phases': {
      'phase_name':['Mobilization','Enabling Works','Piling','Substructure','Superstructure','MEP Rough-In','Facade','Internal Finishes','External Works','Testing & Commissioning'],
      'status':['Not Started','In Progress','Completed','Delayed']
    },
    'Permits_Approvals': {
      'permit_type':['Dubai Municipality - Building Permit','Dubai Municipality - NOC','Civil Defense NOC','DEWA Connection','RTA Access Permit','Crane Permit','DDA Approval','Environmental Permit','Etisalat NOC','Trakhees Permit'],
      'status':['Pending','Submitted','Approved','Expired','Rejected']
    },
    'Inspections': {
      'inspection_type':['Concrete Pour','Rebar Placement','Formwork','Structural','Waterproofing','MEP Rough-In','Fire Safety','Final Handover'],
      'result':['Pass','Fail','Conditional Pass','Pending']
    },
    'Safety_Incidents': {
      'incident_type':['Fall from Height','Struck by Object','Electrical','Heat Stroke','Equipment Malfunction','Caught In/Between','Slip/Trip','Chemical Exposure','Fire','Vehicle Accident'],
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

  // 1. Build _Lookups sheet
  ss.toast('1/7 Building lookup tables...', '⚙️', -1);
  buildLookupsSheet(ss);

  // 2. Apply base styling
  ss.toast('2/7 Styling all sheets...', '⚙️', -1);
  styleAllSheets(ss);

  // 3. Apply dropdown validations
  ss.toast('3/7 Adding dropdowns...', '⚙️', -1);
  applyAllDropdowns(ss);

  // 4. Apply smart FK dropdowns (name+ID)
  ss.toast('4/7 Smart FK dropdowns...', '⚙️', -1);
  applySmartFKDropdowns(ss);

  // 5. Inject formulas (auto-ID, timestamps, calculations)
  ss.toast('5/7 Injecting formulas...', '⚙️', -1);
  injectAllFormulas(ss);

  // 6. Apply status cell coloring
  ss.toast('6/7 Status formatting...', '⚙️', -1);
  colorAllStatusCells(ss);

  // 7. Auto-resize columns
  ss.toast('7/7 Resizing columns...', '⚙️', -1);
  autoResizeAll(ss);

  ss.toast('✅ ERP Setup Complete!', '⚙️ Done', 10);
  ui.alert('✅ Formula-Based ERP System Ready!\n\n• _Lookups helper sheet created\n• Smart dropdowns (Name + ID) on all FK fields\n• Auto-ID formulas for new rows\n• Auto-calculated fields (duration, risk, health, profit)\n• Colorblind-safe status badges\n• All formulas pre-filled for ' + ERP.EXTRA_ROWS + ' new rows');
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
    var endRow = sheet.getLastRow() + ERP.EXTRA_ROWS;
    var dds = ERP.DROPDOWNS[sheetName];

    for (var col in dds) {
      var idx = headers.indexOf(col);
      if (idx < 0) continue;
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(dds[col], true)
        .setAllowInvalid(true)
        .setHelpText('Select: ' + dds[col].join(', '))
        .build();
      sheet.getRange(2, idx + 1, endRow - 1, 1).setDataValidation(rule);
    }
    SpreadsheetApp.flush();
  }
}

// ============================================================
// MODULE 4: SMART FK DROPDOWNS (Name + ID)
// ============================================================
function applySmartFKDropdowns(ss) {
  var ls = ss.getSheetByName(ERP.LOOKUP_SHEET);
  if (!ls) return;

  // Map: which lookup column has the display values for each FK
  // Format: { sheetName: { fkColumnHeader: lookupColumnLetter } }
  var FK_MAP = {
    'Projects': {
      'client_id': {displayCol: 4, idCol: 5},        // _Lookups D:E
      'project_manager_id': {displayCol: 1, idCol: 2} // _Lookups A:B
    },
    'Contracts': {
      'project_id': {displayCol: 13, idCol: 14},      // _Lookups M:N
      'contractor_id': {displayCol: 7, idCol: 8}      // _Lookups G:H
    },
    'Project_Phases': {
      'project_id': {displayCol: 13, idCol: 14}
    },
    'Work_Packages': {
      'project_id': {displayCol: 13, idCol: 14},
      'contractor_id': {displayCol: 7, idCol: 8}
    },
    'Permits_Approvals': {
      'project_id': {displayCol: 13, idCol: 14}
    },
    'Inspections': {
      'project_id': {displayCol: 13, idCol: 14},
      'inspector_id': {displayCol: 1, idCol: 2}
    },
    'Safety_Incidents': {
      'project_id': {displayCol: 13, idCol: 14}
    },
    'Payment_Applications': {
      'project_id': {displayCol: 13, idCol: 14}
    },
    'Variation_Orders': {
      'project_id': {displayCol: 13, idCol: 14}
    },
    'Purchase_Orders': {
      'project_id': {displayCol: 13, idCol: 14}
    },
    'Daily_Site_Reports': {
      'project_id': {displayCol: 13, idCol: 14},
      'prepared_by': {displayCol: 1, idCol: 2}
    },
    'Project_Documents': {
      'project_id': {displayCol: 13, idCol: 14}
    }
  };

  for (var sheetName in FK_MAP) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var endRow = sheet.getLastRow() + ERP.EXTRA_ROWS;
    var fks = FK_MAP[sheetName];

    for (var fkCol in fks) {
      var idx = headers.indexOf(fkCol);
      if (idx < 0) continue;
      var info = fks[fkCol];

      // Get lookup values
      var lookupLastRow = ls.getLastRow();
      var displayValues = ls.getRange(2, info.displayCol, Math.max(1, lookupLastRow - 1), 1).getValues();
      var valueList = [];
      for (var v = 0; v < displayValues.length; v++) {
        if (displayValues[v][0] && String(displayValues[v][0]).trim() !== '') {
          valueList.push(String(displayValues[v][0]));
        }
      }

      if (valueList.length > 0) {
        // For existing data: keep current IDs as-is
        // For new rows below data: set dropdown with Name (ID) format
        var newRowStart = sheet.getLastRow() + 1;
        if (newRowStart <= endRow) {
          var rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(valueList, true)
            .setAllowInvalid(true)
            .setHelpText('Search & select ' + fkCol)
            .build();
          sheet.getRange(newRowStart, idx + 1, endRow - newRowStart + 1, 1).setDataValidation(rule);
        }
      }
    }
    SpreadsheetApp.flush();
  }
}

// ============================================================
// MODULE 5: FORMULA INJECTION
// ============================================================
function injectAllFormulas(ss) {
  injectProjectFormulas(ss);
  injectClientFormulas(ss);
  injectContractFormulas(ss);
  injectPhaseFormulas(ss);
  injectWorkPackageFormulas(ss);
  injectPermitFormulas(ss);
  injectPaymentFormulas(ss);
  injectAutoIDFormulas(ss);
  injectTimestampFormulas(ss);
}

// --- Projects: duration, elapsed, remaining, variance, profit, risk, health ---
function injectProjectFormulas(ss) {
  var sheet = ss.getSheetByName('Projects');
  if (!sheet || sheet.getLastRow() <= 1) return;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var lastData = sheet.getLastRow();
  var endRow = lastData + ERP.EXTRA_ROWS;

  var h = {};
  headers.forEach(function(name, idx) { h[name] = idx + 1; }); // 1-indexed

  // Helper: column letter from 1-indexed number
  function cl(colNum) {
    var l = '';
    while (colNum > 0) { var m = (colNum - 1) % 26; l = String.fromCharCode(65 + m) + l; colNum = Math.floor((colNum - 1) / 26); }
    return l;
  }

  var startC = cl(h['start_date']), endC = cl(h['end_date']);
  var valC = cl(h['contract_value_aed']), compC = cl(h['completion_pct']);
  var costC = cl(h['total_cost']), revC = cl(h['total_revenue']);
  var riskC = cl(h['risk_score']);

  for (var r = 2; r <= endRow; r++) {
    var row = r;
    var trigger = cl(h['project_name']) + row; // B column — main trigger

    // duration_days = end - start
    if (h['duration_days']) {
      sheet.getRange(row, h['duration_days']).setFormula(
        '=IF(AND(' + startC + row + '<>"",' + endC + row + '<>""),' + endC + row + '-' + startC + row + ',"")'
      );
    }
    // elapsed_days
    if (h['elapsed_days']) {
      sheet.getRange(row, h['elapsed_days']).setFormula(
        '=IF(AND(' + startC + row + '<>"",' + endC + row + '<>""),MAX(0,MIN(TODAY(),' + endC + row + ')-' + startC + row + '),"")'
      );
    }
    // remaining_days
    if (h['remaining_days']) {
      sheet.getRange(row, h['remaining_days']).setFormula(
        '=IF(' + endC + row + '<>"",MAX(0,' + endC + row + '-TODAY()),"")'
      );
    }
    // budget_variance_pct
    if (h['budget_variance_pct'] && h['total_cost'] && h['contract_value_aed'] && h['completion_pct']) {
      sheet.getRange(row, h['budget_variance_pct']).setFormula(
        '=IF(AND(' + valC + row + '>0,' + compC + row + '>0),ROUND(((' + costC + row + '/(' + valC + row + '*' + compC + row + '/100))-1)*100,2),0)'
      );
    }
    // profit = revenue - cost
    if (h['profit']) {
      sheet.getRange(row, h['profit']).setFormula(
        '=IF(' + revC + row + '<>"",' + revC + row + '-' + costC + row + ',"")'
      );
    }
    // health_status based on risk_score
    if (h['health_status'] && h['risk_score']) {
      sheet.getRange(row, h['health_status']).setFormula(
        '=IF(' + riskC + row + '="","",IF(' + riskC + row + '<=3,"Green",IF(' + riskC + row + '<=6,"Amber","Red")))'
      );
    }
  }
  SpreadsheetApp.flush();
  Logger.log('Project formulas injected');
}

// --- Clients: lifetime_revenue, active_projects, risk_projects ---
function injectClientFormulas(ss) {
  var sheet = ss.getSheetByName('Clients');
  var prjSheet = ss.getSheetByName('Projects');
  if (!sheet || !prjSheet || sheet.getLastRow() <= 1) return;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var prjHeaders = prjSheet.getRange(1, 1, 1, prjSheet.getLastColumn()).getValues()[0];

  var h = {};
  headers.forEach(function(name, idx) { h[name] = idx + 1; });
  function cl(n) { var l=''; while(n>0){var m=(n-1)%26;l=String.fromCharCode(65+m)+l;n=Math.floor((n-1)/26);}return l; }

  var cidCol = cl(h['client_id']);
  var prjCidCol = cl(prjHeaders.indexOf('client_id') + 1);
  var prjRevCol = cl(prjHeaders.indexOf('total_revenue') + 1);
  var prjStatusCol = cl(prjHeaders.indexOf('status') + 1);
  var prjHealthCol = cl(prjHeaders.indexOf('health_status') + 1);
  var prjLast = prjSheet.getLastRow();
  var endRow = sheet.getLastRow() + ERP.EXTRA_ROWS;

  for (var r = 2; r <= endRow; r++) {
    // lifetime_revenue = SUMIF
    if (h['lifetime_revenue']) {
      sheet.getRange(r, h['lifetime_revenue']).setFormula(
        '=IF(' + cidCol + r + '<>"",SUMIF(Projects!' + prjCidCol + '$2:' + prjCidCol + '$' + (prjLast + ERP.EXTRA_ROWS) + ',' + cidCol + r + ',Projects!' + prjRevCol + '$2:' + prjRevCol + '$' + (prjLast + ERP.EXTRA_ROWS) + '),"")'
      );
    }
    // active_projects = COUNTIFS with multiple statuses
    if (h['active_projects']) {
      sheet.getRange(r, h['active_projects']).setFormula(
        '=IF(' + cidCol + r + '<>"",COUNTIFS(Projects!' + prjCidCol + '$2:' + prjCidCol + '$' + (prjLast + ERP.EXTRA_ROWS) + ',' + cidCol + r + ',Projects!' + prjStatusCol + '$2:' + prjStatusCol + '$' + (prjLast + ERP.EXTRA_ROWS) + ',"<>Completed",Projects!' + prjStatusCol + '$2:' + prjStatusCol + '$' + (prjLast + ERP.EXTRA_ROWS) + ',"<>Cancelled"),"")'
      );
    }
    // risk_projects = COUNTIFS health=Red
    if (h['risk_projects']) {
      sheet.getRange(r, h['risk_projects']).setFormula(
        '=IF(' + cidCol + r + '<>"",COUNTIFS(Projects!' + prjCidCol + '$2:' + prjCidCol + '$' + (prjLast + ERP.EXTRA_ROWS) + ',' + cidCol + r + ',Projects!' + prjHealthCol + '$2:' + prjHealthCol + '$' + (prjLast + ERP.EXTRA_ROWS) + ',"Red"),"")'
      );
    }
  }
  SpreadsheetApp.flush();
}

// --- Contracts: remaining_value, payment_progress_pct ---
function injectContractFormulas(ss) {
  var sheet = ss.getSheetByName('Contracts');
  if (!sheet || sheet.getLastRow() <= 1) return;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var h = {};
  headers.forEach(function(name, idx) { h[name] = idx + 1; });
  function cl(n) { var l=''; while(n>0){var m=(n-1)%26;l=String.fromCharCode(65+m)+l;n=Math.floor((n-1)/26);}return l; }

  var valC = cl(h['contract_value_aed']);
  var endRow = sheet.getLastRow() + ERP.EXTRA_ROWS;

  for (var r = 2; r <= endRow; r++) {
    if (h['remaining_value'] && h['payment_progress_pct']) {
      var ppC = cl(h['payment_progress_pct']);
      sheet.getRange(r, h['remaining_value']).setFormula(
        '=IF(' + valC + r + '<>"",' + valC + r + '*(1-' + ppC + r + '/100),"")'
      );
    }
  }
  SpreadsheetApp.flush();
}

// --- Phases: no complex formulas needed beyond auto-ID/timestamps ---
function injectPhaseFormulas(ss) { /* Auto-ID handled in injectAutoIDFormulas */ }

// --- Work Packages: total_amount, executed_amount ---
function injectWorkPackageFormulas(ss) {
  var sheet = ss.getSheetByName('Work_Packages');
  if (!sheet || sheet.getLastRow() <= 1) return;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var h = {};
  headers.forEach(function(name, idx) { h[name] = idx + 1; });
  function cl(n) { var l=''; while(n>0){var m=(n-1)%26;l=String.fromCharCode(65+m)+l;n=Math.floor((n-1)/26);}return l; }

  var qtyC = cl(h['quantity']), rateC = cl(h['unit_rate_aed']);
  var exQtyC = cl(h['executed_qty']);
  var endRow = sheet.getLastRow() + ERP.EXTRA_ROWS;

  for (var r = 2; r <= endRow; r++) {
    // total_amount = qty × rate
    if (h['total_amount_aed']) {
      sheet.getRange(r, h['total_amount_aed']).setFormula(
        '=IF(AND(' + qtyC + r + '<>"",' + rateC + r + '<>""),' + qtyC + r + '*' + rateC + r + ',"")'
      );
    }
    // executed_amount = exec_qty × rate
    if (h['executed_amount_aed']) {
      sheet.getRange(r, h['executed_amount_aed']).setFormula(
        '=IF(AND(' + exQtyC + r + '<>"",' + rateC + r + '<>""),' + exQtyC + r + '*' + rateC + r + ',"")'
      );
    }
  }
  SpreadsheetApp.flush();
}

// --- Permits: is_expired formula ---
function injectPermitFormulas(ss) {
  var sheet = ss.getSheetByName('Permits_Approvals');
  if (!sheet || sheet.getLastRow() <= 1) return;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var h = {};
  headers.forEach(function(name, idx) { h[name] = idx + 1; });
  function cl(n) { var l=''; while(n>0){var m=(n-1)%26;l=String.fromCharCode(65+m)+l;n=Math.floor((n-1)/26);}return l; }

  var expC = cl(h['expiry_date']);
  var endRow = sheet.getLastRow() + ERP.EXTRA_ROWS;

  for (var r = 2; r <= endRow; r++) {
    if (h['is_expired']) {
      sheet.getRange(r, h['is_expired']).setFormula(
        '=IF(' + expC + r + '<>"",IF(TODAY()>' + expC + r + ',TRUE,FALSE),"")'
      );
    }
  }
  SpreadsheetApp.flush();
}

// --- Payments: remaining_balance ---
function injectPaymentFormulas(ss) {
  var sheet = ss.getSheetByName('Payment_Applications');
  if (!sheet || sheet.getLastRow() <= 1) return;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var h = {};
  headers.forEach(function(name, idx) { h[name] = idx + 1; });
  function cl(n) { var l=''; while(n>0){var m=(n-1)%26;l=String.fromCharCode(65+m)+l;n=Math.floor((n-1)/26);}return l; }

  var grossC = cl(h['gross_value_aed']), cumC = cl(h['cumulative_value_aed']);
  var retC = cl(h['retention_aed']), dedC = cl(h['deductions_aed']);
  var endRow = sheet.getLastRow() + ERP.EXTRA_ROWS;

  for (var r = 2; r <= endRow; r++) {
    // net_certified = gross - retention - deductions
    if (h['net_certified_aed']) {
      sheet.getRange(r, h['net_certified_aed']).setFormula(
        '=IF(' + grossC + r + '<>"",' + grossC + r + '-IF(' + retC + r + '<>"",' + retC + r + ',0)-IF(' + dedC + r + '<>"",' + dedC + r + ',0),"")'
      );
    }
  }
  SpreadsheetApp.flush();
}

// --- Auto-ID formulas for all sheets (new rows only) ---
function injectAutoIDFormulas(ss) {
  for (var sheetName in ERP.SHEETS) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    var cfg = ERP.SHEETS[sheetName];
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var pkIdx = headers.indexOf(cfg.pk);
    if (pkIdx < 0) continue;

    // Find a "trigger" column — the main data column user would fill first
    var triggerCols = {
      'Clients':'client_name', 'Employees':'full_name', 'Contractors':'company_name',
      'Suppliers':'company_name', 'Equipment':'equipment_type', 'Projects':'project_name',
      'Contracts':'contract_type', 'Project_Phases':'phase_name', 'Work_Packages':'item_description',
      'Permits_Approvals':'permit_type', 'Inspections':'inspection_type', 'Safety_Incidents':'incident_type',
      'Payment_Applications':'ipc_number', 'Variation_Orders':'vo_number', 'Purchase_Orders':'po_type',
      'Daily_Site_Reports':'report_date', 'Project_Documents':'document_type'
    };
    var trigCol = triggerCols[sheetName] || headers[1]; // Default to 2nd column
    var trigIdx = headers.indexOf(trigCol);
    if (trigIdx < 0) trigIdx = 1; // Fallback

    function cl(n) { var l=''; while(n>0){var m=(n-1)%26;l=String.fromCharCode(65+m)+l;n=Math.floor((n-1)/26);}return l; }

    var pkLetter = cl(pkIdx + 1);
    var trigLetter = cl(trigIdx + 1);
    var padStr = new Array(cfg.pad + 1).join('0'); // e.g. "000" or "0000"

    // Only inject into NEW rows (below existing data)
    var dataEnd = sheet.getLastRow();
    var endRow = dataEnd + ERP.EXTRA_ROWS;

    for (var r = dataEnd + 1; r <= endRow; r++) {
      // Auto-ID: =IF(triggerCol<>"", PREFIX & TEXT(MAX(existing IDs)+ROW-dataEnd, "000"), "")
      sheet.getRange(r, pkIdx + 1).setFormula(
        '=IF(' + trigLetter + r + '<>"","' + cfg.prefix + '"&TEXT(COUNTA(' + pkLetter + '$2:' + pkLetter + (r - 1) + ')+1,"' + padStr + '"),"")'
      );
    }
    SpreadsheetApp.flush();
  }
}

// --- Timestamp formulas for new rows ---
function injectTimestampFormulas(ss) {
  for (var sheetName in ERP.SHEETS) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var cfg = ERP.SHEETS[sheetName];

    var triggerCols = {
      'Clients':'client_name', 'Employees':'full_name', 'Contractors':'company_name',
      'Suppliers':'company_name', 'Equipment':'equipment_type', 'Projects':'project_name',
      'Contracts':'contract_type', 'Project_Phases':'phase_name', 'Work_Packages':'item_description',
      'Permits_Approvals':'permit_type', 'Inspections':'inspection_type', 'Safety_Incidents':'incident_type',
      'Payment_Applications':'ipc_number', 'Variation_Orders':'vo_number', 'Purchase_Orders':'po_type',
      'Daily_Site_Reports':'report_date', 'Project_Documents':'document_type'
    };
    var trigCol = triggerCols[sheetName] || headers[1];
    var trigIdx = headers.indexOf(trigCol);
    if (trigIdx < 0) trigIdx = 1;
    function cl(n) { var l=''; while(n>0){var m=(n-1)%26;l=String.fromCharCode(65+m)+l;n=Math.floor((n-1)/26);}return l; }
    var trigLetter = cl(trigIdx + 1);

    var catIdx = headers.indexOf('created_at');
    var uatIdx = headers.indexOf('last_updated_at');
    var cbIdx = headers.indexOf('created_by');
    var rvIdx = headers.indexOf('record_version');
    var iaIdx = headers.indexOf('is_active');

    var dataEnd = sheet.getLastRow();
    var endRow = dataEnd + ERP.EXTRA_ROWS;

    for (var r = dataEnd + 1; r <= endRow; r++) {
      if (catIdx >= 0) {
        sheet.getRange(r, catIdx + 1).setFormula('=IF(' + trigLetter + r + '<>"",NOW(),"")');
      }
      if (uatIdx >= 0) {
        sheet.getRange(r, uatIdx + 1).setFormula('=IF(' + trigLetter + r + '<>"",NOW(),"")');
      }
      if (cbIdx >= 0) {
        sheet.getRange(r, cbIdx + 1).setFormula('=IF(' + trigLetter + r + '<>"","' + ERP.ADMIN + '","")');
      }
      if (rvIdx >= 0) {
        sheet.getRange(r, rvIdx + 1).setFormula('=IF(' + trigLetter + r + '<>"",1,"")');
      }
      if (iaIdx >= 0) {
        sheet.getRange(r, iaIdx + 1).setFormula('=IF(' + trigLetter + r + '<>"",TRUE,"")');
      }
    }
    SpreadsheetApp.flush();
    Utilities.sleep(200);
  }
}

// ============================================================
// MODULE 6: STATUS CELL COLORING
// ============================================================
function colorAllStatusCells(ss) {
  for (var sheetName in ERP.SHEETS) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var data = sheet.getDataRange().getValues();
    var cfg = ERP.SHEETS[sheetName];

    // Find status/result/severity/health columns
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
    sheet.clearConditionalFormatRules();
    SpreadsheetApp.flush();
    Utilities.sleep(100);
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
