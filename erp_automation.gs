/**
 * ============================================================
 * DUBAI ERP - ENTERPRISE AUTOMATION ENGINE
 * ============================================================
 * Production-grade Google Apps Script for civil project ERP
 * 
 * MODULES:
 *  1. System Setup & Configuration
 *  2. Edit Triggers (onEdit/onChange) — timestamps, versioning, audit
 *  3. Insert Handling — auto PK, defaults, tags
 *  4. Delete Handling — soft delete
 *  5. Validation Engine — per-table business rules
 *  6. Conditional Formatting — visual intelligence
 *  7. Derived Intelligence — auto-calculations
 *  8. Smart Tagging — dynamic tag engine
 *  9. Privacy & Security — column protection, role-based access
 * 10. Audit & Governance — hidden audit log
 *
 * INSTRUCTIONS:
 *  1. Open Google Sheets → Extensions → Apps Script
 *  2. Paste this code into a new file called "erp_automation.gs"
 *  3. Run setupERPSystem() once to initialize
 *  4. All automation runs automatically via triggers
 *
 * Admin: anandhu7833@gmail.com
 * Execution Limit: 15 minutes
 */

// ============================================================
// MODULE 1: CONFIGURATION
// ============================================================
var ERP_CONFIG = {
  ADMIN_EMAIL: 'anandhu7833@gmail.com',
  SYSTEM_USER: 'system@alrashdeng.ae',
  AUDIT_SHEET: 'Audit_Log',
  
  // Sheet configurations: name, PK column header, PK prefix, default status
  SHEETS: {
    'Clients':              { pk: 'client_id',     prefix: 'CL-',   defaultStatus: 'Active',      statusCol: 'status' },
    'Employees':            { pk: 'employee_id',   prefix: 'EMP-',  defaultStatus: 'Active',      statusCol: 'status' },
    'Contractors':          { pk: 'contractor_id', prefix: 'CON-',  defaultStatus: 'Active',      statusCol: 'status' },
    'Suppliers':            { pk: 'supplier_id',   prefix: 'SUP-',  defaultStatus: 'Active',      statusCol: 'status' },
    'Equipment':            { pk: 'equipment_id',  prefix: 'EQ-',   defaultStatus: 'Available',   statusCol: 'status' },
    'Projects':             { pk: 'project_id',    prefix: 'PRJ-',  defaultStatus: 'Awarded',     statusCol: 'status' },
    'Contracts':            { pk: 'contract_id',   prefix: 'MC-',   defaultStatus: 'Active',      statusCol: 'status' },
    'Project_Phases':       { pk: 'phase_id',      prefix: 'PH-',   defaultStatus: 'Not Started', statusCol: 'status' },
    'Work_Packages':        { pk: 'package_id',    prefix: 'WP-',   defaultStatus: 'Active',      statusCol: null },
    'Permits_Approvals':    { pk: 'permit_id',     prefix: 'PM-',   defaultStatus: 'Pending',     statusCol: 'status' },
    'Inspections':          { pk: 'inspection_id', prefix: 'INS-',  defaultStatus: 'Pending',     statusCol: 'result' },
    'Safety_Incidents':     { pk: 'incident_id',   prefix: 'INC-',  defaultStatus: 'Open',        statusCol: 'status' },
    'Payment_Applications': { pk: 'ipc_id',        prefix: 'IPC-',  defaultStatus: 'Submitted',   statusCol: 'status' },
    'Variation_Orders':     { pk: 'vo_id',         prefix: 'VO-',   defaultStatus: 'Draft',       statusCol: 'status' },
    'Purchase_Orders':      { pk: 'po_id',         prefix: 'LPO-',  defaultStatus: 'Draft',       statusCol: 'status' },
    'Daily_Site_Reports':   { pk: 'report_id',     prefix: 'DSR-',  defaultStatus: null,          statusCol: null },
    'Project_Documents':    { pk: 'document_id',   prefix: 'DOC-',  defaultStatus: 'Draft',       statusCol: 'status' }
  },

  // Protected columns (cannot be manually edited)
  PROTECTED_COLS: ['created_at', 'created_by', 'record_version'],

  // Private/sensitive columns (admin-only, gray shading)
  PRIVATE_COLS: ['salary_aed', 'contract_value_aed', 'gross_value_aed', 'net_certified_aed',
                 'amount_aed', 'risk_score', 'internal_notes', 'credit_limit_aed'],

  // FK relationships for validation
  FK_RULES: {
    'Projects':             [{ col: 'client_id', ref: 'Clients', refCol: 'client_id' },
                             { col: 'project_manager_id', ref: 'Employees', refCol: 'employee_id' }],
    'Contracts':            [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
    'Project_Phases':       [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
    'Work_Packages':        [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
    'Permits_Approvals':    [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
    'Inspections':          [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
    'Safety_Incidents':     [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
    'Payment_Applications': [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
    'Variation_Orders':     [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
    'Purchase_Orders':      [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
    'Daily_Site_Reports':   [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
    'Project_Documents':    [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }]
  },

  // Colorblind-safe palette (Wong/Okabe-Ito inspired, WCAG AA contrast)
  COLORS: {
    HEADER_BG: '#2d3748', HEADER_FG: '#f7fafc',
    ALT_ROW: '#f7fafc', WHITE: '#ffffff',
    SYSTEM_COL_BG: '#edf2f7', PRIVATE_COL_BG: '#fef5e7',
    MISSING_BG: '#fff3cd',
    // Status cell badges (colorblind-safe: distinct hue + lightness)
    STATUS: {
      // Blue family (positive)
      'Completed': {bg:'#d4edda',fg:'#155724'}, 'Active': {bg:'#d4edda',fg:'#155724'},
      'Paid': {bg:'#d4edda',fg:'#155724'}, 'Approved': {bg:'#d4edda',fg:'#155724'},
      'Pass': {bg:'#d4edda',fg:'#155724'},
      // Cyan/Blue (in progress)
      'In Progress': {bg:'#cce5ff',fg:'#004085'}, 'Certified': {bg:'#cce5ff',fg:'#004085'},
      'Under Review': {bg:'#cce5ff',fg:'#004085'}, 'Issued': {bg:'#cce5ff',fg:'#004085'},
      'Delivered': {bg:'#cce5ff',fg:'#004085'},
      // Yellow (caution)
      'On Hold': {bg:'#fff3cd',fg:'#856404'}, 'Submitted': {bg:'#fff3cd',fg:'#856404'},
      'Pending': {bg:'#fff3cd',fg:'#856404'}, 'Draft': {bg:'#fff3cd',fg:'#856404'},
      'Under Negotiation': {bg:'#fff3cd',fg:'#856404'}, 'For Review': {bg:'#fff3cd',fg:'#856404'},
      // Orange (warning)
      'Delayed': {bg:'#f8d7da',fg:'#721c24'}, 'Overdue': {bg:'#f8d7da',fg:'#721c24'},
      'Rejected': {bg:'#f8d7da',fg:'#721c24'}, 'Fail': {bg:'#f8d7da',fg:'#721c24'},
      'Expired': {bg:'#f8d7da',fg:'#721c24'}, 'Blacklisted': {bg:'#f8d7da',fg:'#721c24'},
      // Gray (neutral)
      'Cancelled': {bg:'#e2e3e5',fg:'#383d41'}, 'Inactive': {bg:'#e2e3e5',fg:'#383d41'},
      'Not in Contact': {bg:'#e2e3e5',fg:'#383d41'}, 'Superseded': {bg:'#e2e3e5',fg:'#383d41'},
      // Purple (awarded/special)
      'Awarded': {bg:'#e8daef',fg:'#4a235a'}, 'Not Started': {bg:'#e8daef',fg:'#4a235a'},
      'Invoiced': {bg:'#d6eaf8',fg:'#1b4f72'},
      // Safety severity (distinct shapes reinforced by text prefix)
      'High': {bg:'#f8d7da',fg:'#721c24'}, 'Medium': {bg:'#fff3cd',fg:'#856404'},
      'Low': {bg:'#d4edda',fg:'#155724'},
      // Health
      'Green': {bg:'#d4edda',fg:'#155724'}, 'Amber': {bg:'#fff3cd',fg:'#856404'},
      'Red': {bg:'#f8d7da',fg:'#721c24'},
      // Conditional pass
      'Conditional Pass': {bg:'#fff3cd',fg:'#856404'}, 'Open': {bg:'#fff3cd',fg:'#856404'},
      'Closed': {bg:'#d4edda',fg:'#155724'}, 'Available': {bg:'#d4edda',fg:'#155724'},
      'In Use': {bg:'#cce5ff',fg:'#004085'}, 'Under Maintenance': {bg:'#fff3cd',fg:'#856404'},
      'Decommissioned': {bg:'#e2e3e5',fg:'#383d41'}
    }
  },

  // Dropdown definitions: { sheetName: { columnHeader: [values] } }
  DROPDOWNS: {
    'Clients': {
      'client_type': ['Private Developer','Government','Semi-Government','Contractor'],
      'status': ['Active','Inactive','Not in Contact','Blacklisted'],
      'payment_terms': ['Net 30','Net 45','Net 60','Net 90']
    },
    'Employees': {
      'department': ['Engineering','Construction','HSE','Finance','HR','Procurement','QA/QC','Document Control','Admin','IT'],
      'status': ['Active','On Leave','Resigned','Terminated'],
      'visa_status': ['Valid','Expiring Soon','Expired','In Process','Cancelled']
    },
    'Contractors': {
      'specialty': ['Concrete Works','Demolition','Electrical','Facade & Cladding','Fire Fighting','Fit-Out','HVAC','Landscaping','MEP','Painting','Piling & Foundations','Plumbing','Structural Steel','Waterproofing'],
      'status': ['Active','Suspended','Blacklisted'],
      'prequalified': ['TRUE','FALSE'],
      'rating': ['A','B','C','D']
    },
    'Suppliers': {
      'category': ['Cement & Aggregates','Construction Chemicals','Facade Materials','Finishing Materials','Formwork & Scaffolding','MEP Materials','Ready-Mix Concrete','Safety Equipment','Steel Reinforcement','Structural Steel'],
      'status': ['Active','Suspended','Inactive'],
      'approved': ['TRUE','FALSE']
    },
    'Equipment': {
      'equipment_type': ['Backhoe Loader','Batching Plant','Boom Lift','Compactor','Concrete Pump','Dump Truck','Excavator','Forklift','Generator','Mobile Crane','Scaffolding Set','Tower Crane','Welding Machine','Wheel Loader'],
      'status': ['Available','In Use','Under Maintenance','Decommissioned']
    },
    'Projects': {
      'project_type': ['Commercial Office Building','High-Rise Residential Tower','Hotel & Hospitality','Industrial Warehouse','Infrastructure - Bridges','Infrastructure - Roads','Marine & Coastal Works','Mixed-Use Development','Retail Mall','School/Educational','Utility Infrastructure','Villa Development'],
      'status': ['Awarded','In Progress','On Hold','Delayed','Completed','Cancelled'],
      'health_status': ['Green','Amber','Red']
    },
    'Contracts': {
      'contract_type': ['Main Contract','Subcontract'],
      'status': ['Active','Completed','Terminated','Rejected']
    },
    'Project_Phases': {
      'phase_name': ['Mobilization','Enabling Works','Piling','Substructure','Superstructure','MEP Rough-In','Facade','Internal Finishes','External Works','Testing & Commissioning'],
      'status': ['Not Started','In Progress','Completed','Delayed']
    },
    'Permits_Approvals': {
      'permit_type': ['Dubai Municipality - Building Permit','Dubai Municipality - NOC','Civil Defense NOC','DEWA Connection','RTA Access Permit','Crane Permit','DDA Approval','Environmental Permit','Etisalat NOC','Trakhees Permit'],
      'status': ['Pending','Submitted','Approved','Expired','Rejected']
    },
    'Inspections': {
      'inspection_type': ['Concrete Pour','Rebar Placement','Formwork','Structural','Waterproofing','MEP Rough-In','Fire Safety','Final Handover'],
      'result': ['Pass','Fail','Conditional Pass','Pending']
    },
    'Safety_Incidents': {
      'incident_type': ['Fall from Height','Struck by Object','Electrical','Heat Stroke','Equipment Malfunction','Caught In/Between','Slip/Trip','Chemical Exposure','Fire','Vehicle Accident'],
      'severity': ['Low','Medium','High'],
      'status': ['Open','Under Investigation','Closed']
    },
    'Payment_Applications': {
      'status': ['Submitted','Under Review','Certified','Paid','Rejected','Overdue']
    },
    'Variation_Orders': {
      'reason': ['Client Request','Design Change','Site Condition','Authority Requirement','Scope Addition','Material Substitution'],
      'status': ['Draft','Submitted','Under Negotiation','Approved','Rejected']
    },
    'Purchase_Orders': {
      'po_type': ['Material','Subcontract'],
      'status': ['Draft','Issued','Delivered','Invoiced','Paid']
    },
    'Daily_Site_Reports': {
      'weather': ['Clear','Sunny','Hot','Humid','Windy','Sandstorm','Rain']
    },
    'Project_Documents': {
      'document_type': ['Shop Drawing','Material Submittal','RFI','Method Statement','Inspection Request','NCR','Progress Report','Meeting Minutes','Site Instruction','Test Report','As-Built Drawing','Variation Order'],
      'status': ['Draft','For Review','Approved','Superseded'],
      'discipline': ['Civil','Structural','Architectural','MEP','HSE']
    }
  }
};

// ============================================================
// MODULE 1: SYSTEM SETUP
// ============================================================

/**
 * Main setup function — run once to initialize the ERP system
 */
function setupERPSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  ss.toast('Setting up Enterprise ERP System...', '⚙️ Setup', -1);
  
  // 1. Create Audit Log sheet
  createAuditLogSheet(ss);
  ss.toast('1/9 Audit Log created', '⚙️ Setup', -1);
  
  // 2. Install triggers
  installTriggers(ss);
  ss.toast('2/9 Triggers installed', '⚙️ Setup', -1);
  
  // 3. Apply base styling (headers, alternating rows, column widths)
  applyBaseStyleToAllSheets(ss);
  ss.toast('3/9 Base styling applied', '⚙️ Setup', -1);
  
  // 4. Apply dropdown data validation
  applyAllDropdowns(ss);
  ss.toast('4/9 Dropdowns applied', '⚙️ Setup', -1);
  
  // 5. Handle missing data
  handleAllMissingData(ss);
  ss.toast('5/9 Missing data handled', '⚙️ Setup', -1);
  
  // 6. Apply status cell formatting (colorblind-safe)
  applyAllStatusFormatting(ss);
  ss.toast('6/9 Status formatting applied', '⚙️ Setup', -1);
  
  // 7. Apply privacy protections
  applyPrivacyProtections(ss);
  ss.toast('7/9 Privacy protections applied', '⚙️ Setup', -1);
  
  // 8. Run derived calculations
  recalculateAllDerived(ss);
  ss.toast('8/9 Derived calculations complete', '⚙️ Setup', -1);
  
  // 9. Auto-resize columns
  autoResizeAllColumns(ss);
  ss.toast('9/9 Columns resized', '⚙️ Setup', -1);
  
  ss.toast('✅ ERP System Setup Complete!', '⚙️ Done', 10);
  ui.alert('ERP System Setup Complete!\n\n✅ Audit Log\n✅ Edit Triggers\n✅ Elegant Base Styling\n✅ Dropdown Validation (all categorical fields)\n✅ Missing Data Handling\n✅ Colorblind-Safe Status Formatting\n✅ Privacy Protections\n✅ Derived Calculations\n✅ Auto Column Widths');
}

function createAuditLogSheet(ss) {
  var sheet = ss.getSheetByName(ERP_CONFIG.AUDIT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ERP_CONFIG.AUDIT_SHEET);
  } else {
    sheet.clear();
  }
  var headers = ['Timestamp', 'User', 'Action', 'Table', 'RowID', 'Field', 'Old_Value', 'New_Value'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#e0e0e0');
  sheet.setFrozenRows(1);
  sheet.hideSheet();
  sheet.autoResizeColumns(1, headers.length);
}

function installTriggers(ss) {
  // Remove existing triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) { ScriptApp.deleteTrigger(t); });
  
  // Install onEdit trigger (simple — runs on every edit)
  ScriptApp.newTrigger('onEditTrigger')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  
  // Install time-based trigger for auto-delayed detection (runs hourly)
  ScriptApp.newTrigger('autoDelayedDetection')
    .timeBased()
    .everyHours(1)
    .create();
  
  // Install time-based trigger for permit expiry check (runs daily)
  ScriptApp.newTrigger('autoPermitExpiryCheck')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  
  Logger.log('Triggers installed successfully');
}

// ============================================================
// MODULE 2: EDIT TRIGGERS
// ============================================================

/**
 * Main onEdit handler — runs on every cell edit
 */
function onEditTrigger(e) {
  if (!e || !e.range) return;
  
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  
  // Skip system sheets
  if (sheetName === ERP_CONFIG.AUDIT_SHEET || sheetName === 'Import_Summary' || 
      sheetName === 'Import_Log' || sheetName === 'Temp') return;
  
  var config = ERP_CONFIG.SHEETS[sheetName];
  if (!config) return; // Not a managed sheet
  
  var row = e.range.getRow();
  var col = e.range.getColumn();
  
  // Skip header row
  if (row === 1) return;
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var editedHeader = headers[col - 1];
  var user = Session.getActiveUser().getEmail() || ERP_CONFIG.SYSTEM_USER;
  
  // Check if editing a protected column (non-admin)
  if (ERP_CONFIG.PROTECTED_COLS.indexOf(editedHeader) >= 0 && user !== ERP_CONFIG.ADMIN_EMAIL) {
    e.range.setValue(e.oldValue || '');
    SpreadsheetApp.getActiveSpreadsheet().toast('❌ ' + editedHeader + ' is read-only', 'Protected Field', 3);
    return;
  }
  
  // Check if editing PK column (never allowed)
  if (editedHeader === config.pk) {
    e.range.setValue(e.oldValue || '');
    SpreadsheetApp.getActiveSpreadsheet().toast('❌ Primary Key cannot be modified', 'Protected Field', 3);
    return;
  }
  
  // Detect new row insertion (PK cell is empty)
  var pkColIdx = headers.indexOf(config.pk);
  if (pkColIdx >= 0) {
    var pkValue = sheet.getRange(row, pkColIdx + 1).getValue();
    if (!pkValue || pkValue === '') {
      handleNewRow(sheet, sheetName, row, headers, config, user);
      return;
    }
  }
  
  // Run validation for the edited field
  var validationResult = validateEdit(sheetName, editedHeader, e.range.getValue(), sheet, row, headers);
  if (validationResult !== true) {
    e.range.setValue(e.oldValue || '');
    SpreadsheetApp.getActiveSpreadsheet().toast('❌ ' + validationResult, 'Validation Error', 5);
    return;
  }
  
  // Update timestamps and version
  var updatedAtIdx = headers.indexOf('last_updated_at');
  if (updatedAtIdx >= 0) {
    sheet.getRange(row, updatedAtIdx + 1).setValue(new Date().toISOString());
  }
  
  var versionIdx = headers.indexOf('record_version');
  if (versionIdx >= 0) {
    var currentVersion = sheet.getRange(row, versionIdx + 1).getValue();
    sheet.getRange(row, versionIdx + 1).setValue((parseInt(currentVersion) || 0) + 1);
  }
  
  // Log to audit
  var pkVal = pkColIdx >= 0 ? sheet.getRange(row, pkColIdx + 1).getValue() : 'ROW-' + row;
  logAudit(user, 'UPDATE', sheetName, pkVal, editedHeader, e.oldValue, e.range.getValue());
  
  // Recalculate derived fields if needed
  recalculateDerivedForRow(sheet, sheetName, row, headers, editedHeader);
  
  // Update tags
  updateTagsForRow(sheet, sheetName, row, headers);
}

// ============================================================
// MODULE 3: INSERT HANDLING
// ============================================================

function handleNewRow(sheet, sheetName, row, headers, config, user) {
  var now = new Date().toISOString();
  
  // Auto-generate PK
  var pkColIdx = headers.indexOf(config.pk);
  if (pkColIdx >= 0) {
    var newPk = generateNextPK(sheet, headers, config);
    sheet.getRange(row, pkColIdx + 1).setValue(newPk);
  }
  
  // Auto-set governance fields
  setIfExists(sheet, row, headers, 'created_at', now);
  setIfExists(sheet, row, headers, 'last_updated_at', now);
  setIfExists(sheet, row, headers, 'created_by', user);
  setIfExists(sheet, row, headers, 'record_version', 1);
  setIfExists(sheet, row, headers, 'is_active', true);
  
  // Auto-set default status
  if (config.statusCol && config.defaultStatus) {
    setIfExists(sheet, row, headers, config.statusCol, config.defaultStatus);
  }
  
  // Auto-suggest tags
  updateTagsForRow(sheet, sheetName, row, headers);
  
  // Log insertion
  var pkVal = pkColIdx >= 0 ? sheet.getRange(row, pkColIdx + 1).getValue() : 'ROW-' + row;
  logAudit(user, 'INSERT', sheetName, pkVal, '', '', 'New record created');
  
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ New record created: ' + (pkVal || 'ROW-' + row), sheetName, 3);
}

function generateNextPK(sheet, headers, config) {
  var pkColIdx = headers.indexOf(config.pk);
  if (pkColIdx < 0) return '';
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return config.prefix + '001';
  
  var pkValues = sheet.getRange(2, pkColIdx + 1, lastRow - 1, 1).getValues();
  var maxNum = 0;
  
  for (var i = 0; i < pkValues.length; i++) {
    var val = String(pkValues[i][0]);
    // Extract numeric part from PK
    var match = val.match(/(\d+)$/);
    if (match) {
      var num = parseInt(match[1]);
      if (num > maxNum) maxNum = num;
    }
  }
  
  var nextNum = maxNum + 1;
  var padLen = config.prefix === 'EMP-' ? 4 : (config.prefix === 'WP-' || config.prefix === 'EQ-' ? 4 : 3);
  return config.prefix + String(nextNum).padStart(padLen, '0');
}

function setIfExists(sheet, row, headers, colName, value) {
  var idx = headers.indexOf(colName);
  if (idx >= 0) {
    sheet.getRange(row, idx + 1).setValue(value);
  }
}

// ============================================================
// MODULE 4: DELETE HANDLING (SOFT DELETE)
// ============================================================

/**
 * Convert physical delete to soft delete
 * Called from a custom menu or can intercept onChange
 */
function softDeleteRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getName();
  var config = ERP_CONFIG.SHEETS[sheetName];
  if (!config) { ss.toast('Not a managed table', 'Error', 3); return; }
  
  var row = sheet.getActiveRange().getRow();
  if (row <= 1) { ss.toast('Cannot delete header row', 'Error', 3); return; }
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var user = Session.getActiveUser().getEmail() || ERP_CONFIG.SYSTEM_USER;
  
  // Set IsActive = FALSE
  var isActiveIdx = headers.indexOf('is_active');
  if (isActiveIdx >= 0) {
    sheet.getRange(row, isActiveIdx + 1).setValue(false);
  }
  
  // Update timestamp
  setIfExists(sheet, row, headers, 'last_updated_at', new Date().toISOString());
  
  // Increment version
  var versionIdx = headers.indexOf('record_version');
  if (versionIdx >= 0) {
    var ver = sheet.getRange(row, versionIdx + 1).getValue();
    sheet.getRange(row, versionIdx + 1).setValue((parseInt(ver) || 0) + 1);
  }
  
  // Log deletion
  var pkIdx = headers.indexOf(config.pk);
  var pkVal = pkIdx >= 0 ? sheet.getRange(row, pkIdx + 1).getValue() : 'ROW-' + row;
  logAudit(user, 'SOFT_DELETE', sheetName, pkVal, 'is_active', 'TRUE', 'FALSE');
  
  // Gray out the row
  sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground('#d5d8dc').setFontColor('#95a5a6');
  
  ss.toast('🗑️ Record soft-deleted: ' + pkVal, sheetName, 3);
}

// ============================================================
// MODULE 5: VALIDATION ENGINE
// ============================================================

function validateEdit(sheetName, column, value, sheet, row, headers) {
  if (value === '' || value === null || value === undefined) return true;
  
  switch (sheetName) {
    case 'Clients':
      return validateClient(column, value, sheet, row, headers);
    case 'Projects':
      return validateProject(column, value, sheet, row, headers);
    case 'Contracts':
      return validateContract(column, value, sheet, row, headers);
    case 'Payment_Applications':
      return validatePayment(column, value, sheet, row, headers);
    case 'Inspections':
      return validateInspection(column, value, sheet, row, headers);
    case 'Safety_Incidents':
      return validateSafetyIncident(column, value, sheet, row, headers);
    case 'Permits_Approvals':
      return validatePermit(column, value, sheet, row, headers);
    default:
      return validateForeignKeys(sheetName, column, value);
  }
}

function validateClient(column, value, sheet, row, headers) {
  switch (column) {
    case 'client_name':
      if (!value || String(value).trim() === '') return 'Client name cannot be empty';
      // Check uniqueness
      var nameCol = headers.indexOf('client_name') + 1;
      var names = sheet.getRange(2, nameCol, sheet.getLastRow() - 1, 1).getValues();
      var count = 0;
      names.forEach(function(n) { if (String(n[0]).toLowerCase() === String(value).toLowerCase()) count++; });
      if (count > 1) return 'Client name must be unique';
      return true;
    
    case 'email':
      var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (!emailRegex.test(String(value))) return 'Invalid email format';
      return true;
    
    case 'phone':
      var phoneRegex = /^\+971-\d{2}-\d{7}$/;
      if (!phoneRegex.test(String(value))) return 'Phone must be +971-XX-XXXXXXX format';
      return true;
    
    case 'status':
      var validStatuses = ['Active', 'Inactive', 'Not in Contact', 'Blacklisted'];
      if (validStatuses.indexOf(value) === -1) return 'Status must be: ' + validStatuses.join(', ');
      return true;
    
    default:
      return true;
  }
}

function validateProject(column, value, sheet, row, headers) {
  switch (column) {
    case 'contract_value_aed':
      if (isNaN(value) || Number(value) <= 0) return 'Budget must be > 0';
      return true;
    
    case 'start_date':
      var endIdx = headers.indexOf('end_date');
      if (endIdx >= 0) {
        var endDate = new Date(sheet.getRange(row, endIdx + 1).getValue());
        if (new Date(value) > endDate) return 'Start date must be ≤ End date';
      }
      return true;
    
    case 'end_date':
      var startIdx = headers.indexOf('start_date');
      if (startIdx >= 0) {
        var startDate = new Date(sheet.getRange(row, startIdx + 1).getValue());
        if (new Date(value) < startDate) return 'End date must be ≥ Start date';
      }
      return true;
    
    case 'status':
      if (value === 'Completed') {
        // Cannot complete if work packages are incomplete
        var projId = getRowValue(sheet, row, headers, 'project_id');
        if (projId && hasIncompleteWorkPackages(projId)) {
          return 'Cannot mark Completed: incomplete work packages exist';
        }
      }
      if (value === 'Cancelled') {
        var projId2 = getRowValue(sheet, row, headers, 'project_id');
        if (projId2 && hasPayments(projId2)) {
          return 'Cannot cancel: payments exist for this project';
        }
      }
      var validStatuses = ['Awarded', 'In Progress', 'On Hold', 'Delayed', 'Completed', 'Cancelled'];
      if (validStatuses.indexOf(value) === -1) return 'Invalid project status';
      return true;
    
    case 'client_id':
      return validateFKValue('Clients', 'client_id', value);
    
    case 'project_manager_id':
      return validateFKValue('Employees', 'employee_id', value);
    
    default:
      return true;
  }
}

function validateContract(column, value, sheet, row, headers) {
  switch (column) {
    case 'contract_value_aed':
      if (isNaN(value) || Number(value) <= 0) return 'Contract value must be > 0';
      return true;
    case 'contractor_id':
      if (value && value !== '') return validateFKValue('Contractors', 'contractor_id', value);
      return true;
    case 'project_id':
      return validateFKValue('Projects', 'project_id', value);
    default:
      return true;
  }
}

function validatePayment(column, value, sheet, row, headers) {
  switch (column) {
    case 'gross_value_aed':
    case 'net_certified_aed':
      if (isNaN(value) || Number(value) < 0) return 'Payment amount must be ≥ 0';
      // Check against contract value
      var projId = getRowValue(sheet, row, headers, 'project_id');
      if (projId) {
        var contractValue = getProjectContractValue(projId);
        var cumIdx = headers.indexOf('cumulative_value_aed');
        if (cumIdx >= 0) {
          var cumulative = Number(sheet.getRange(row, cumIdx + 1).getValue()) || 0;
          if (cumulative > contractValue * 1.01) {
            return 'Cumulative payments exceed contract value (' + contractValue + ')';
          }
        }
      }
      return true;
    
    case 'status':
      // Check that contract is approved before payment
      var projId2 = getRowValue(sheet, row, headers, 'project_id');
      if (value === 'Paid' || value === 'Certified') {
        if (projId2 && !hasApprovedContract(projId2)) {
          return 'Cannot process payment: no approved contract exists';
        }
      }
      return true;
    
    default:
      return true;
  }
}

function validateInspection(column, value, sheet, row, headers) {
  if (column === 'inspection_date') {
    var inspDate = new Date(value);
    if (inspDate > new Date()) return 'Inspection date cannot be in the future';
  }
  return true;
}

function validateSafetyIncident(column, value, sheet, row, headers) {
  if (column === 'severity') {
    var validSeverities = ['Low', 'Medium', 'High'];
    if (validSeverities.indexOf(value) === -1) return 'Severity must be: Low, Medium, or High';
  }
  return true;
}

function validatePermit(column, value, sheet, row, headers) {
  if (column === 'expiry_date') {
    if (!value || value === '') return 'Expiry date is required for permits';
    // Auto-flag expired
    var expiryDate = new Date(value);
    if (expiryDate < new Date()) {
      setIfExists(sheet, row, headers, 'is_expired', true);
      setIfExists(sheet, row, headers, 'status', 'Expired');
    }
  }
  return true;
}

// FK validation helpers
function validateForeignKeys(sheetName, column, value) {
  var rules = ERP_CONFIG.FK_RULES[sheetName];
  if (!rules) return true;
  for (var i = 0; i < rules.length; i++) {
    if (rules[i].col === column) {
      return validateFKValue(rules[i].ref, rules[i].refCol, value);
    }
  }
  return true;
}

function validateFKValue(refSheet, refCol, value) {
  if (!value || value === '') return true;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(refSheet);
  if (!sheet) return 'Reference table "' + refSheet + '" not found';
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colIdx = headers.indexOf(refCol);
  if (colIdx < 0) return true;
  
  var values = sheet.getRange(2, colIdx + 1, Math.max(1, sheet.getLastRow() - 1), 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(value)) return true;
  }
  return 'Invalid ' + refCol + ': "' + value + '" not found in ' + refSheet;
}

// Business rule helpers
function hasIncompleteWorkPackages(projectId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Work_Packages');
  if (!sheet || sheet.getLastRow() <= 1) return false;
  var data = sheet.getDataRange().getValues();
  var projIdx = data[0].indexOf('project_id');
  var qtyIdx = data[0].indexOf('quantity');
  var execIdx = data[0].indexOf('executed_qty');
  if (projIdx < 0) return false;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][projIdx]) === String(projectId)) {
      if (qtyIdx >= 0 && execIdx >= 0 && Number(data[i][execIdx]) < Number(data[i][qtyIdx])) return true;
    }
  }
  return false;
}

function hasPayments(projectId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Payment_Applications');
  if (!sheet || sheet.getLastRow() <= 1) return false;
  var data = sheet.getDataRange().getValues();
  var projIdx = data[0].indexOf('project_id');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][projIdx]) === String(projectId)) return true;
  }
  return false;
}

function hasApprovedContract(projectId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Contracts');
  if (!sheet || sheet.getLastRow() <= 1) return true; // Default allow if no contracts sheet
  var data = sheet.getDataRange().getValues();
  var projIdx = data[0].indexOf('project_id');
  var statusIdx = data[0].indexOf('status');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][projIdx]) === String(projectId) &&
        String(data[i][statusIdx]) !== 'Rejected') return true;
  }
  return false;
}

function getProjectContractValue(projectId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Projects');
  if (!sheet || sheet.getLastRow() <= 1) return Infinity;
  var data = sheet.getDataRange().getValues();
  var projIdx = data[0].indexOf('project_id');
  var valIdx = data[0].indexOf('contract_value_aed');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][projIdx]) === String(projectId)) return Number(data[i][valIdx]) || Infinity;
  }
  return Infinity;
}

function getRowValue(sheet, row, headers, colName) {
  var idx = headers.indexOf(colName);
  return idx >= 0 ? sheet.getRange(row, idx + 1).getValue() : null;
}

// ============================================================
// MODULE 6: VISUAL STYLING & FORMATTING
// ============================================================

/**
 * Apply elegant base styling to all managed sheets:
 * - Dark header row with readable white text
 * - Subtle alternating row colors
 * - Proper font family & size
 * - System/private column shading
 */
function applyBaseStyleToAllSheets(ss) {
  for (var sheetName in ERP_CONFIG.SHEETS) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    applyBaseStyle(sheet, sheetName);
    SpreadsheetApp.flush();
    Utilities.sleep(300);
  }
}

function applyBaseStyle(sheet, sheetName) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol <= 0) return;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // Set font for entire sheet
  sheet.getRange(1, 1, lastRow, lastCol).setFontFamily('Inter').setFontSize(10);
  
  // Header row: dark slate with white text
  sheet.getRange(1, 1, 1, lastCol)
    .setBackground(ERP_CONFIG.COLORS.HEADER_BG)
    .setFontColor(ERP_CONFIG.COLORS.HEADER_FG)
    .setFontWeight('bold')
    .setFontSize(11)
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 36);
  
  // Alternating row colors (data area)
  if (lastRow > 1) {
    for (var r = 2; r <= lastRow; r++) {
      var bg = (r % 2 === 0) ? ERP_CONFIG.COLORS.ALT_ROW : ERP_CONFIG.COLORS.WHITE;
      sheet.getRange(r, 1, 1, lastCol).setBackground(bg);
    }
  }
  
  // Shade system columns (IDs, timestamps, version) light gray
  var config = ERP_CONFIG.SHEETS[sheetName];
  for (var c = 0; c < headers.length; c++) {
    var h = headers[c];
    if (h === config.pk || ERP_CONFIG.PROTECTED_COLS.indexOf(h) >= 0 || h === 'is_active' || h === 'tags') {
      sheet.getRange(2, c + 1, Math.max(1, lastRow - 1), 1).setBackground(ERP_CONFIG.COLORS.SYSTEM_COL_BG).setFontColor('#718096');
    }
    if (ERP_CONFIG.PRIVATE_COLS.indexOf(h) >= 0) {
      sheet.getRange(2, c + 1, Math.max(1, lastRow - 1), 1).setBackground(ERP_CONFIG.COLORS.PRIVATE_COL_BG);
    }
  }
}

/**
 * Apply colorblind-safe status cell formatting:
 * - Only the STATUS cell gets colored (not the whole row)
 * - High contrast bg/fg pairs from COLORS.STATUS
 */
function applyAllStatusFormatting(ss) {
  for (var sheetName in ERP_CONFIG.SHEETS) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    applyStatusCellFormatting(sheet, sheetName);
    SpreadsheetApp.flush();
    Utilities.sleep(200);
  }
}

function applyStatusCellFormatting(sheet, sheetName) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getDataRange().getValues();
  
  // Find all columns that have status/category coloring
  var colorCols = [];
  var config = ERP_CONFIG.SHEETS[sheetName];
  if (config && config.statusCol) colorCols.push(config.statusCol);
  // Also color health_status, severity, result, is_expired
  ['health_status','severity','result'].forEach(function(col) {
    if (headers.indexOf(col) >= 0 && colorCols.indexOf(col) < 0) colorCols.push(col);
  });
  
  for (var ci = 0; ci < colorCols.length; ci++) {
    var colIdx = headers.indexOf(colorCols[ci]);
    if (colIdx < 0) continue;
    for (var r = 1; r < data.length; r++) {
      var val = String(data[r][colIdx]).trim();
      var style = ERP_CONFIG.COLORS.STATUS[val];
      if (style) {
        sheet.getRange(r + 1, colIdx + 1)
          .setBackground(style.bg).setFontColor(style.fg)
          .setFontWeight('bold').setHorizontalAlignment('center');
      }
    }
  }
  
  // Clear existing conditional format rules (we use direct cell styling now)
  sheet.clearConditionalFormatRules();
}

/**
 * Apply dropdown data validation to all categorical columns
 */
function applyAllDropdowns(ss) {
  for (var sheetName in ERP_CONFIG.DROPDOWNS) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var lastRow = sheet.getLastRow();
    var dropdowns = ERP_CONFIG.DROPDOWNS[sheetName];
    
    for (var colName in dropdowns) {
      var colIdx = headers.indexOf(colName);
      if (colIdx < 0) continue;
      var values = dropdowns[colName];
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(values, true)
        .setAllowInvalid(true) // Allow existing data that may not match
        .setHelpText('Select from: ' + values.join(', '))
        .build();
      sheet.getRange(2, colIdx + 1, Math.max(1, lastRow - 1), 1).setDataValidation(rule);
    }
    SpreadsheetApp.flush();
    Utilities.sleep(200);
  }
}

/**
 * Handle missing/empty data: highlight required empty cells, fill optional with "—"
 */
function handleAllMissingData(ss) {
  // Required fields per table (if empty, highlight yellow)
  var REQUIRED = {
    'Clients': ['client_name','client_type','email','status'],
    'Employees': ['full_name','department','designation','status'],
    'Contractors': ['company_name','specialty','status'],
    'Suppliers': ['company_name','category','status'],
    'Equipment': ['equipment_type','model','status'],
    'Projects': ['project_name','project_type','client_id','status'],
    'Contracts': ['project_id','contract_type','contract_value_aed','status'],
    'Inspections': ['project_id','inspection_type','result'],
    'Safety_Incidents': ['project_id','incident_type','severity'],
    'Payment_Applications': ['project_id','ipc_number','status'],
    'Permits_Approvals': ['project_id','permit_type','status']
  };
  
  // Optional text fields: if empty, fill with "—"
  var OPTIONAL_TEXT = ['internal_notes','issues','description','approved_by','approved_date',
                       'payment_date','file_path'];
  
  for (var sheetName in ERP_CONFIG.SHEETS) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var data = sheet.getDataRange().getValues();
    var lastRow = data.length;
    
    // Highlight required empty fields
    var reqFields = REQUIRED[sheetName] || [];
    for (var ri = 0; ri < reqFields.length; ri++) {
      var ci = headers.indexOf(reqFields[ri]);
      if (ci < 0) continue;
      for (var r = 1; r < lastRow; r++) {
        var v = data[r][ci];
        if (v === '' || v === null || v === undefined) {
          sheet.getRange(r + 1, ci + 1).setBackground(ERP_CONFIG.COLORS.MISSING_BG).setFontColor('#856404');
        }
      }
    }
    
    // Fill optional empty text fields with "—"
    for (var oi = 0; oi < OPTIONAL_TEXT.length; oi++) {
      var optIdx = headers.indexOf(OPTIONAL_TEXT[oi]);
      if (optIdx < 0) continue;
      for (var row = 1; row < lastRow; row++) {
        var val = data[row][optIdx];
        if (val === '' || val === null || val === undefined) {
          sheet.getRange(row + 1, optIdx + 1).setValue('—').setFontColor('#a0aec0').setFontStyle('italic');
        }
      }
    }
    SpreadsheetApp.flush();
    Utilities.sleep(300);
  }
}

/**
 * Auto-resize all columns across all managed sheets
 */
function autoResizeAllColumns(ss) {
  for (var sheetName in ERP_CONFIG.SHEETS) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    var lastCol = sheet.getLastColumn();
    // Auto-resize up to 25 columns to avoid timeout on wide sheets
    var colsToResize = Math.min(lastCol, 25);
    try {
      sheet.autoResizeColumns(1, colsToResize);
      // Cap max width to 250px for readability
      for (var c = 1; c <= colsToResize; c++) {
        if (sheet.getColumnWidth(c) > 250) sheet.setColumnWidth(c, 250);
      }
    } catch(e) { Logger.log('Resize error for ' + sheetName + ': ' + e.message); }
    SpreadsheetApp.flush();
    Utilities.sleep(200);
  }
}

// ============================================================
// MODULE 7: DERIVED INTELLIGENCE ENGINE
// ============================================================

function recalculateAllDerived(ss) {
  recalculateProjectMetrics(ss);
  recalculateClientMetrics(ss);
}

function recalculateProjectMetrics(ss) {
  var sheet = ss.getSheetByName('Projects');
  if (!sheet || sheet.getLastRow() <= 1) return;
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var today = new Date();
  
  var h = {};
  headers.forEach(function(name, idx) { h[name] = idx; });
  
  for (var i = 1; i < data.length; i++) {
    var startDate = new Date(data[i][h['start_date']]);
    var endDate = new Date(data[i][h['end_date']]);
    var budget = Number(data[i][h['contract_value_aed']]) || 0;
    var completion = Number(data[i][h['completion_pct']]) || 0;
    var status = data[i][h['status']];
    
    var durationDays = Math.round((endDate - startDate) / 86400000);
    var elapsedDays = startDate <= today ? Math.max(0, Math.round((Math.min(today, endDate) - startDate) / 86400000)) : 0;
    var remainingDays = endDate > today ? Math.max(0, Math.round((endDate - today) / 86400000)) : 0;
    
    if (h['duration_days'] !== undefined) sheet.getRange(i + 1, h['duration_days'] + 1).setValue(durationDays);
    if (h['elapsed_days'] !== undefined) sheet.getRange(i + 1, h['elapsed_days'] + 1).setValue(elapsedDays);
    if (h['remaining_days'] !== undefined) sheet.getRange(i + 1, h['remaining_days'] + 1).setValue(remainingDays);
  }
}

function recalculateClientMetrics(ss) {
  var clientSheet = ss.getSheetByName('Clients');
  var projSheet = ss.getSheetByName('Projects');
  if (!clientSheet || !projSheet || clientSheet.getLastRow() <= 1 || projSheet.getLastRow() <= 1) return;
  
  var clientData = clientSheet.getDataRange().getValues();
  var projData = projSheet.getDataRange().getValues();
  var cHeaders = clientData[0];
  var pHeaders = projData[0];
  
  var ch = {}, ph = {};
  cHeaders.forEach(function(name, idx) { ch[name] = idx; });
  pHeaders.forEach(function(name, idx) { ph[name] = idx; });
  
  for (var i = 1; i < clientData.length; i++) {
    var clientId = clientData[i][ch['client_id']];
    var totalRevenue = 0, activeCount = 0, riskCount = 0;
    
    for (var j = 1; j < projData.length; j++) {
      if (String(projData[j][ph['client_id']]) === String(clientId)) {
        totalRevenue += Number(projData[j][ph['total_revenue']]) || 0;
        if (['In Progress', 'Delayed', 'On Hold'].indexOf(projData[j][ph['status']]) >= 0) activeCount++;
        if (projData[j][ph['health_status']] === 'Red') riskCount++;
      }
    }
    
    if (ch['lifetime_revenue'] !== undefined) clientSheet.getRange(i + 1, ch['lifetime_revenue'] + 1).setValue(totalRevenue);
    if (ch['active_projects'] !== undefined) clientSheet.getRange(i + 1, ch['active_projects'] + 1).setValue(activeCount);
    if (ch['risk_projects'] !== undefined) clientSheet.getRange(i + 1, ch['risk_projects'] + 1).setValue(riskCount);
  }
}

function recalculateDerivedForRow(sheet, sheetName, row, headers, editedColumn) {
  // Only recalculate if relevant columns were edited
  if (sheetName === 'Projects') {
    var relevantCols = ['start_date', 'end_date', 'contract_value_aed', 'completion_pct', 'status'];
    if (relevantCols.indexOf(editedColumn) >= 0) {
      var today = new Date();
      var h = {};
      headers.forEach(function(name, idx) { h[name] = idx; });
      
      var startDate = new Date(sheet.getRange(row, h['start_date'] + 1).getValue());
      var endDate = new Date(sheet.getRange(row, h['end_date'] + 1).getValue());
      
      if (h['duration_days'] !== undefined) 
        sheet.getRange(row, h['duration_days'] + 1).setValue(Math.round((endDate - startDate) / 86400000));
      if (h['elapsed_days'] !== undefined)
        sheet.getRange(row, h['elapsed_days'] + 1).setValue(startDate <= today ? Math.max(0, Math.round((Math.min(today, endDate) - startDate) / 86400000)) : 0);
      if (h['remaining_days'] !== undefined)
        sheet.getRange(row, h['remaining_days'] + 1).setValue(endDate > today ? Math.max(0, Math.round((endDate - today) / 86400000)) : 0);
    }
  }
}

// ============================================================
// MODULE 8: SMART TAGGING ENGINE
// ============================================================

function updateTagsForRow(sheet, sheetName, row, headers) {
  var tagsIdx = headers.indexOf('tags');
  if (tagsIdx < 0) return;
  
  var h = {};
  headers.forEach(function(name, idx) { h[name] = idx; });
  
  var tags = [];
  var getVal = function(col) {
    return h[col] !== undefined ? sheet.getRange(row, h[col] + 1).getValue() : null;
  };
  
  switch (sheetName) {
    case 'Projects':
      if (Number(getVal('contract_value_aed')) >= 100000000) tags.push('High Budget');
      if (getVal('status') === 'Delayed') tags.push('Delayed');
      if (Number(getVal('risk_score')) >= 7) tags.push('High Risk');
      if (Number(getVal('duration_days')) >= 1080) tags.push('Long Duration');
      break;
    
    case 'Clients':
      if (Number(getVal('credit_limit_aed')) >= 5000000) tags.push('VIP Client');
      if (getVal('status') === 'Inactive') tags.push('Inactive Client');
      if (getVal('status') === 'Blacklisted') tags.push('Blacklisted');
      if (Number(getVal('risk_projects')) > 0) tags.push('Has Risk Projects');
      break;
    
    case 'Payment_Applications':
      if (getVal('status') === 'Overdue') tags.push('Overdue Payment');
      break;
    
    case 'Safety_Incidents':
      if (getVal('severity') === 'High') tags.push('High Severity');
      break;
    
    case 'Permits_Approvals':
      if (getVal('is_expired') === true || getVal('is_expired') === 'TRUE') tags.push('Expired Permit');
      break;
    
    case 'Project_Phases':
      var phaseName = getVal('phase_name');
      if (phaseName === 'Superstructure' || phaseName === 'Piling') tags.push('Critical Phase');
      break;
  }
  
  sheet.getRange(row, tagsIdx + 1).setValue(tags.join(','));
}

// ============================================================
// MODULE 9: PRIVACY & SECURITY
// ============================================================

function applyPrivacyProtections(ss) {
  var sheets = ss.getSheets();
  
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var sheetName = sheet.getName();
    if (!ERP_CONFIG.SHEETS[sheetName]) continue;
    if (sheet.getLastRow() <= 1) continue;
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Protect private columns
    for (var c = 0; c < headers.length; c++) {
      if (ERP_CONFIG.PRIVATE_COLS.indexOf(headers[c]) >= 0) {
        // Gray shading for private columns
        var range = sheet.getRange(1, c + 1, sheet.getLastRow(), 1);
        range.setBackground('#f0f0f0');
        
        // Protection
        try {
          var protection = sheet.protect();
          protection.setDescription('ERP Private Column: ' + headers[c]);
          var unprotected = [];
          for (var r = 0; r < headers.length; r++) {
            if (ERP_CONFIG.PRIVATE_COLS.indexOf(headers[r]) < 0 && 
                ERP_CONFIG.PROTECTED_COLS.indexOf(headers[r]) < 0 &&
                headers[r] !== ERP_CONFIG.SHEETS[sheetName].pk) {
              // This column is editable
            }
          }
        } catch(e) { /* Protection may fail for non-owner */ }
      }
      
      // Protect system columns (PK, created_at, record_version)
      if (headers[c] === ERP_CONFIG.SHEETS[sheetName].pk || 
          ERP_CONFIG.PROTECTED_COLS.indexOf(headers[c]) >= 0) {
        sheet.getRange(1, c + 1, sheet.getLastRow(), 1).setBackground('#e8e8e8');
      }
    }
    
    // Protect ID and system columns via sheet-level protection
    try {
      var existingProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      existingProtections.forEach(function(p) { p.remove(); });
      
      var sheetProtection = sheet.protect().setDescription('ERP Protection: ' + sheetName);
      // Allow admin to edit everything
      sheetProtection.addEditor(ERP_CONFIG.ADMIN_EMAIL);
      
      // Set unprotected ranges (everything except protected columns)
      var unprotectedRanges = [];
      for (var col = 0; col < headers.length; col++) {
        if (ERP_CONFIG.PROTECTED_COLS.indexOf(headers[col]) < 0 &&
            headers[col] !== ERP_CONFIG.SHEETS[sheetName].pk) {
          unprotectedRanges.push(sheet.getRange(2, col + 1, Math.max(1, sheet.getLastRow() - 1), 1));
        }
      }
      if (unprotectedRanges.length > 0) {
        sheetProtection.setUnprotectedRanges(unprotectedRanges);
      }
    } catch(e) { Logger.log('Protection error for ' + sheetName + ': ' + e.message); }
  }
}

// ============================================================
// MODULE 10: AUDIT & GOVERNANCE
// ============================================================

function logAudit(user, action, table, rowId, field, oldValue, newValue) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ERP_CONFIG.AUDIT_SHEET);
    if (!sheet) {
      createAuditLogSheet(ss);
      sheet = ss.getSheetByName(ERP_CONFIG.AUDIT_SHEET);
    }
    
    sheet.appendRow([
      new Date().toISOString(),
      user || ERP_CONFIG.SYSTEM_USER,
      action,
      table,
      String(rowId),
      field || '',
      String(oldValue || ''),
      String(newValue || '')
    ]);
  } catch(e) {
    Logger.log('Audit log error: ' + e.message);
  }
}

// ============================================================
// SCHEDULED: AUTO-DELAYED DETECTION
// ============================================================

function autoDelayedDetection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Projects');
  if (!sheet || sheet.getLastRow() <= 1) return;
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var h = {};
  headers.forEach(function(name, idx) { h[name] = idx; });
  
  var today = new Date();
  var changed = 0;
  
  for (var i = 1; i < data.length; i++) {
    var endDate = new Date(data[i][h['end_date']]);
    var status = data[i][h['status']];
    
    if (today > endDate && status !== 'Completed' && status !== 'Cancelled' && status !== 'Delayed') {
      sheet.getRange(i + 1, h['status'] + 1).setValue('Delayed');
      logAudit('SYSTEM', 'AUTO_STATUS_CHANGE', 'Projects', data[i][h['project_id']], 'status', status, 'Delayed');
      changed++;
    }
  }
  
  if (changed > 0) {
    Logger.log('Auto-Delayed Detection: ' + changed + ' projects marked as Delayed');
  }
}

// ============================================================
// SCHEDULED: AUTO PERMIT EXPIRY CHECK
// ============================================================

function autoPermitExpiryCheck() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Permits_Approvals');
  if (!sheet || sheet.getLastRow() <= 1) return;
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var h = {};
  headers.forEach(function(name, idx) { h[name] = idx; });
  
  var today = new Date();
  var flagged = 0;
  
  for (var i = 1; i < data.length; i++) {
    var expiryDate = new Date(data[i][h['expiry_date']]);
    var currentExpired = data[i][h['is_expired']];
    
    if (today > expiryDate && currentExpired !== true && currentExpired !== 'TRUE') {
      sheet.getRange(i + 1, h['is_expired'] + 1).setValue(true);
      sheet.getRange(i + 1, h['status'] + 1).setValue('Expired');
      logAudit('SYSTEM', 'AUTO_EXPIRY', 'Permits_Approvals', data[i][h['permit_id']], 'is_expired', 'FALSE', 'TRUE');
      flagged++;
    }
  }
  
  if (flagged > 0) {
    Logger.log('Permit Expiry Check: ' + flagged + ' permits flagged as expired');
  }
}

// ============================================================
// MENU SETUP
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏗️ Dubai ERP System')
    .addItem('⚙️ Setup ERP System', 'setupERPSystem')
    .addSeparator()
    .addItem('📥 Import All Data', 'importAllData')
    .addItem('🔍 Validate Foreign Keys', 'runFKValidation')
    .addSeparator()
    .addItem('🔄 Recalculate All Metrics', 'recalculateAllFromMenu')
    .addItem('🏷️ Refresh All Tags', 'refreshAllTags')
    .addItem('🎨 Reapply Formatting & Dropdowns', 'applyFormattingFromMenu')
    .addSeparator()
    .addItem('🗑️ Soft Delete Selected Row', 'softDeleteRow')
    .addItem('👁️ Show Audit Log', 'showAuditLog')
    .addSeparator()
    .addItem('🔒 Reapply Privacy Protections', 'reapplyPrivacy')
    .addToUi();
}

function recalculateAllFromMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Recalculating all derived metrics...', '🔄 Recalculating', -1);
  recalculateAllDerived(ss);
  ss.toast('✅ All metrics recalculated!', 'Done', 5);
}

function refreshAllTags() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Refreshing all tags...', '🏷️ Tags', -1);
  
  for (var sheetName in ERP_CONFIG.SHEETS) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) continue;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    for (var i = 2; i <= sheet.getLastRow(); i++) {
      updateTagsForRow(sheet, sheetName, i, headers);
    }
  }
  
  ss.toast('✅ All tags refreshed!', 'Done', 5);
}

function applyFormattingFromMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Applying styling, dropdowns & formatting...', '🎨 Formatting', -1);
  applyBaseStyleToAllSheets(ss);
  applyAllDropdowns(ss);
  handleAllMissingData(ss);
  applyAllStatusFormatting(ss);
  autoResizeAllColumns(ss);
  ss.toast('✅ All formatting & dropdowns applied!', 'Done', 5);
}

function showAuditLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ERP_CONFIG.AUDIT_SHEET);
  if (sheet) {
    sheet.showSheet();
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('Audit Log not found. Run Setup ERP System first.');
  }
}

function reapplyPrivacy() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Reapplying privacy protections...', '🔒 Privacy', -1);
  applyPrivacyProtections(ss);
  ss.toast('✅ Privacy protections applied!', 'Done', 5);
}
