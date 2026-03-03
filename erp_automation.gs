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

  // Status color maps
  STATUS_COLORS: {
    'Projects': {
      'Completed': '#27ae60', 'In Progress': '#2980b9', 'On Hold': '#f39c12',
      'Delayed': '#e74c3c', 'Cancelled': '#95a5a6', 'Awarded': '#8e44ad'
    },
    'Clients': {
      'Active': '#27ae60', 'Inactive': '#95a5a6', 'Not in Contact': '#f1c40f', 'Blacklisted': '#e74c3c'
    },
    'Payments': {
      'Paid': '#27ae60', 'Certified': '#2ecc71', 'Submitted': '#f1c40f',
      'Under Review': '#f39c12', 'Rejected': '#e74c3c', 'Overdue': '#c0392b'
    },
    'Safety': {
      'High': '#e74c3c', 'Medium': '#f39c12', 'Low': '#f1c40f'
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
  ss.toast('Audit Log created', '⚙️ Setup', -1);
  
  // 2. Install triggers
  installTriggers(ss);
  ss.toast('Triggers installed', '⚙️ Setup', -1);
  
  // 3. Apply conditional formatting to all sheets
  applyAllConditionalFormatting(ss);
  ss.toast('Conditional formatting applied', '⚙️ Setup', -1);
  
  // 4. Apply privacy protections
  applyPrivacyProtections(ss);
  ss.toast('Privacy protections applied', '⚙️ Setup', -1);
  
  // 5. Run derived calculations
  recalculateAllDerived(ss);
  ss.toast('Derived calculations complete', '⚙️ Setup', -1);
  
  ss.toast('✅ ERP System Setup Complete!', '⚙️ Done', 10);
  ui.alert('ERP System Setup Complete!\n\nAll automation is now active:\n• Edit triggers (timestamps, versioning, audit)\n• Data validation rules\n• Conditional formatting\n• Privacy protections\n• Derived calculations');
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
// MODULE 6: CONDITIONAL FORMATTING
// ============================================================

function applyAllConditionalFormatting(ss) {
  applyProjectFormatting(ss);
  applyClientFormatting(ss);
  applyPaymentFormatting(ss);
  applySafetyFormatting(ss);
  applyPermitFormatting(ss);
}

function applyProjectFormatting(ss) {
  var sheet = ss.getSheetByName('Projects');
  if (!sheet || sheet.getLastRow() <= 1) return;
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var statusIdx = headers.indexOf('status');
  var healthIdx = headers.indexOf('health_status');
  var bvIdx = headers.indexOf('budget_variance_pct');
  if (statusIdx < 0) return;
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var statusCol = String.fromCharCode(65 + statusIdx);
  
  // Clear existing rules
  sheet.clearConditionalFormatRules();
  var rules = [];
  
  // Status-based row coloring
  var statusColors = ERP_CONFIG.STATUS_COLORS['Projects'];
  for (var status in statusColors) {
    var range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + statusCol + '2="' + status + '"')
      .setBackground(statusColors[status])
      .setFontColor(status === 'Cancelled' ? '#2c3e50' : '#ffffff')
      .setRanges([range])
      .build();
    rules.push(rule);
  }
  
  // Budget variance coloring (if column exists)
  if (bvIdx >= 0) {
    var bvCol = String.fromCharCode(65 + bvIdx);
    var bvRange = sheet.getRange(2, bvIdx + 1, lastRow - 1, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(10).setBackground('#e74c3c').setFontColor('#ffffff').setRanges([bvRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(5, 10).setBackground('#f39c12').setFontColor('#ffffff').setRanges([bvRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0, 5).setBackground('#f1c40f').setRanges([bvRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground('#27ae60').setFontColor('#ffffff').setRanges([bvRange]).build());
  }
  
  // Health status coloring
  if (healthIdx >= 0) {
    var healthRange = sheet.getRange(2, healthIdx + 1, lastRow - 1, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Green').setBackground('#27ae60').setFontColor('#ffffff').setRanges([healthRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Amber').setBackground('#f39c12').setFontColor('#ffffff').setRanges([healthRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Red').setBackground('#e74c3c').setFontColor('#ffffff').setRanges([healthRange]).build());
  }
  
  sheet.setConditionalFormatRules(rules);
}

function applyClientFormatting(ss) {
  var sheet = ss.getSheetByName('Clients');
  if (!sheet || sheet.getLastRow() <= 1) return;
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var statusIdx = headers.indexOf('status');
  if (statusIdx < 0) return;
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var statusCol = String.fromCharCode(65 + statusIdx);
  
  sheet.clearConditionalFormatRules();
  var rules = [];
  
  var statusColors = ERP_CONFIG.STATUS_COLORS['Clients'];
  for (var status in statusColors) {
    var range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + statusCol + '2="' + status + '"')
      .setBackground(statusColors[status])
      .setFontColor(status === 'Blacklisted' ? '#ffffff' : (status === 'Inactive' ? '#2c3e50' : '#ffffff'))
      .setRanges([range])
      .build());
  }
  
  sheet.setConditionalFormatRules(rules);
}

function applyPaymentFormatting(ss) {
  var sheet = ss.getSheetByName('Payment_Applications');
  if (!sheet || sheet.getLastRow() <= 1) return;
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var statusIdx = headers.indexOf('status');
  if (statusIdx < 0) return;
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var statusCol = String.fromCharCode(65 + statusIdx);
  
  sheet.clearConditionalFormatRules();
  var rules = [];
  
  var colors = ERP_CONFIG.STATUS_COLORS['Payments'];
  for (var status in colors) {
    var range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + statusCol + '2="' + status + '"')
      .setBackground(colors[status])
      .setFontColor('#ffffff')
      .setRanges([range])
      .build());
  }
  
  sheet.setConditionalFormatRules(rules);
}

function applySafetyFormatting(ss) {
  var sheet = ss.getSheetByName('Safety_Incidents');
  if (!sheet || sheet.getLastRow() <= 1) return;
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var sevIdx = headers.indexOf('severity');
  if (sevIdx < 0) return;
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var sevCol = String.fromCharCode(65 + sevIdx);
  
  sheet.clearConditionalFormatRules();
  var rules = [];
  
  var colors = ERP_CONFIG.STATUS_COLORS['Safety'];
  for (var sev in colors) {
    var range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + sevCol + '2="' + sev + '"')
      .setBackground(colors[sev])
      .setFontColor(sev === 'Low' ? '#2c3e50' : '#ffffff')
      .setRanges([range])
      .build());
  }
  
  sheet.setConditionalFormatRules(rules);
}

function applyPermitFormatting(ss) {
  var sheet = ss.getSheetByName('Permits_Approvals');
  if (!sheet || sheet.getLastRow() <= 1) return;
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var expiredIdx = headers.indexOf('is_expired');
  if (expiredIdx < 0) return;
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var expCol = String.fromCharCode(65 + expiredIdx);
  
  sheet.clearConditionalFormatRules();
  var rules = [];
  
  var range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + expCol + '2=TRUE')
    .setBackground('#e74c3c').setFontColor('#ffffff')
    .setRanges([range]).build());
  
  sheet.setConditionalFormatRules(rules);
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
    .addItem('📊 Apply Conditional Formatting', 'applyFormattingFromMenu')
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
  ss.toast('Applying conditional formatting...', '📊 Formatting', -1);
  applyAllConditionalFormatting(ss);
  ss.toast('✅ Formatting applied!', 'Done', 5);
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
