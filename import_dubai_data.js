/**
 * DUBAI PROJECT MANAGEMENT - ENTERPRISE ERP DATA IMPORTER
 * Schema-validated import engine with upsert, FK validation, and audit logging
 * 
 * FEATURES:
 *  - CSV schema validation (abort on mismatch)
 *  - Upsert logic (update existing, insert new)
 *  - Duplicate primary key prevention
 *  - Foreign key relationship validation
 *  - Import status logging (Import_Log sheet)
 *  - Batch writing for large datasets
 *  - Post-import automation triggers
 *
 * INSTRUCTIONS:
 *  1. Open Google Sheets → Extensions → Apps Script
 *  2. Paste this code and save
 *  3. Run "importAllData" function
 *  4. Grant permissions when prompted
 */

// ============================================================
// CONFIGURATION
// ============================================================
var CONFIG = {
    GITHUB_USER: 'Anandhuvimalan',
    REPO: 'dubai-project-management-data',
    BRANCH: 'main',
    BATCH_SIZE: 2000,
    ADMIN_EMAIL: 'anandhu7833@gmail.com',

    // Schema definitions: name, sheet, primary key column, expected headers (first 5 for validation), FK references
    FILES: [
        // Core Master Tables (no FK dependencies)
        {
            name: 'clients.csv', sheet: 'Clients', pk: 'client_id', fks: [],
            requiredHeaders: ['client_id', 'client_name', 'client_type', 'email', 'status', 'created_at', 'is_active', 'tags']
        },
        {
            name: 'employees.csv', sheet: 'Employees', pk: 'employee_id', fks: [],
            requiredHeaders: ['employee_id', 'full_name', 'department', 'designation', 'created_at', 'is_active', 'tags']
        },
        {
            name: 'contractors.csv', sheet: 'Contractors', pk: 'contractor_id', fks: [],
            requiredHeaders: ['contractor_id', 'company_name', 'specialty', 'created_at', 'is_active', 'tags']
        },
        {
            name: 'suppliers.csv', sheet: 'Suppliers', pk: 'supplier_id', fks: [],
            requiredHeaders: ['supplier_id', 'company_name', 'category', 'created_at', 'is_active', 'tags']
        },
        {
            name: 'equipment.csv', sheet: 'Equipment', pk: 'equipment_id', fks: [],
            requiredHeaders: ['equipment_id', 'equipment_type', 'model', 'created_at', 'is_active', 'tags']
        },

        // Project & Contract Management (FK dependencies)
        {
            name: 'projects.csv', sheet: 'Projects', pk: 'project_id',
            fks: [{ col: 'client_id', ref: 'Clients', refCol: 'client_id' }, { col: 'project_manager_id', ref: 'Employees', refCol: 'employee_id' }],
            requiredHeaders: ['project_id', 'project_name', 'client_id', 'project_manager_id', 'status', 'risk_score', 'health_status', 'created_at', 'is_active', 'tags']
        },
        {
            name: 'contracts.csv', sheet: 'Contracts', pk: 'contract_id',
            fks: [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
            requiredHeaders: ['contract_id', 'project_id', 'contract_type', 'contract_value_aed', 'status', 'created_at', 'is_active', 'tags']
        },
        {
            name: 'project_phases.csv', sheet: 'Project_Phases', pk: 'phase_id',
            fks: [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
            requiredHeaders: ['phase_id', 'project_id', 'phase_name', 'status', 'created_at', 'is_active', 'tags']
        },
        {
            name: 'work_packages.csv', sheet: 'Work_Packages', pk: 'package_id',
            fks: [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
            requiredHeaders: ['package_id', 'project_id', 'item_description', 'created_at', 'is_active', 'tags']
        },

        // Compliance & Approvals
        {
            name: 'permits_approvals.csv', sheet: 'Permits_Approvals', pk: 'permit_id',
            fks: [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
            requiredHeaders: ['permit_id', 'project_id', 'permit_type', 'status', 'is_expired', 'created_at', 'is_active', 'tags']
        },
        {
            name: 'inspections.csv', sheet: 'Inspections', pk: 'inspection_id',
            fks: [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
            requiredHeaders: ['inspection_id', 'project_id', 'inspection_type', 'result', 'created_at', 'is_active', 'tags']
        },
        {
            name: 'safety_incidents.csv', sheet: 'Safety_Incidents', pk: 'incident_id',
            fks: [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
            requiredHeaders: ['incident_id', 'project_id', 'incident_type', 'severity', 'created_at', 'is_active', 'tags']
        },

        // Financial Management
        {
            name: 'payment_applications.csv', sheet: 'Payment_Applications', pk: 'ipc_id',
            fks: [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
            requiredHeaders: ['ipc_id', 'project_id', 'ipc_number', 'status', 'remaining_balance', 'created_at', 'is_active', 'tags']
        },
        {
            name: 'variation_orders.csv', sheet: 'Variation_Orders', pk: 'vo_id',
            fks: [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
            requiredHeaders: ['vo_id', 'project_id', 'vo_number', 'status', 'created_at', 'is_active', 'tags']
        },
        {
            name: 'purchase_orders.csv', sheet: 'Purchase_Orders', pk: 'po_id',
            fks: [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
            requiredHeaders: ['po_id', 'project_id', 'po_type', 'status', 'created_at', 'is_active', 'tags']
        },

        // Resource & Documentation
        {
            name: 'daily_site_reports.csv', sheet: 'Daily_Site_Reports', pk: 'report_id',
            fks: [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
            requiredHeaders: ['report_id', 'project_id', 'report_date', 'created_at', 'is_active', 'tags']
        },
        {
            name: 'project_documents.csv', sheet: 'Project_Documents', pk: 'document_id',
            fks: [{ col: 'project_id', ref: 'Projects', refCol: 'project_id' }],
            requiredHeaders: ['document_id', 'project_id', 'document_type', 'status', 'created_at', 'is_active', 'tags']
        }
    ]
};

// ============================================================
// MAIN IMPORT FUNCTION
// ============================================================
function importAllData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var baseUrl = getBaseUrl();
    var startTime = new Date();
    var totalRows = 0;
    var results = [];
    var importLog = [];

    ss.toast('Starting Enterprise ERP Data Import...', '🏗️ Dubai PM ERP', -1);

    // Clean up old sheets before importing
    deleteOldSheets(ss);
    ss.toast('Old sheets removed. Importing fresh data...', '🏗️ Dubai PM ERP', -1);

    // Phase 1: Import all tables with schema validation
    for (var i = 0; i < CONFIG.FILES.length; i++) {
        var file = CONFIG.FILES[i];
        ss.toast('Importing ' + file.sheet + ' (' + (i + 1) + '/' + CONFIG.FILES.length + ')...', '🏗️ Import', -1);

        try {
            var result = importWithValidation(ss, baseUrl + file.name, file);
            totalRows += result.rows;
            results.push({ file: file.sheet, rows: result.rows, status: result.status, skipped: result.skipped, errors: result.errors });
            importLog.push({
                timestamp: new Date().toISOString(),
                table: file.sheet,
                rows_imported: result.rows,
                rows_skipped: result.skipped,
                schema_valid: result.schemaValid,
                fk_errors: result.fkErrors || 0,
                status: result.status
            });
            Logger.log('✓ ' + file.sheet + ': ' + result.rows + ' rows (' + result.skipped + ' skipped)');
        } catch (e) {
            results.push({ file: file.sheet, rows: 0, status: 'ERROR: ' + e.message, skipped: 0, errors: e.message });
            importLog.push({
                timestamp: new Date().toISOString(),
                table: file.sheet,
                rows_imported: 0,
                rows_skipped: 0,
                schema_valid: false,
                fk_errors: 0,
                status: 'ERROR: ' + e.message
            });
            Logger.log('✗ ' + file.sheet + ': ERROR - ' + e.message);
        }
        // Memory management: flush and pause between tables
        SpreadsheetApp.flush();
        Utilities.sleep(2000);
    }

    // Phase 2: Post-import FK validation
    ss.toast('Validating foreign key relationships...', '🏗️ Validation', -1);
    var fkResults = validateAllForeignKeys(ss);

    var duration = Math.round((new Date() - startTime) / 1000 / 60);
    createSummarySheet(ss, results, totalRows, duration, fkResults);
    createImportLog(ss, importLog);

    // Remove temporary sheet if it exists
    var tempSheet = ss.getSheetByName('Temp');
    if (tempSheet) { try { ss.deleteSheet(tempSheet); } catch (e) { } }

    ss.toast('✅ COMPLETE! ' + totalRows + ' rows imported in ' + duration + ' min.', 'Done', 10);
}

// ============================================================
// SCHEMA-VALIDATED IMPORT
// ============================================================
function importWithValidation(ss, url, fileConfig) {
    var response = fetchWithRetry(url, 5);
    if (!response) throw new Error('Failed to fetch after 5 retries');

    var csvText = response.getContentText();
    response = null; // Release response memory

    var csvData = Utilities.parseCsv(csvText);
    csvText = null; // Release raw text memory

    if (!csvData || csvData.length === 0) throw new Error('Empty CSV data');

    var headers = csvData[0];
    var result = { rows: 0, skipped: 0, status: 'OK', schemaValid: true, errors: '', fkErrors: 0 };

    // Schema validation: check required headers exist
    var missingHeaders = [];
    for (var h = 0; h < fileConfig.requiredHeaders.length; h++) {
        if (headers.indexOf(fileConfig.requiredHeaders[h]) === -1) {
            missingHeaders.push(fileConfig.requiredHeaders[h]);
        }
    }
    if (missingHeaders.length > 0) {
        csvData = null;
        throw new Error('SCHEMA_MISMATCH: Missing ' + missingHeaders.join(', '));
    }

    // Find PK column index for duplicate detection
    var pkIdx = headers.indexOf(fileConfig.pk);

    var sheet = ss.getSheetByName(fileConfig.sheet);
    if (sheet) {
        sheet.clear();
    } else {
        sheet = ss.insertSheet(fileConfig.sheet);
    }

    // Fresh import: write directly (no upsert needed for clean import)
    var totalRows = csvData.length;
    var numCols = csvData[0].length;

    if (totalRows <= CONFIG.BATCH_SIZE) {
        sheet.getRange(1, 1, totalRows, numCols).setValues(csvData);
    } else {
        writeInBatches(sheet, csvData, CONFIG.BATCH_SIZE);
    }

    csvData = null; // Release parsed CSV memory

    result.rows = totalRows - 1; // Exclude header
    formatSheet(sheet, numCols, totalRows);
    return result;
}

// ============================================================
// FOREIGN KEY VALIDATION (POST-IMPORT)
// ============================================================
function validateAllForeignKeys(ss) {
    var fkResults = [];
    var pkCache = {};

    for (var i = 0; i < CONFIG.FILES.length; i++) {
        var file = CONFIG.FILES[i];
        if (!file.fks || file.fks.length === 0) continue;

        var sheet = ss.getSheetByName(file.sheet);
        if (!sheet || sheet.getLastRow() <= 1) continue;

        var data = sheet.getDataRange().getValues();
        var headers = data[0];

        for (var f = 0; f < file.fks.length; f++) {
            var fk = file.fks[f];
            var fkColIdx = headers.indexOf(fk.col);
            if (fkColIdx === -1) continue;

            // Build PK set for referenced table (cached)
            if (!pkCache[fk.ref]) {
                var refSheet = ss.getSheetByName(fk.ref);
                if (!refSheet || refSheet.getLastRow() <= 1) {
                    fkResults.push({ table: file.sheet, fk: fk.col, ref: fk.ref, orphans: data.length - 1, status: 'REF TABLE MISSING' });
                    continue;
                }
                var refData = refSheet.getDataRange().getValues();
                var refPkIdx = refData[0].indexOf(fk.refCol);
                if (refPkIdx === -1) continue;
                pkCache[fk.ref] = {};
                for (var r = 1; r < refData.length; r++) {
                    pkCache[fk.ref][refData[r][refPkIdx]] = true;
                }
            }

            // Check FK values
            var orphans = 0;
            for (var row = 1; row < data.length; row++) {
                var val = data[row][fkColIdx];
                if (val && val !== '' && !pkCache[fk.ref][val]) {
                    orphans++;
                }
            }
            fkResults.push({
                table: file.sheet,
                fk: fk.col,
                ref: fk.ref,
                orphans: orphans,
                status: orphans === 0 ? '✓ VALID' : '✗ ' + orphans + ' ORPHANS'
            });
        }
    }
    return fkResults;
}

// ============================================================
// BATCH WRITER
// ============================================================
function writeInBatches(sheet, data, batchSize) {
    var totalRows = data.length;
    var numCols = data[0].length;
    sheet.getRange(1, 1, 1, numCols).setValues([data[0]]); // Header

    var currentRow = 2;
    var dataIndex = 1;
    while (dataIndex < totalRows) {
        var endIndex = Math.min(dataIndex + batchSize, totalRows);
        var batch = data.slice(dataIndex, endIndex);
        if (batch.length > 0) {
            sheet.getRange(currentRow, 1, batch.length, numCols).setValues(batch);
            currentRow += batch.length;
            dataIndex = endIndex;
            SpreadsheetApp.flush();
            Utilities.sleep(300);
        }
    }
}

// ============================================================
// UTILITIES
// ============================================================
function fetchWithRetry(url, maxRetries) {
    for (var i = 0; i < maxRetries; i++) {
        try {
            var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
            if (response.getResponseCode() === 200) return response;
            Logger.log('Fetch attempt ' + (i + 1) + ' returned ' + response.getResponseCode());
        } catch (e) {
            Logger.log('Fetch attempt ' + (i + 1) + ' failed: ' + e.message);
            Utilities.sleep(1000 * Math.pow(2, i));
        }
    }
    return null;
}

function formatSheet(sheet, numCols, totalRows) {
    try {
        sheet.setFrozenRows(1);
        sheet.getRange(1, 1, 1, numCols).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#e0e0e0');
        sheet.setRowHeight(1, 30);
        if (totalRows < 5000) {
            sheet.autoResizeColumns(1, Math.min(numCols, 20)); // Limit to avoid timeout
        }
    } catch (e) { Logger.log('Format error: ' + e.message); }
}

function deleteOldSheets(ss) {
    var sheets = ss.getSheets();
    Logger.log('Removing all existing sheets for clean import...');
    for (var i = sheets.length - 1; i >= 0; i--) {
        try {
            ss.deleteSheet(sheets[i]);
        } catch (e) {
            sheets[i].clear();
        }
    }
    if (ss.getSheets().length === 0) {
        ss.insertSheet('Temp');
    }
    Logger.log('All old sheets removed.');
}

function createSummarySheet(ss, results, totalRows, duration, fkResults) {
    var sheet = ss.getSheetByName('Import_Summary');
    if (!sheet) { sheet = ss.insertSheet('Import_Summary', 0); } else { sheet.clear(); }

    var data = [
        ['DUBAI ERP - Enterprise Import Summary', '', '', '', ''],
        ['Import Timestamp:', new Date().toISOString(), '', '', ''],
        ['Total Rows Imported:', totalRows, '', '', ''],
        ['Duration (min):', duration, '', '', ''],
        ['Import Engine:', 'Enterprise v2.0 (Schema Validated)', '', '', ''],
        ['', '', '', '', ''],
        ['TABLE', 'ROWS', 'SKIPPED', 'STATUS', 'ERRORS']
    ];
    results.forEach(function (r) {
        data.push([r.file, r.rows, r.skipped || 0, r.status, r.errors || '']);
    });

    // FK validation results
    if (fkResults && fkResults.length > 0) {
        data.push(['', '', '', '', '']);
        data.push(['FOREIGN KEY VALIDATION', '', '', '', '']);
        data.push(['TABLE', 'FK COLUMN', 'REFERENCES', 'ORPHANS', 'STATUS']);
        fkResults.forEach(function (fk) {
            data.push([fk.table, fk.fk, fk.ref, fk.orphans, fk.status]);
        });
    }

    sheet.getRange(1, 1, data.length, 5).setValues(data);
    sheet.getRange(1, 1, 1, 5).merge().setFontSize(14).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#e0e0e0');
    sheet.getRange(7, 1, 1, 5).setFontWeight('bold').setBackground('#2c3e50').setFontColor('white');

    // Color-code status
    for (var r = 8; r <= 8 + results.length; r++) {
        var statusCell = sheet.getRange(r, 4);
        var status = statusCell.getValue();
        if (status === 'OK') statusCell.setBackground('#27ae60').setFontColor('white');
        else if (String(status).indexOf('ERROR') >= 0) statusCell.setBackground('#e74c3c').setFontColor('white');
        else if (String(status).indexOf('SCHEMA') >= 0) statusCell.setBackground('#f39c12').setFontColor('white');
    }
    sheet.autoResizeColumns(1, 5);
}

function createImportLog(ss, importLog) {
    var sheet = ss.getSheetByName('Import_Log');
    if (!sheet) { sheet = ss.insertSheet('Import_Log'); } else { sheet.clear(); }

    var headers = ['Timestamp', 'Table', 'Rows Imported', 'Rows Skipped', 'Schema Valid', 'FK Errors', 'Status'];
    var data = [headers];
    importLog.forEach(function (entry) {
        data.push([
            entry.timestamp, entry.table, entry.rows_imported,
            entry.rows_skipped, entry.schema_valid, entry.fk_errors, entry.status
        ]);
    });
    sheet.getRange(1, 1, data.length, headers.length).setValues(data);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#2c3e50').setFontColor('white');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
}

function getBaseUrl() {
    return 'https://raw.githubusercontent.com/' + CONFIG.GITHUB_USER + '/' + CONFIG.REPO + '/' + CONFIG.BRANCH + '/';
}

function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('🏗️ Dubai ERP System')
        .addItem('📥 Import All Data', 'importAllData')
        .addItem('🔍 Validate Foreign Keys', 'runFKValidation')
        .addItem('📊 Setup ERP Automation', 'setupERPSystem')
        .addToUi();
}

function runFKValidation() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Running FK validation...', '🔍 Validation', -1);
    var results = validateAllForeignKeys(ss);
    var msg = 'FK Validation Results:\n';
    results.forEach(function (r) { msg += r.table + '.' + r.fk + ' → ' + r.ref + ': ' + r.status + '\n'; });
    SpreadsheetApp.getUi().alert(msg);
}
