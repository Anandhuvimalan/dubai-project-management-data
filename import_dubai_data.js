/**
 * DUBAI PROJECT MANAGEMENT DATA IMPORTER
 * Imports generated Dubai PM dataset from GitHub to Google Sheets
 * 
 * INSTRUCTIONS:
 * 1. Open Google Sheets → Extensions → Apps Script
 * 2. Paste this code and save
 * 3. Run "importAllData" function
 * 4. Grant permissions when asked
 */

// ============================================================
// CONFIGURATION
// ============================================================
var CONFIG = {
    GITHUB_USER: 'Anandhuvimalan',
    REPO: 'dubai-project-management-data',
    BRANCH: 'main',
    BATCH_SIZE: 5000,

    // Files ordered by dependency/importance
    FILES: [
        { name: 'clients.csv', sheet: 'Clients', expectedRows: 25 },
        { name: 'employees.csv', sheet: 'Employees', expectedRows: 250 },
        { name: 'vendors.csv', sheet: 'Vendors', expectedRows: 40 },
        { name: 'projects.csv', sheet: 'Projects', expectedRows: 50 },
        { name: 'project_milestones.csv', sheet: 'Milestones', expectedRows: 400 },
        { name: 'tasks.csv', sheet: 'Tasks', expectedRows: 1000 },
        { name: 'assignments.csv', sheet: 'Assignments', expectedRows: 1500 },
        { name: 'timesheets.csv', sheet: 'Timesheets', expectedRows: 500 },
        { name: 'project_documents.csv', sheet: 'Documents', expectedRows: 600 },
        { name: 'purchase_orders.csv', sheet: 'Purchase_Orders', expectedRows: 300 },
        { name: 'expenses.csv', sheet: 'Expenses', expectedRows: 800 },
        { name: 'risks.csv', sheet: 'Risks', expectedRows: 200 }
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

    ss.toast('Starting Dubai PM Data Import...', '🏗️ Dubai PM', -1);

    for (var i = 0; i < CONFIG.FILES.length; i++) {
        var file = CONFIG.FILES[i];
        ss.toast('Importing ' + file.sheet + '...', '🏗️ Import', -1);

        try {
            var rowCount = importFullCSV(ss, baseUrl + file.name, file.sheet);
            totalRows += rowCount;
            results.push({ file: file.sheet, rows: rowCount, status: 'OK' });
            Logger.log('✓ ' + file.sheet + ': ' + rowCount + ' rows imported');
            SpreadsheetApp.flush();
            Utilities.sleep(1000);
        } catch (e) {
            results.push({ file: file.sheet, rows: 0, status: 'ERROR: ' + e.message });
            Logger.log('✗ ' + file.sheet + ': ERROR - ' + e.message);
        }
    }

    var duration = Math.round((new Date() - startTime) / 1000 / 60);
    createSummarySheet(ss, results, totalRows, duration);
    ss.toast('✅ COMPLETE! ' + totalRows + ' rows imported.', 'Done', 5);
}

// ============================================================
// IMPORT SINGLE CSV
// ============================================================
function importFullCSV(ss, url, sheetName) {
    var response = fetchWithRetry(url, 5);
    if (!response) throw new Error('Failed to fetch after 5 retries');

    var csvData = Utilities.parseCsv(response.getContentText());
    if (!csvData || csvData.length === 0) throw new Error('Empty CSV data');

    var totalRows = csvData.length;
    var numCols = csvData[0].length;

    var sheet = ss.getSheetByName(sheetName);
    if (sheet) sheet.clear();
    else sheet = ss.insertSheet(sheetName);

    if (totalRows <= CONFIG.BATCH_SIZE) {
        sheet.getRange(1, 1, totalRows, numCols).setValues(csvData);
    } else {
        writeInBatches(sheet, csvData, CONFIG.BATCH_SIZE);
    }

    formatSheet(sheet, numCols);
    return totalRows;
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
            Utilities.sleep(500);
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
        } catch (e) {
            Utilities.sleep(1000 * Math.pow(2, i));
        }
    }
    return null;
}

function formatSheet(sheet, numCols) {
    try {
        sheet.setFrozenRows(1);
        sheet.getRange(1, 1, 1, numCols).setFontWeight('bold').setBackground('#2c3e50').setFontColor('white');
        sheet.setRowHeight(1, 30);
    } catch (e) { }
}

function createSummarySheet(ss, results, totalRows, duration) {
    var sheet = ss.getSheetByName('Import_Summary');
    if (!sheet) { sheet = ss.insertSheet('Import_Summary', 0); } else { sheet.clear(); }

    var data = [
        ['Dubai Project Management - Import Summary'],
        ['Import Date:', new Date().toLocaleString()],
        ['Total Rows:', totalRows],
        ['Duration (min):', duration],
        [''],
        ['Table', 'Rows', 'Status']
    ];
    results.forEach(function (r) { data.push([r.file, r.rows, r.status]); });

    sheet.getRange(1, 1, data.length, 3).setValues(data);
    sheet.getRange(1, 1, 1, 3).merge().setFontSize(14).setFontWeight('bold');
    sheet.autoResizeColumns(1, 3);
}

function getBaseUrl() {
    return 'https://raw.githubusercontent.com/' + CONFIG.GITHUB_USER + '/' + CONFIG.REPO + '/' + CONFIG.BRANCH + '/';
}

function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('🏗️ Dubai PM Data')
        .addItem('Import All Data', 'importAllData')
        .addToUi();
}
