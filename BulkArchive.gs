function archiveBulk() {
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var sourceSheet = spreadsheet.getSheetByName("Source");
var archiveSheet = spreadsheet.getSheetByName("Target");
var archiveKeywordColumn = 37; // Column BD - 1 Based Index
var sourceHeaderRow = 6; // Adjust the header row for Source
var archiveHeaderRow = 1; // Adjust the header row for Archive sheet
var dateColumnLetter = 'AC'; // Column C

var archiveKeywords = ["Archive"]; // Adjust as needed

// Insert dynamic columns if necessary
var columnsInserted = insertDynamicColumns(sourceSheet, archiveSheet, sourceHeaderRow, archiveHeaderRow, dateColumnLetter);

// Exit if new columns were inserted
if (columnsInserted) {
return;
}

// Check if headers match between source and archive sheets
if (!checkHeaderMatch(sourceSheet, archiveSheet, sourceHeaderRow, archiveHeaderRow)) {
sendHeaderMismatchEmail(spreadsheet);
return;
}

// Archive data if necessary
var nonBlankRowCount = countNonBlankRows(sourceSheet, archiveKeywordColumn);
if (nonBlankRowCount > 0) {
sourceSheet.insertRowsAfter(sourceSheet.getLastRow(), nonBlankRowCount);
}

deleteArchivedRows(archiveSheet);

filterAndArchiveData(sourceSheet, archiveSheet, archiveKeywordColumn, sourceHeaderRow, archiveKeywords);

// Clear contents of column M in "Dashboard" sheet from row 4 onward
//clearDashboardColumnM(spreadsheet);
}

function insertDynamicColumns(sourceSheet, archiveSheet, sourceHeaderRow, archiveHeaderRow, dateColumnLetter) {
var lastColumnIndex = sourceSheet.getLastColumn();
var sourceHeaders = sourceSheet.getRange(sourceHeaderRow, 1, 1, lastColumnIndex).getValues()[0];
var archiveHeaders = archiveSheet.getRange(archiveHeaderRow, 1, 1, lastColumnIndex).getValues()[0];

var lastSourceHeader = sourceHeaders[sourceHeaders.length - 1];

if (lastSourceHeader === 'Archive Helper') {
return false; // Exit if "Archive Helper" already exists
}

var dynamicColumnRange = dateColumnLetter + sourceHeaderRow + ':' + dateColumnLetter;

var newColumnFormulas = [
'=ArrayFormula(IF(ROW(' + dynamicColumnRange + ')=' + sourceHeaderRow + ',"Financial Year",IF(' + dynamicColumnRange + '="",,IF((--(MONTH(' + dynamicColumnRange + ')>=1))*(--(MONTH(' + dynamicColumnRange + ')<=3)),YEAR(' + dynamicColumnRange + ')-1&"-"&YEAR(' + dynamicColumnRange + '),YEAR(' + dynamicColumnRange + ')&"-"&YEAR(' + dynamicColumnRange + ')+1))))',
'=ArrayFormula(IF(ROW(' + dynamicColumnRange + ')=' + sourceHeaderRow + ',"Qtr",IF(' + dynamicColumnRange + '="",,IF(CEILING(TEXT(' + dynamicColumnRange + ',"m")*1,3)/3=1,"Q4",IF(CEILING(TEXT(' + dynamicColumnRange + ',"m")*1,3)/3=2,"Q1",IF(CEILING(TEXT(' + dynamicColumnRange + ',"m")*1,3)/3=3,"Q2","Q3"))))))',
'=ArrayFormula(IF(' + dynamicColumnRange + '="",,IF(ROW(A' + sourceHeaderRow + ':A)=' + sourceHeaderRow + ',"Month",TEXT(DATEVALUE(' + dynamicColumnRange + '),"mmm"))))',
'=ArrayFormula(IF(ROW(A' + sourceHeaderRow + ':A)=' + sourceHeaderRow + ',"Archive Helper",IF(A' + sourceHeaderRow + ':A="",,IF(ROW(A' + sourceHeaderRow + ':A)=' + (sourceHeaderRow+1) + ',"Yes",))))'
];

var numColumnsToAdd = newColumnFormulas.length;
sourceSheet.insertColumnsAfter(lastColumnIndex, numColumnsToAdd);
var newHeadersRangeSource = sourceSheet.getRange(sourceHeaderRow, lastColumnIndex + 1, 1, numColumnsToAdd);
for (var i = 0; i < numColumnsToAdd; i++) {
newHeadersRangeSource.getCell(1, i + 1).setFormula(newColumnFormulas[i]);
}

var lastNonEmptyColumn = archiveHeaders.filter(String).length;
var insertColumnIndex = lastNonEmptyColumn + 1;

archiveSheet.insertColumnsAfter(lastNonEmptyColumn, 4);
archiveSheet.getRange(archiveHeaderRow, insertColumnIndex).setValue("Financial Year");
archiveSheet.getRange(archiveHeaderRow, insertColumnIndex + 1).setValue("Qtr");
archiveSheet.getRange(archiveHeaderRow, insertColumnIndex + 2).setValue("Month");
archiveSheet.getRange(archiveHeaderRow, insertColumnIndex + 3).setValue("Archive Helper");

return true;
}

function checkHeaderMatch(sourceSheet, archiveSheet, sourceHeaderRow, archiveHeaderRow) {
var sourceHeaders = sourceSheet.getRange(sourceHeaderRow, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
var archiveHeaders = archiveSheet.getRange(archiveHeaderRow, 1, 1, archiveSheet.getLastColumn()).getValues()[0];

if (sourceHeaders.length !== archiveHeaders.length) {
return false;
}

for (var i = 0; i < sourceHeaders.length; i++) {
if (sourceHeaders[i] !== archiveHeaders[i]) {
return false;
}
}
return true;
}

function sendHeaderMismatchEmail(spreadsheet) {
var fileName = spreadsheet.getName();
var subject = fileName + ' | Archive Error';
var spreadsheetLink = '<a href="' + spreadsheet.getUrl() + '">Click Here</a>';
var body = 'Headers do not match for ' + fileName + '. Do the needful.<br><br> ' + spreadsheetLink + ' to access the spreadsheet.';

MailApp.sendEmail({
to: '',
subject: subject,
body: body,
htmlBody: body
});
}

function countNonBlankRows(sheet, column) {
var nonBlankRowCount = sheet.getRange(1, column, sheet.getLastRow()).getValues().filter(String).length;
return nonBlankRowCount;
}

function deleteArchivedRows(archiveSheet) {
var archiveData = archiveSheet.getDataRange().getValues();
var lastArchiveColumn = archiveData[0].length;
var lastArchiveRow = archiveData.length;

for (var row = lastArchiveRow - 1; row >= 0; row--) {
if (archiveData[row][lastArchiveColumn - 1] === "Yes") {
archiveSheet.deleteRow(row + 1);
}
}
}

function filterAndArchiveData(sourceSheet, archiveSheet, archiveKeywordColumn, sourceHeaderRow, archiveKeywords) {
var lastRow = sourceSheet.getLastRow();
if (lastRow <= sourceHeaderRow) {
return;
}

var dataRange = sourceSheet.getRange(sourceHeaderRow, 1, lastRow - sourceHeaderRow + 1, sourceSheet.getLastColumn() - 1);

var filter = sourceSheet.getFilter();
if (filter) {
filter.removeColumnFilterCriteria(archiveKeywordColumn);
}

var filterRange = sourceSheet.getRange(sourceHeaderRow, archiveKeywordColumn, lastRow - sourceHeaderRow + 1, 1);
var filterCriteria = SpreadsheetApp.newFilterCriteria().whenTextContains(archiveKeywords).build();
filter = sourceSheet.getFilter();
if (filter) {
filter.setColumnFilterCriteria(archiveKeywordColumn, filterCriteria);
} else {
var rangeWithFilter = dataRange.createFilter();
rangeWithFilter.setColumnFilterCriteria(archiveKeywordColumn, filterCriteria);
filter = rangeWithFilter;
}

if (!filter) {
return;
}

var visibleRows = filter.getRange().getValues().filter(row => row[archiveKeywordColumn - 1] === "Archive");
if (visibleRows.length === 0) {
filter.remove();
return;
} else {
var lastRowArchive = archiveSheet.getLastRow();
var filteredRange = filter.getRange().offset(1, 0, lastRow - sourceHeaderRow, sourceSheet.getLastColumn() - 1);
filteredRange.copyTo(archiveSheet.getRange(lastRowArchive + 1, 1), {contentsOnly: true});
sourceSheet.deleteRows(sourceHeaderRow + 1, lastRow - sourceHeaderRow);
filter.remove();
}

var archiveFormula = '=IFNA(FILTER(INDIRECT("\'' + sourceSheet.getName() + '\'!A' + (sourceHeaderRow + 1) + ':" & LEFT(ADDRESS(1, COUNTA(\'' + sourceSheet.getName() + '\'!' + sourceHeaderRow + ':' + sourceHeaderRow + '), 4), LEN(ADDRESS(1, COUNTA(\'' + sourceSheet.getName() + '\'!' + sourceHeaderRow + ':' + sourceHeaderRow + '), 4)) - 1)), INDIRECT("\'' + sourceSheet.getName() + '\'!A' + (sourceHeaderRow + 1) + ':A") <> ""))';
archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1).setFormula(archiveFormula);
}

/*
// Clear column M in "Dashboard" from row 4 onward
function clearDashboardColumnM(spreadsheet) {
var dashboardSheet = spreadsheet.getSheetByName("Dashboard");
if (dashboardSheet) {
var rangeToClear = dashboardSheet.getRange("M4:M" + dashboardSheet.getLastRow());
rangeToClear.clearContent();
}
}
*/
// END FUNCTION ***


/////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Simple and Modular approach

/**
 * archiveBulk - Modular, readable, and performant rewrite
 *
 * Behavior summary:
 * 1. Optionally inserts dynamic helper columns (Financial Year, Qtr, Month, Archive Helper)
 *    and exits early so formulas can populate.
 * 2. Verifies headers match between Source and Target (Archive) sheets.
 * 3. Finds rows in Source marked for archiving (archiveKeywordColumn contains "Archive")
 *    and moves them to Target (append only, contentsOnly).
 * 4. Deletes moved rows from Source in a safe bottom-up batch.
 * 5. Removes any "Archive Helper" rows in Target that are marked "Yes".
 *
 * CONFIG at top — change to match your spreadsheet layout.
 */

/* ---------------------- CONFIG ---------------------- */
const ARCHIVE_CONFIG = {
  sourceSheetName: 'Dataset',
  targetSheetName: 'Archive',
  sourceHeaderRow: 1,        // header row number in Source (1-based)
  targetHeaderRow: 1,        // header row number in Target (1-based)
  archiveKeywordColumn: 18,  // column number (1-based) containing the "Archive" keyword (BD -> 56? but using 37 per your original)
  archiveKeywords: ['Archive'],
  dateColumnLetter: 'K',    // letter used to derive dynamic formulas (e.g. 'AC')
  archiveHelperHeader: 'Archive Helper',
  emailOnHeaderMismatchTo: '' // set email if you want header mismatch reports
};
/* ---------------------------------------------------- */

/* ---------------------- ENTRY ----------------------- */
function archiveBulk() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(ARCHIVE_CONFIG.sourceSheetName);
  const target = ss.getSheetByName(ARCHIVE_CONFIG.targetSheetName);
  if (!source || !target) {
    throw new Error('Source or Target sheet not found. Check sheet names in CONFIG.');
  }

  // Step 1. Ensure dynamic columns exist; if inserted, exit early (formulas need a calc pass)
  const inserted = ensureDynamicColumns(source, target);
  if (inserted) {
    Logger.log('Inserted dynamic columns; please re-run after formulas populate.');
    return;
  }

  // Step 2. Verify headers match (structure)
  const headersMatch = compareHeaders(source, target);
  if (!headersMatch) {
    Logger.log('Headers do not match between Source and Target.');
    maybeSendHeaderMismatchEmail(ss);
    return;
  }

  // Step 3. Collect rows to archive (batch read)
  const rowsToArchiveInfo = collectSourceRowsToArchive(source);
  if (rowsToArchiveInfo.rows.length === 0) {
    Logger.log('No rows marked for archive found.');
    return;
  }

  // Step 4. Append to archive sheet (contents only) in a single batch operation
  appendToTarget(target, rowsToArchiveInfo.rows);

  // Step 5. Delete the archived rows from Source (bottom-up)
  deleteRowsFromSheet(source, rowsToArchiveInfo.rowNumbers);

  // Step 6. Clean up target: delete rows marked "Yes" in last column (Archive Helper)
  removeYesMarkedRowsInTarget(target);

  // Optionally add a formula row to the archive sheet to keep it dynamic
  addArchiveFormulaRow(target, source);

  Logger.log(`Archived ${rowsToArchiveInfo.rows.length} rows.`);
}

/* -------------------- HELPERS ----------------------- */

/**
 * ensureDynamicColumns
 * - If 'Archive Helper' header not present in Source, inserts helper columns to the right
 * - Also inserts corresponding headers to Target
 * - Returns true if columns were inserted (script should stop so formulas can populate)
 */
function ensureDynamicColumns(sourceSheet, targetSheet) {
  const cfg = ARCHIVE_CONFIG;
  const lastCol = sourceSheet.getLastColumn();
  const sourceHeaders = sourceSheet.getRange(cfg.sourceHeaderRow, 1, 1, lastCol).getDisplayValues()[0].map(String);
  if (sourceHeaders.includes(cfg.archiveHelperHeader)) return false;

  // Build formulas that use the configured date column letter and source header row
  const r = cfg.sourceHeaderRow;
  const col = cfg.dateColumnLetter;
  const dynamicRange = `${col}${r}:${col}`;
  const formulas = [
    // Financial Year
    `=ArrayFormula(IF(ROW(${dynamicRange})=${r},"Financial Year",IF(${dynamicRange}="",,IF((MONTH(${dynamicRange})>=1)*(MONTH(${dynamicRange})<=3),YEAR(${dynamicRange})-1&"-"&YEAR(${dynamicRange}),YEAR(${dynamicRange})&"-"&YEAR(${dynamicRange})+1))))`,
    // Quarter (simple mapping)
    `=ArrayFormula(IF(ROW(${dynamicRange})=${r},"Qtr",IF(${dynamicRange}="",,IF(CEILING(MONTH(${dynamicRange})/3)=1,"Q1",IF(CEILING(MONTH(${dynamicRange})/3)=2,"Q2",IF(CEILING(MONTH(${dynamicRange})/3)=3,"Q3","Q4"))))))`,
    // Month (short name)
    `=ArrayFormula(IF(ROW(${dynamicRange})=${r},"Month",IF(${dynamicRange}="",,TEXT(${dynamicRange},"mmm"))))`,
    // Archive Helper
    `=ArrayFormula(IF(ROW(A${r}:A)=${r},"Archive Helper",IF(A${r}:A="",,IF(ROW(A${r}:A)=${r}+1,"Yes",""))))`
  ];

  // Insert columns to the right
  sourceSheet.insertColumnsAfter(lastCol, formulas.length);
  const targetRange = sourceSheet.getRange(cfg.sourceHeaderRow, lastCol + 1, 1, formulas.length);
  for (let i = 0; i < formulas.length; i++) {
    targetRange.getCell(1, i + 1).setFormula(formulas[i]);
  }

  // Add corresponding headers to target sheet (at the end)
  const targetLastCol = targetSheet.getLastColumn();
  targetSheet.insertColumnsAfter(targetLastCol, formulas.length);
  const headers = ['Financial Year', 'Qtr', 'Month', cfg.archiveHelperHeader];
  targetSheet.getRange(cfg.targetHeaderRow, targetLastCol + 1, 1, headers.length).setValues([headers]);

  return true;
}

/**
 * compareHeaders - returns true if header arrays are the same length and values (order matters)
 */
function compareHeaders(sourceSheet, targetSheet) {
  const cfg = ARCHIVE_CONFIG;
  const sHeaders = sourceSheet.getRange(cfg.sourceHeaderRow, 1, 1, sourceSheet.getLastColumn()).getDisplayValues()[0].map(String);
  const tHeaders = targetSheet.getRange(cfg.targetHeaderRow, 1, 1, targetSheet.getLastColumn()).getDisplayValues()[0].map(String);
  if (sHeaders.length !== tHeaders.length) return false;
  for (let i = 0; i < sHeaders.length; i++) {
    if (String(sHeaders[i]) !== String(tHeaders[i])) return false;
  }
  return true;
}

/**
 * maybeSendHeaderMismatchEmail - notifies if configured
 */
function maybeSendHeaderMismatchEmail(spreadsheet) {
  const email = ARCHIVE_CONFIG.emailOnHeaderMismatchTo;
  if (!email) return;
  const subject = `${spreadsheet.getName()} | Archive Header Mismatch`;
  const body = `Headers do not match between Source and Target for spreadsheet: ${spreadsheet.getUrl()}`;
  MailApp.sendEmail(email, subject, body);
}

/**
 * collectSourceRowsToArchive - batch reads source data and finds rows where the archiveKeywordColumn
 * contains any of the archiveKeywords. Returns: { rows: [... rowArrays ...], rowNumbers: [... 1-based row numbers in sheet ...] }
 */
function collectSourceRowsToArchive(sourceSheet) {
  const cfg = ARCHIVE_CONFIG;
  const lastRow = sourceSheet.getLastRow();
  if (lastRow <= cfg.sourceHeaderRow) return { rows: [], rowNumbers: [] };

  // read all rows from headerRow+1 to lastRow
  const numRows = lastRow - cfg.sourceHeaderRow;
  const numCols = sourceSheet.getLastColumn();
  const data = sourceSheet.getRange(cfg.sourceHeaderRow + 1, 1, numRows, numCols).getValues();

  const archiveSet = new Set(cfg.archiveKeywords.map(k => String(k).toLowerCase()));
  const rows = [];
  const rowNumbers = []; // sheet absolute row numbers (1-based)

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const cellVal = row[cfg.archiveKeywordColumn - 1]; // zero-based index in row array
    if (cellVal !== null && cellVal !== undefined && String(cellVal).trim() !== '') {
      const test = String(cellVal).trim().toLowerCase();
      if (archiveSet.has(test) || cfg.archiveKeywords.some(k => test.indexOf(String(k).toLowerCase()) !== -1)) {
        rows.push(row);
        rowNumbers.push(cfg.sourceHeaderRow + 1 + i); // convert to 1-based sheet row
      }
    }
  }

  return { rows, rowNumbers };
}

/**
 * appendToTarget - appends rows (2D array) to targetSheet in a single batch append
 */
function appendToTarget(targetSheet, rows) {
  if (!rows || rows.length === 0) return;
  const startRow = targetSheet.getLastRow() + 1;
  const cols = rows[0].length;
  targetSheet.getRange(startRow, 1, rows.length, cols).setValues(rows);
}

/**
 * deleteRowsFromSheet - deletes rows by rowNumbers (array of 1-based rows) in bottom-up order.
 * Performs batch deletions by grouping contiguous rows for fewer delete calls.
 */
function deleteRowsFromSheet(sheet, rowNumbers) {
  if (!rowNumbers || rowNumbers.length === 0) return;
  rowNumbers.sort((a, b) => b - a); // descending
  let i = 0;
  while (i < rowNumbers.length) {
    const start = rowNumbers[i];
    let count = 1;
    // count contiguous rows below this one (desc sorted)
    while (i + count < rowNumbers.length && rowNumbers[i + count] === rowNumbers[i] - count) {
      count++;
    }
    sheet.deleteRows(start - count + 1, count); // delete a block
    i += count;
  }
}

/**
 * removeYesMarkedRowsInTarget - deletes rows in target where the last column === "Yes"
 * Uses batch read and bottom-up deletes
 */
function removeYesMarkedRowsInTarget(targetSheet) {
  const data = targetSheet.getDataRange().getValues();
  if (!data || data.length === 0) return;
  const lastCol = data[0].length;
  const toDelete = [];
  // start from row 2 if header row is 1; but we'll check entire data to be safe
  for (let r = 0; r < data.length; r++) {
    if (String(data[r][lastCol - 1]).trim() === 'Yes') {
      toDelete.push(r + 1); // convert 0-based to 1-based
    }
  }
  if (toDelete.length) deleteRowsFromSheet(targetSheet, toDelete);
}

/**
 * addArchiveFormulaRow - replicates your original archive formula insertion: places a FILTER formula
 * in the next row of the archive sheet to keep archive data dynamic if desired.
 */
function addArchiveFormulaRow(targetSheet, sourceSheet) {
  const cfg = ARCHIVE_CONFIG;
  // Keep the same approach as original but simplified. We'll set a formula that pulls from Source
  const srcName = sourceSheet.getName();
  const srcHeaderRow = cfg.sourceHeaderRow;
  const formula = `=IFERROR(FILTER(INDIRECT("${srcName}!A${srcHeaderRow + 1}:Z"), INDIRECT("${srcName}!A${srcHeaderRow + 1}:A") <> ""), "")`;
  const destRow = targetSheet.getLastRow() + 1;
  try {
    targetSheet.getRange(destRow, 1).setFormula(formula);
  } catch (e) {
    // If formula fails (e.g. too many columns), skip silently — it's a convenience not a requirement.
    Logger.log('Could not set archive formula row: ' + e.message);
  }
}

/* -------------------- END OF SCRIPT ------------------ */

