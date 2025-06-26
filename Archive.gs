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
