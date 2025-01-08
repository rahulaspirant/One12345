function simpleCopyPasteSpecificFieldsWithStartColumn() {
  const sourceSpreadsheetUrl = 'https://docs.google.com/spreadsheets/d/1yxjESHn65gjXQRTVwy-8U51-pHtl1OkCFu_XJtEsCks/edit?gid=1741190118#gid=1741190118'; // Replace with your source spreadsheet URL
  const targetSpreadsheetUrl = 'https://docs.google.com/spreadsheets/d/1yxjESHn65gjXQRTVwy-8U51-pHtl1OkCFu_XJtEsCks/edit?gid=298091860#gid=298091860'; // Replace with your target spreadsheet URL

  const sourceSheetName = 'NPD Initiation'; // Replace with your source sheet name
  const targetSheetName = 'FMS'; // Replace with your target sheet name

  const targetStartColumn = 4; // Specify the starting column for pasting (e.g., column D = 4)

  // Open the spreadsheets and sheets
  const sourceSpreadsheet = SpreadsheetApp.openByUrl(sourceSpreadsheetUrl);
  const targetSpreadsheet = SpreadsheetApp.openByUrl(targetSpreadsheetUrl);
  const sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  const targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);

  // Define the columns to copy from the source sheet (1-based indexing)
  const sourceColumns = [3, 4, 5, 6, 7]; // Specify the columns to copy, e.g., C, D, E, F, G
  const lastSourceRow = sourceSheet.getLastRow();
  const sourceData = sourceSheet.getRange(2, 1, lastSourceRow - 1, sourceSheet.getLastColumn()).getValues(); // Data starting from row 2

  if (sourceData.length === 0) return; // Exit if no data to copy

  // Filter rows that are not yet marked as "Copied" in Column H (index 7)
  const rowsToCopy = sourceData.filter(row => row[7] !== 'Copied'); // Column H is the 8th column (index 7)

  if (rowsToCopy.length === 0) return; // Exit if no rows to copy

  // Extract only the specified columns
  const selectedData = rowsToCopy.map(row => sourceColumns.map(col => row[col - 1])); // Convert 1-based to 0-based index

  // Define the last row of Column D in the target sheet (considering data starting from row 11)
  const lastTargetRow = targetSheet.getRange('D:D').getValues().filter((val, index) => index >= 11 && val[0] !== '').length + 11;

  // Determine the start row for pasting data
  const startRow = lastTargetRow + 1;
  const targetRange = targetSheet.getRange(startRow, targetStartColumn, selectedData.length, selectedData[0].length);

  // Paste the data
  targetRange.setValues(selectedData);

  // Mark the processed rows as "Copied" in the source sheet
  const copiedColumn = sourceSheet.getRange(2, 8, sourceData.length);
  copiedColumn.setValues(sourceData.map(row => (row[7] === 'Copied' ? ['Copied'] : ['Copied'])));
}
