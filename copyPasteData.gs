function copyPasteData() {
  //This script copy pastes data by indexing column but its unable to identify last row
  const sourceSpreadsheetUrl = 'https://docs.google.com/spreadsheets/d/1yxjESHn65gjXQRTVwy-8U51-pHtl1OkCFu_XJtEsCks/edit?resourcekey=&gid=1741190118#gid=1741190118'; // Replace with your source spreadsheet URL
  const targetSpreadsheetUrl = 'https://docs.google.com/spreadsheets/d/1yxjESHn65gjXQRTVwy-8U51-pHtl1OkCFu_XJtEsCks/edit?resourcekey=&gid=298091860#gid=298091860'; // Replace with your target spreadsheet URL
  const sourceSheetName = 'NPD Initiation'; // Replace with your source sheet name
  const targetSheetName = 'FMS'; // Replace with your target sheet name

  // Open the spreadsheets and sheets
  const sourceSpreadsheet = SpreadsheetApp.openByUrl(sourceSpreadsheetUrl);
  const targetSpreadsheet = SpreadsheetApp.openByUrl(targetSpreadsheetUrl);
  const sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  const targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);

  // Get headers from both sheets
  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0]; // First row
  const targetHeaders = targetSheet.getRange(11, 1, 1, targetSheet.getLastColumn()).getValues()[0]; // First row

  // Get source data (excluding headers)
  const sourceData = sourceSheet.getDataRange().getValues();
  const lastRow = sourceSheet.getRange('D:D').getValues().filter(String).length; // Last non-empty row in column D
  if (lastRow <= 1) return; // Exit if no data

  const rowsToCopy = sourceData.slice(1, lastRow); // Exclude header row

  // Identify the "Copied" column index in the source sheet (Column H assumed here)
  const copiedColumnIndex = sourceHeaders.indexOf('Copied');
  if (copiedColumnIndex === -1) throw new Error('Column "Copied" not found in source sheet.');

  // Filter rows that are not yet marked as "copied"
  const rowsToProcess = rowsToCopy.filter(row => row[copiedColumnIndex] !== 'Copied');

  // Map source columns to target columns based on headers
  const columnMapping = targetHeaders.map(header => sourceHeaders.indexOf(header)); // Match target headers to source headers

  // Prepare data to paste, filling missing columns with empty strings
  const preparedData = rowsToProcess.map(row =>
    columnMapping.map(index => (index !== -1 ? row[index] : '')) // If header is not found in source, leave blank
  );

  // Determine the last row in the target sheet and append the data
  if (preparedData.length > 0) {
    //const lastTargetRow = targetSheet.getLastRow();
    const lastTargetRow = targetSheet.getRange('D:D').getValues().filter(String).length+9;

    const startRow = lastTargetRow + 1; // Append to the row after the last one
    targetSheet.getRange(startRow, 1, preparedData.length, preparedData[0].length).setValues(preparedData);

    // Mark the processed rows as "Copied" in the source sheet
    const copiedRange = sourceSheet.getRange(2, copiedColumnIndex + 1, rowsToCopy.length);
    const updatedCopiedStatus = rowsToCopy.map(row => (row[copiedColumnIndex] === 'Copied' ? 'Copied' : 'Copied'));
    copiedRange.setValues(updatedCopiedStatus.map(status => [status]));
  }
}
