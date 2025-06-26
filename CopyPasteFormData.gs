function CopyPasteFormData() {
  // User settings for data copy behavior
  const copyMode = "New"; // Options: "All" (copy all rows) or "New" (only uncopied rows)
  const copyColumnsMode = "All"; // Options: "All" (all columns) or "Specific" (selected columns only)

  // Sheet and range setup
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheetName = 'FormData';
  const targetSheetName = 'FMS';
  const targetStartColumn = 1; // 1-based index for column A

  const sourceSheet = ss.getSheetByName(sourceSheetName);
  const targetSheet = ss.getSheetByName(targetSheetName);

  const specificColumns = [3, 4, 5, 6, 7]; // Used only if copyColumnsMode is "Specific"

  // Get header row to dynamically find the "Copied" column
  const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  const copiedColIndex = headers.indexOf("Copied Status") + 1; // Convert to 1-based index for getRange
  const copiedDataIndex = copiedColIndex - 1; // 0-based index for working with arrays

  if (copiedColIndex === 0) throw new Error(" 'Copied' column not found.");

  // Read all data (excluding header)
  const lastSourceRow = sourceSheet.getLastRow();
  const sourceData = sourceSheet.getRange(2, 1, lastSourceRow - 1, sourceSheet.getLastColumn()).getValues();

  if (sourceData.length === 0) return; // Exit if no data

  let rowsToCopy = [];

  // Filter rows based on copyMode
  if (copyMode.toLowerCase() === 'all') {
    rowsToCopy = sourceData;
  } else if (copyMode.toLowerCase() === 'new') {
    rowsToCopy = sourceData.filter(row => row[copiedDataIndex] !== 'Copied');
  } else {
    throw new Error('Invalid copy mode. Use "All" or "New".');
  }

  if (rowsToCopy.length === 0) return; // Nothing to copy

  let selectedData = [];

  // Decide which columns to copy
  if (copyColumnsMode.toLowerCase() === 'all') {
    selectedData = rowsToCopy;
  } else if (copyColumnsMode.toLowerCase() === 'specific') {
    selectedData = rowsToCopy.map(row => specificColumns.map(col => row[col - 1]));
  } else {
    throw new Error('Invalid column copy mode. Use "All" or "Specific".');
  }

  // Determine the insertion point in the target sheet
  const targetColumnD = targetSheet.getRange('D:D').getValues();
  const lastTargetRow = targetColumnD.filter((val, idx) => idx >= 4 && val[0] !== '').length + 5;
  const startRow = lastTargetRow + 1;

  // Paste copied data into the target sheet
  targetRange = targetSheet.getRange(startRow, targetStartColumn, selectedData.length, selectedData[0].length);
  targetRange.setValues(selectedData);

  // Mark copied rows in source sheet (only if "New" mode is selected)
  if (copyMode.toLowerCase() === 'new') {
    let updateCount = 0;
    const updatedFlags = sourceData.map(row => {
      if (row[copiedDataIndex] !== 'Copied' && updateCount < rowsToCopy.length) {
        updateCount++;
        return ['Copied']; // Mark this row
      } else {
        return [row[copiedDataIndex]]; // Retain original value
      }
    });

    // Update the "Copied" column with new values
    sourceSheet.getRange(2, copiedColIndex, updatedFlags.length, 1).setValues(updatedFlags);
  }
}
