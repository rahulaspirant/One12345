function fetchAndFilterDataModified() {
  // Replace with your source and target sheet details
  const sourceSpreadsheetId = "1fOIMOvwpsOdsqN6TMA1HZeizl-Bx_Aq8DxjmJHzoPfw";
  const sourceRange = "Combined!A2:AN";
  const targetSheetName = "Exhibition Indiamart Website"; // Replace with your target sheet name
  const excludeColumns = [28, 35]; // Specify column indices (1-based) to exclude
  
  // Open the target spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = ss.getSheetByName(targetSheetName);

  // If the target sheet doesn't exist, create it
  if (!targetSheet) {
    targetSheet = ss.insertSheet(targetSheetName);
  }

  // Fetch the source data
  const sourceData = SpreadsheetApp.openById(sourceSpreadsheetId)
    .getRangeByName(sourceRange)
    .getValues();

  // Filter data based on the condition and exclude specified columns
  const headers = sourceData[0];
  const filteredData = sourceData.filter((row, index) => {
    if (index === 0) return true; // Keep headers
    const col9 = row[8]; // Column 9 is index 8 (0-based)
    return (col9 === "Indiamart" || col9 === "Exhibition" || col9 === "Website");
  }).map(row => {
    return row.filter((_, colIndex) => !excludeColumns.includes(colIndex + 1)); // Exclude columns dynamically
  });

  // Clear existing data in the target sheet (excluding headers)
if (targetSheet.getLastRow() > 1) {
  targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clearContent();
}

// Write the filtered data to the target sheet
if (filteredData.length > 0) { // Ensure there are rows to append
  targetSheet.getRange(2, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
}}
