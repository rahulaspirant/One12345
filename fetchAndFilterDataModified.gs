function fetchAndFilterDataModified() {
  // Replace with your source and target sheet details
  const sourceSpreadsheetId = "1fOIMOvwpsOdsqN6TMA1HZeizl-Bx_Aq8DxjmJHzoPfw";
  const sourceRange = "Combined!A1:AQ";
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

  // Fetch existing data in the target sheet for comparison
  const existingData = targetSheet.getRange(1, 1, targetSheet.getLastRow(), 1).getValues().flat();
  
  // Filter data based on the condition and exclude specified columns
  const headers = sourceData[0];
  const filteredData = sourceData.filter((row, index) => {
    if (index === 0) return true; // Keep headers
    const col9 = row[8]; // Column 9 is index 8 (0-based)
    return (col9 === "Indiamart" || col9 === "Exhibition" || col9 === "Website");
  }).map(row => {
    return row.filter((_, colIndex) => !excludeColumns.includes(colIndex + 1)); // Exclude columns dynamically
  });

  // Append only missing rows
  const newRows = filteredData.filter((row, index) => {
    if (index === 0) return true; // Keep headers
    const uniqueId = row[0]; // Unique ID is in the first column
    return !existingData.includes(uniqueId);
  });

  // Write the new rows to the target sheet
  if (newRows.length > 1) { // Ensure there are rows to append
    targetSheet.getRange(targetSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }
}


Inquiry No
Timestamp
Company
Contact Name
Phone
Email
Category
Details
Lead Source
Sales Person Email
Sales Stage
Update Remarks
Next Steps
Next Followup Date
Forecast Amount
Win Probability
Expected Close Date
Attachments
Delay Days
Dealer Name
Site Name
City
Address
State
Weighted Forecast
Closure Month
Send Status
TS Backup
Sales Person Name
Followup Date
IQ Date
Next Follow up Date
Expected Closing Date
Region
Sales Stage Filter
Inquiry Month
Time Frame
null
Financial Year
Qtr
Month
Week
Date
Yesterday Helper
