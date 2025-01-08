function TimeFrameLookup() {
  // Open the source spreadsheet by its ID
  var sourceSpreadsheetId = '1vHIBTGFRAQ2bGzXO8GuWsfdn99J_JLUBUllwk8D1Qz8'; // Replace with your source spreadsheet ID
  var sourceSheetName = 'File'; // Replace with your source sheet name
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the data from the source sheet, excluding the header row (starting from row 2)
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  var sourceData = sourceSheet.getRange('A2:P' + sourceSheet.getLastRow()).getValues(); // Start from row 2
  
  // Create a map for VLOOKUP
  var lookupMap = {};
  sourceData.forEach(function(row) {
    lookupMap[row[0]] = row[15]; // Assuming Inquiry No is in column A (index 0) and value to fetch is in column P (index 15, and is zero based indexing)
  });
  
  // Get the Inquiry No from the target sheet and perform VLOOKUP
  var targetData = targetSheet.getRange('A3:A' + targetSheet.getLastRow()).getValues(); // Start from row 3
  var results = [];
  
  targetData.forEach(function(row) {
    var inquiryNo = row[0];
    results.push([lookupMap[inquiryNo] || '']); // Push the result or empty string if not found
  });
  
  // Clear previous results in the target range
  targetSheet.getRange(3, 39, targetSheet.getLastRow() - 2, 1).clearContent();

  // Set the results in the target sheet (assuming you want to place results in column B, starting from row 3)
  targetSheet.getRange(3, 39, results.length, 1).setValues(results); // Start from row 3
}
