function CreateUniqueList() {
  // Get the active spreadsheet and target sheet (where new numbers will be added)
  var targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = targetSpreadsheet.getSheetByName('File'); // Replace 'TargetSheet' with your target sheet name

  // Ensure the target sheet exists
  if (!targetSheet) {
    Logger.log('Target sheet "TargetSheet" does not exist!');
    return;
  }

  // Clear existing data in the target sheet (optional: use .clearContents() to only clear data, not formatting)
  //targetSheet.getRange('A2:A').clearContents();  // Clears data from column A starting from row 2

  // Fetch the source spreadsheet by its ID
  var sourceSpreadsheetId = '1fOIMOvwpsOdsqN6TMA1HZeizl-Bx_Aq8DxjmJHzoPfw'; // Replace with your source spreadsheet ID
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSpreadsheet.getSheetByName('Combined'); // Replace with your source sheet name

  // Ensure the source sheet exists
  if (!sourceSheet) {
    Logger.log('Source sheet "Sheet1" does not exist!');
    return;
  }

  // Get all the values from the source sheet (column A)
  var sourceRange = sourceSheet.getRange('A2:A'); // Assuming header row is in A1
  var sourceValues = sourceRange.getValues();

  // Get all existing values in the target sheet (column A)
  var targetRange = targetSheet.getRange('A2:A');
  var targetValues = targetRange.getValues().flat(); // Flatten the 2D array to 1D
  
  // Create a Set for faster lookup
  var targetSet = new Set(targetValues);

  // Prepare an array to hold new numbers to add
  var newNumbers = [];

  // Iterate through the source values and check for new numbers
  for (var i = 0; i < sourceValues.length; i++) {
    var number = sourceValues[i][0];

    // If the number does not exist in the target sheet, add it to the newNumbers array
    if (number && !targetSet.has(number)) {
      newNumbers.push([number]);
      targetSet.add(number); // Add the number to the Set to avoid future duplication
    }
  }

  // If there are new numbers, append them to the target sheet
  if (newNumbers.length > 0) {
  var lastRow = targetSheet.getRange("A:A").getValues().filter(String).length;
   targetSheet.getRange(lastRow + 1, 1, newNumbers.length, 1).setValues(newNumbers);


    //targetSheet.getRange(targetSheet.getLastRow() + 1, 1, newNumbers.length, 1).setValues(newNumbers);
  }
}
