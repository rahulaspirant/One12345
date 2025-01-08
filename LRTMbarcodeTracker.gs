function copyDataToProduction() {
  // Source file
  var sourceFile = SpreadsheetApp.openById("1z_o8occHdJ1Hpv7LQJpEIwKmjAgjl_tMQjMQs9wKJnM"); // Replace with the ID of your source file
  var sourceSheet = sourceFile.getSheetByName("Production Helper");

  // Destination file
  var destFile = SpreadsheetApp.openById("1-WRSKJichxX_lP_zaWtiIzE_QHtBy_zMYlB6jFVFNgw"); // Replace with the ID of your destination file
  var destSheet = destFile.getSheetByName("Master");

  // Check if cell A2 in destination sheet is blank
  var destA2Value = destSheet.getRange("A2").getValue();
  var destRowIndex = destA2Value === "" ? 2 : destSheet.getLastRow() + 1;

  // Get data from source sheet
  var data = sourceSheet.getDataRange().getValues();

  // Iterate through rows starting from the second row (skipping header)
  for (var i = 1; i < data.length; i++) {
    // Check if "Copied" is mentioned in column H (index 7)
    if (data[i][7] === "Copied") {
      // Skip this row as it is already copied
      continue;
    }

    // Check if column H is blank
    if (data[i][7] === "") {
      // Copy data from column A to I to destination sheet
      destSheet.getRange(destRowIndex, 1, 1, 7).setValues([data[i].slice(0, 7)]); // Set values for the destination row

      // Write "Copied" in column H of source sheet
      sourceSheet.getRange(i + 1, 8).setValue("Copied");

      // Increment the destination row index
      destRowIndex++;
    }
  }
}





function copyDataToPrintBarcode() {
  // Get source and destination sheets
  var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master");
  var destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Print Barcode");

  // Get data from source sheet
  var data = sourceSheet.getDataRange().getValues();

  // Get the last row in the destination sheet
  var destLastRow = destSheet.getLastRow();

  // Calculate the starting row in the destination sheet (starting from row 3)
  var destStartRow = destLastRow < 2 || destSheet.getRange(3, 2).getValue() === '' ? 3 : destLastRow + 1;

  // Iterate through rows starting from the second row (skipping header)
  for (var i = 1; i < data.length; i++) {
    // Check if "Copied" is mentioned in column O (0 Based index 11)
    if (data[i][15] !== "Copied") {
      // Copy data from columns A, B, D, and E to destination sheet
      destSheet.getRange(destStartRow, 2, 1, 2).setValues([[data[i][4], data[i][0]]]); // Columns B and C
      destSheet.getRange(destStartRow, 4, 1, 2).setValues([[data[i][1], data[i][3]]]); // Columns D and E

      // Write "Copied" in column O of source sheet
      sourceSheet.getRange(i + 1, 16).setValue("Copied");

      // Increment the destination start row for the next iteration
      destStartRow++;
    }
  }
}






function clearDataInPrintBarcode() {
  // Get destination sheet
  var destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Print Barcode");

  // Get the data range starting from row 2 in columns B and C
  var dataRange = destSheet.getRange(3, 2, destSheet.getLastRow() - 1, 4);

  // Clear the data in the range
  dataRange.clearContent();
}





function archiveProduction() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName("Master");
  var archiveSheet = spreadsheet.getSheetByName("Archive");

  var archiveKeywordColumn = 18; // Column R 1 Based Index
  var sourceHeaderRow = 1;  // Adjust the header row for Source
  var archiveHeaderRow = 1; // Adjust the header row for Archive sheet

  // Define an array of archive keywords
  var archiveKeywords = ["Archive","Done","Copied"]; // Adjust as needed

  var lastColumnIndex = sourceSheet.getDataRange().getLastColumn();
  var sourceHeaders = sourceSheet.getRange(sourceHeaderRow, 1, 1, lastColumnIndex).getValues()[0];
  var archiveHeaders = archiveSheet.getRange(archiveHeaderRow, 1, 1, lastColumnIndex).getValues()[0];

  // Check header match
  for (var i = 0; i < sourceHeaders.length; i++) {
    if (sourceHeaders[i] !== archiveHeaders[i]) {
      var fileName = spreadsheet.getName();
      var subject = fileName + ' | Headers do not match for Archive';
      var body = 'Headers do not match for ' + fileName + '. Do the needful.';
      MailApp.sendEmail({ to: 'central.data@arihant.com', subject: subject, body: body });
      return;
    }
  }

  // Delete all rows in Archive sheet where last column contains "Yes"
  var archiveData = archiveSheet.getDataRange().getValues();
  var lastArchiveRow = archiveData.length;

  for (var row = lastArchiveRow - 1; row >= 0; row--) {
    var lastColumnValue = archiveData[row][lastColumnIndex - 1];

    if (lastColumnValue === "Yes") {
      archiveSheet.deleteRow(row + 1); // Adding 1 for one-based index
    }
  }


  // Copy data as values to Archive sheet, skipping the last column
  var sourceData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, lastColumnIndex - 1).getValues();
  for (var row = sourceData.length - 1; row >= 0; row--) {
    var conditionValue = sourceData[row][archiveKeywordColumn - 1];
    if (archiveKeywords.indexOf(conditionValue) !== -1) {
      archiveSheet.appendRow(sourceData[row].slice(0, lastColumnIndex - 1)); // Copy all columns except the last one
    }
  }

  // Delete rows from Master that were copied to Archive
  for (var row = sourceData.length - 1; row >= 0; row--) {
    var conditionValue = sourceData[row][archiveKeywordColumn - 1];
    if (archiveKeywords.indexOf(conditionValue) !== -1) {
      sourceSheet.deleteRow(row + 2); // Adding 2 for one-based index and starting from row 2
    }
  }

  // Update Archive sheet formula with the dynamically obtained sheet name
  var sheetName = sourceSheet.getName();
  var lastNonBlankRow = archiveSheet.getLastRow() + 1;
  archiveSheet.getRange(lastNonBlankRow, 1).setFormula('=IFNA(FILTER(INDIRECT("\'' + sheetName + '\'!A2:" & LEFT(ADDRESS(1, COUNTA(\'' + sheetName + '\'!1:1), 4), LEN(ADDRESS(1, COUNTA(\'' + sheetName + '\'!1:1), 4)) - 1)), INDIRECT("\'' + sheetName + '\'!A2:A") <> ""))');
}


