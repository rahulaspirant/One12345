function updateEmployeeDB() {
  var sourceSheetId = "1FmdTvIrpzn-KxwDwOpJsd6xwCHUhtUW7QQalbiD1DgU";
  var sourceSheetName = "Dropdown";
  var sourceRange = "E1:O";

  var destinationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Emp DB AutoUpdate");

  // Check if "Emp_DB" sheet doesn't exist
  if (destinationSheet === null) {
    // Create the "Emp_DB" sheet
    destinationSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Emp DB AutoUpdate");
  }

  var destinationRange = "A1:K";

  var sourceSheet = SpreadsheetApp.openById(sourceSheetId).getSheetByName(sourceSheetName);
  var valuesToCopy = sourceSheet.getRange(sourceRange).getValues();

  destinationSheet.getRange(destinationRange).clearContent();
  destinationSheet.getRange(1, 1, valuesToCopy.length, valuesToCopy[0].length).setValues(valuesToCopy);

  // Set font for the first row to bold
  destinationSheet.getRange(1, 1, 1, valuesToCopy[0].length).setFontWeight("bold");

  // Freeze the first row
  destinationSheet.setFrozenRows(1);

  // Change font for the whole "Emp_DB" sheet to "Fira Sans Condensed"
  var range = destinationSheet.getDataRange();
  range.setFontFamily("Fira Sans Condensed");

  // Protect and lock the "Emp_DB" sheet
  var protection = destinationSheet.protect();
  protection.setDescription("Read-only protection"); // Optional, add a description for the protection
  protection.setWarningOnly(false); // Set to 'true' to only show a warning, 'false' to completely lock

  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }

  SpreadsheetApp.flush();
}
