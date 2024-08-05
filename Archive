function archiveContracts() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName("File");
  var archiveSheet = spreadsheet.getSheetByName("Archive");

  var dataRange = sourceSheet.getDataRange();
  var lastColumnIndex = dataRange.getLastColumn();
  var conditionColumnIndex = 35; // Column U (1 based index)

  var sourceHeaders = sourceSheet.getRange(6, 1, 1, lastColumnIndex).getValues()[0]; // get the last element of array header
  var archiveHeaders = archiveSheet.getRange(1, 1, 1, lastColumnIndex).getValues()[0]; //get the last element of array header

//Add Helper Columns if Not Present:
  if (sourceHeaders[sourceHeaders.length - 1] !== "Archive Helper" && sourceHeaders[sourceHeaders.length - 1] !== "") {
    var formulas = [                                   //if header last element is not "Archive Helper" and also not empty
      '=ArrayFormula(IF(ROW(A6:A)=6,"Year",IF(A6:A="",,TEXT(F6:F,"yyy")*1)))',
      '=ArrayFormula(IF(ROW(A6:A)=6,"Qtr",IF(A6:A="",,IF(CEILING(TEXT(A6:A,"m")*1,3)/3=1,"Q4",IF(CEILING(TEXT(A6:A,"m")*1,3)/3=2,"Q1",IF(CEILING(TEXT(A6:A,"m")*1,3)/3=3,"Q2","Q3"))))))',
      '=ArrayFormula(IF(ROW(A6:A)=1,"Year - Qtr",IF(AJ6:AJ="",,AJ6:AJ&" - "&AK6:AK)))', //This needs to be checked everytime
      '=ArrayFormula(IF(A6:A="",,IF(ROW(A6:A)=6,"Financial Year",IF((--(MONTH(A6:A)>=1))*(--(MONTH(A6:A)<=3)),YEAR(A6:A)-1&"-"&YEAR(A6:A),YEAR(A6:A)&"-"&YEAR(A6:A)+1))))',
      '={"Archive Helper";"Yes"}'
    ];

    var headerRange = sourceSheet.getRange(headerRow, lastColumnIndex + 1, 1, formulas.length);
    headerRange.setFormulas([formulas]);
  }

//Check Header Match and Send Email if Not:

  for (var i = 0; i < sourceHeaders.length; i++) {
    if (sourceHeaders[i] !== archiveHeaders[i]) {
      var fileName = spreadsheet.getName();
      var subject = fileName + ' | Headers do not match for Archive';
      var body = 'Headers do not match for ' + fileName + ' for running the Archive function. Do the needful. Link to the spreadsheet: ' + spreadsheet.getUrl();
      MailApp.sendEmail({
        to: 'rahul.solanki@arihant.com',
        subject: subject,
        body: body
      });
      return;
    }
  }

  var archiveData = archiveSheet.getDataRange().getValues();
  for (var row = archiveData.length - 1; row > 0; row--) {
    var lastColumnValue = archiveData[row][lastColumnIndex - 1];
    if (lastColumnValue === "Yes") {
      archiveSheet.deleteRow(row + 1);
    }
  }

  var sourceData = sourceSheet.getDataRange().getValues();
  var hasDataInConditionColumn = false; // Flag to track if there is any data meeting the condition
  for (var row = sourceData.length - 1; row >= 1; row--) { // Start processing from row 2 onwards
    var conditionValue = sourceData[row][conditionColumnIndex - 1]; // Adjust index to match JavaScript's 0-based indexing
    if (conditionValue === "Archive" || conditionValue === "Done") {
      var rowDataToArchive = sourceData[row].slice(0, lastColumnIndex - 1); // Exclude last column
      archiveSheet.appendRow(rowDataToArchive);
      sourceSheet.deleteRow(row + 1);
      hasDataInConditionColumn = true; // Set the flag to true if there is data meeting the condition
    }
  }

  var lastNonBlankRow = archiveSheet.getLastRow() + 1;
  if (!hasDataInConditionColumn) {
    lastNonBlankRow = Math.max(lastNonBlankRow, 2); // Start from row 2 if there is no data meeting the condition
  }

  var formulaCell = archiveSheet.getRange(lastNonBlankRow, 1);
  formulaCell.setFormula('=IFNA(FILTER(INDIRECT("\'File\'!A7:" & LEFT(ADDRESS(1, COUNTA(\'File\'!7:7), 4), LEN(ADDRESS(1, COUNTA(\'File\'!7:7), 4)) - 1)), INDIRECT("\'File\'!A7:A") <> ""))');
  // row number to be changed accordingly
}
