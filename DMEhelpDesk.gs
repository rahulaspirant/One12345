//const updateEmployeeDB = () => AICLScripts.updateEmployeeDB();
//const changeFont = () => AICLScripts.changeFont();
//const generateSrNumbers = () => AICLScripts.generateSrNumbers();
//const copyData = () => AICLScripts.copyData();
//const processTimestamp = () => AICLScripts.processTimestamp();
//const supportTimestamp = () => AICLScripts.supportTimestamp();
//const archiveHelpdeskForm = () => AICLScripts.archiveHelpdeskForm();
//const archiveProcess = () => AICLScripts.archiveProcess();
//const archiveSupport = () => AICLScripts.archiveSupport();
//const createTriggersforDME = () => AICLScripts.createTriggersforDME();



function updateEmployeeDB() {
  var sourceSheetId = "1FmdTvIrpzn-KxwDwOpJsd6xwCHUhtUW7QQalbiD1DgU";
  var sourceSheetName = "Dropdown";
  var sourceRange = "E1:O";

  var destinationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Emp_DB");

  // Check if "Emp_DB" sheet doesn't exist
  if (destinationSheet === null) {
    // Create the "Emp_DB" sheet
    destinationSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Emp_DB");
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


// Changes the font of the sheet to Fira Sans Condensed
function changeFont() {
  var sheetName = "Helpdesk Form";
  var fontName = "Fira Sans Condensed";
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  // Check if the sheet "Purchase Bills" exists
  if (sheet) {
    var range = sheet.getDataRange();
    range.setFontFamily(fontName);
  } else {
    // Sheet "Helpdesk Form" not found in the current spreadsheet.
  }
}





// Generates Unique Sr. No. even if the row was deleted & stores on note of the header
function generateSrNumbers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Helpdesk Form");
  var dataRange = sheet.getRange("A2:A");
  var data = dataRange.getValues();
  var serialNumberRange = sheet.getRange("B2:B");
  var serialNumbers = serialNumberRange.getValues();
  
  var lastSerialNumber = parseInt(sheet.getRange("B1").getNotes(), 10) || 0;

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] !== "" && serialNumbers[i][0] === "") {
      lastSerialNumber++;
      serialNumbers[i][0] = lastSerialNumber;
    }
  }
  
  serialNumberRange.setValues(serialNumbers);
  sheet.getRange("B1").setNotes([[lastSerialNumber.toString()]]);
}



function copyData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var helpdeskFormSheet = ss.getSheetByName("Helpdesk Form");
  var processCreationSheet = ss.getSheetByName("Process Creation");
  var helpdeskSheet = ss.getSheetByName("Support Ticket");
  var lastRow = helpdeskFormSheet.getLastRow();
  
  var lastProcessRow = processCreationSheet.getLastRow();
  var lastHelpdeskRow = helpdeskSheet.getLastRow();

  var destinationRowProcess = lastProcessRow + 1; // Start from the next row after the last one
  var destinationRowHelpdesk = lastHelpdeskRow + 1;

  for (var i = 2; i <= lastRow; i++) {
    var rowData = helpdeskFormSheet.getRange(i, 1, 1, 9).getValues()[0];
    var condition = helpdeskFormSheet.getRange(i, 11).getValue();
    var assigned = helpdeskFormSheet.getRange(i, 12).getValue();

    if (assigned === "Copied") {
      continue; // skip this row and move to the next one
    }

    if (condition === "New Process Creation") {
      processCreationSheet.getRange(destinationRowProcess, 1, 1, rowData.length).setValues([rowData]);
      destinationRowProcess++;
    } else if (condition === "Support Request") {
      helpdeskSheet.getRange(destinationRowHelpdesk, 1, 1, rowData.length).setValues([rowData]);
      destinationRowHelpdesk++;
    }

    helpdeskFormSheet.getRange(i, 12).setValue("Copied");
  }
}




//Process Done Timestamp
function processTimestamp() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getActiveRange();
  
  if (sheet.getName() === "Process Creation") {
    var columnR = 18; // Column R (column number 18)
    
    if (range.getColumn() === columnR && range.getRow() > 1) { // Check for column Q and exclude the header row
      var value = range.getValue();
      var adjacentCell = sheet.getRange(range.getRow(), range.getColumn() + 1); // Get the adjacent cell
      
      if (value === "Done" || value === "Hold") {
        var currentDate = Utilities.formatDate(new Date(), "IST", "dd/MM/yyyy HH:mm:ss");
        adjacentCell.setValue(currentDate);
      } else {
        adjacentCell.clearContent();
      }
    }
  }
}


//Support Done Timestamp
function supportTimestamp() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getActiveRange();
  
  if (sheet.getName() === "Support Ticket") {
    var columnR = 18; // Column R (column number 18)
    
    if (range.getColumn() === columnR && range.getRow() > 1) { // Check for column N and exclude the header row
      var value = range.getValue();
      var adjacentCell = sheet.getRange(range.getRow(), range.getColumn() + 1); // Get the adjacent cell
      
      if (value === "Done" || value === "Hold") {
        var currentDate = Utilities.formatDate(new Date(), "IST", "dd/MM/yyyy HH:mm:ss");
        adjacentCell.setValue(currentDate);
      } else {
        adjacentCell.clearContent();
      }
    }
  }
}






// Archives Support Ticket Sheet
function archiveSupport() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName("Support Ticket");
  var archiveSheet = spreadsheet.getSheetByName("Archive - SP");

  var dataRange = sourceSheet.getDataRange();
  var lastColumnIndex = dataRange.getLastColumn();
  var conditionColumnIndex = 21; // Column U (1 based index)

  var sourceHeaders = sourceSheet.getRange(4, 1, 1, lastColumnIndex).getValues()[0];
  var archiveHeaders = archiveSheet.getRange(1, 1, 1, lastColumnIndex).getValues()[0];

  var lastHeader = sourceHeaders[sourceHeaders.length - 1];
  if (lastHeader !== "Archive Helper") {
    var formulas = [
      '=ArrayFormula(IF(ROW(A4:A)=4,"Year",IF(A4:A="",,TEXT(A4:A,"yyy")*1)))',
      '=ArrayFormula(IF(ROW(A4:A)=4,"Qtr",IF(A4:A="",,IF(CEILING(TEXT(A4:A,"m")*1,3)/3=1,"Q4",IF(CEILING(TEXT(A4:A,"m")*1,3)/3=2,"Q1",IF(CEILING(TEXT(A4:A,"m")*1,3)/3=3,"Q2","Q3"))))))',
      '=ArrayFormula(IF(ROW(A4:A)=4,"Year - Qtr",IF(Z4:Z="",,Z4:Z&" - "&AA4:AA)))',
      '=ArrayFormula(IF(A4:A="",,IF(ROW(A4:A)=4,"Financial Year",IF((--(MONTH(A4:A)>=1))*(--(MONTH(A4:A)<=3)),YEAR(A4:A)-1&"-"&YEAR(A4:A),YEAR(A4:A)&"-"&YEAR(A4:A)+1))))',
      '={"Archive Helper";"Yes"}'
    ];

    var headerRow = 4; // Modify to row 4 for header formulas
    var headerRange = sourceSheet.getRange(headerRow, lastColumnIndex + 1, 1, formulas.length);
    headerRange.setFormulas([formulas]);
  }

  for (var i = 0; i < sourceHeaders.length; i++) {
    if (sourceHeaders[i] !== archiveHeaders[i]) {
      var fileName = spreadsheet.getName();
      var subject = fileName + ' | Headers do not match for Archive';
      var body = 'Headers do not match for ' + fileName + ' for running the Archive function. Do the needful.';
      MailApp.sendEmail({
        to: 'central.data@arihant.com',
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
  for (var row = sourceData.length - 1; row >= 4; row--) { // Start processing from row 5 onwards
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
  formulaCell.setFormula('=IFNA(FILTER(INDIRECT("\'Support Ticket\'!A5:" & LEFT(ADDRESS(1, COUNTA(\'Support Ticket\'!4:4), 4), LEN(ADDRESS(1, COUNTA(\'Support Ticket\'!4:4), 4)) - 1)), INDIRECT("\'Support Ticket\'!A5:A") <> ""))');
}





// Archives Process Creation Sheet
function archiveProcess() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName("Process Creation");
  var archiveSheet = spreadsheet.getSheetByName("Archive - PC");

  var dataRange = sourceSheet.getDataRange();
  var lastColumnIndex = dataRange.getLastColumn();
  var conditionColumnIndex = 27; // Column AA (1 based index)

  var sourceHeaders = sourceSheet.getRange(4, 1, 1, lastColumnIndex).getValues()[0];
  var archiveHeaders = archiveSheet.getRange(1, 1, 1, lastColumnIndex).getValues()[0];

  var lastHeader = sourceHeaders[sourceHeaders.length - 1];
  if (lastHeader !== "Archive Helper") {
    var formulas = [
      '=ArrayFormula(IF(ROW(A4:A)=4,"Year",IF(A4:A="",,TEXT(A4:A,"yyy")*1)))',
      '=ArrayFormula(IF(ROW(A4:A)=4,"Qtr",IF(A4:A="",,IF(CEILING(TEXT(A4:A,"m")*1,3)/3=1,"Q4",IF(CEILING(TEXT(A4:A,"m")*1,3)/3=2,"Q1",IF(CEILING(TEXT(A4:A,"m")*1,3)/3=3,"Q2","Q3"))))))',
      '=ArrayFormula(IF(ROW(A4:A)=4,"Year - Qtr",IF(Z4:Z="",,Z4:Z&" - "&AA4:AA)))',
      '=ArrayFormula(IF(A4:A="",,IF(ROW(A4:A)=4,"Financial Year",IF((--(MONTH(A4:A)>=1))*(--(MONTH(A4:A)<=3)),YEAR(A4:A)-1&"-"&YEAR(A4:A),YEAR(A4:A)&"-"&YEAR(A4:A)+1))))',
      '={"Archive Helper";"Yes"}'
    ];

    var headerRow = 4; // Modify to row 4 for header formulas
    var headerRange = sourceSheet.getRange(headerRow, lastColumnIndex + 1, 1, formulas.length);
    headerRange.setFormulas([formulas]);
  }

  for (var i = 0; i < sourceHeaders.length; i++) {
    if (sourceHeaders[i] !== archiveHeaders[i]) {
      var fileName = spreadsheet.getName();
      var subject = fileName + ' | Headers do not match for Archive';
      var body = 'Headers do not match for ' + fileName + ' for running the Archive function. Do the needful.';
      MailApp.sendEmail({
        to: 'central.data@arihant.com',
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
  for (var row = sourceData.length - 1; row >= 4; row--) { // Start processing from row 5 onwards
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
  formulaCell.setFormula('=IFNA(FILTER(INDIRECT("\'Process Creation\'!A5:" & LEFT(ADDRESS(1, COUNTA(\'Process Creation\'!4:4), 4), LEN(ADDRESS(1, COUNTA(\'Process Creation\'!4:4), 4)) - 1)), INDIRECT("\'Process Creation\'!A5:A") <> ""))');
}






// Archives Helpdesk Form
function archiveHelpdeskForm() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName("Helpdesk Form");
  var archiveSheet = spreadsheet.getSheetByName("Archive - Form");

  var dataRange = sourceSheet.getDataRange();
  var lastColumnIndex = dataRange.getLastColumn();
  var conditionColumnIndex = 30; // Column AD (1 based index)

  var sourceHeaders = sourceSheet.getRange(1, 1, 1, lastColumnIndex).getValues()[0];
  var archiveHeaders = archiveSheet.getRange(1, 1, 1, lastColumnIndex).getValues()[0];

  var lastHeader = sourceHeaders[sourceHeaders.length - 1];
  if (lastHeader !== "Archive Helper") {
    var formulas = [
      '=ArrayFormula(IF(ROW(A1:A)=1,"Year",IF(A1:A="",,TEXT(A1:A,"yyy")*1)))',
      '=ArrayFormula(IF(ROW(A1:A)=1,"Qtr",IF(A1:A="",,IF(CEILING(TEXT(A1:A,"m")*1,3)/3=1,"Q4",IF(CEILING(TEXT(A1:A,"m")*1,3)/3=2,"Q1",IF(CEILING(TEXT(A1:A,"m")*1,3)/3=3,"Q2","Q3"))))))',
      '=ArrayFormula(IF(ROW(A1:A)=1,"Year - Qtr",IF(Z4:Z="",,Z4:Z&" - "&AA1:AA)))',
      '=ArrayFormula(IF(A1:A="",,IF(ROW(A1:A)=1,"Financial Year",IF((--(MONTH(A1:A)>=1))*(--(MONTH(A1:A)<=3)),YEAR(A1:A)-1&"-"&YEAR(A1:A),YEAR(A1:A)&"-"&YEAR(A1:A)+1))))',
      '={"Archive Helper";"Yes"}'
    ];

    var headerRow = 1; // Modify to row 1 for header formulas
    var headerRange = sourceSheet.getRange(headerRow, lastColumnIndex + 1, 1, formulas.length);
    headerRange.setFormulas([formulas]);
  }

  for (var i = 0; i < sourceHeaders.length; i++) {
    if (sourceHeaders[i] !== archiveHeaders[i]) {
      var fileName = spreadsheet.getName();
      var subject = fileName + ' | Headers do not match for Archive';
      var body = 'Headers do not match for ' + fileName + ' for running the Archive function. Do the needful.';
      MailApp.sendEmail({
        to: 'central.data@arihant.com',
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
  formulaCell.setFormula('=IFNA(FILTER(INDIRECT("\'Helpdesk Form\'!A2:" & LEFT(ADDRESS(1, COUNTA(\'Helpdesk Form\'!1:1), 4), LEN(ADDRESS(1, COUNTA(\'Helpdesk Form\'!1:1), 4)) - 1)), INDIRECT("\'Helpdesk Form\'!A2:A") <> ""))');
}




//Auto create DME triggers in the sheet
function createTriggersforDME() {
  const functionsAndTriggers = [
    { name: 'updateEmployeeDB', type: ScriptApp.EventType.CLOCK, hour: 0 },
    { name: 'generateSrNumbers', type: ScriptApp.EventType.ON_FORM_SUBMIT },
    { name: 'changeFont', type: ScriptApp.EventType.ON_FORM_SUBMIT },
    { name: 'copyData', type: ScriptApp.EventType.CLOCK, interval: 5 },
    { name: 'shorten', type: ScriptApp.EventType.CLOCK, interval: 5 },
    { name: 'processTimestamp', type: ScriptApp.EventType.ON_EDIT },
    { name: 'supportTimestamp', type: ScriptApp.EventType.ON_EDIT },
    { name: 'archiveHelpdeskForm', type: ScriptApp.EventType.WEEK_DAY, day: ScriptApp.WeekDay.MONDAY, hour: 21 },
    { name: 'archiveProcess', type: ScriptApp.EventType.WEEK_DAY, day: ScriptApp.WeekDay.MONDAY, hour: 22 },
    { name: 'archiveSupport', type: ScriptApp.EventType.WEEK_DAY, day: ScriptApp.WeekDay.MONDAY, hour: 23 }
  ];

  const timezone = 'Asia/Kolkata'; // Indian Standard Time (IST)

  functionsAndTriggers.forEach(triggerInfo => {
    const { name, type, interval, day, hour = 0 } = triggerInfo;
    let trigger;

    if (type === ScriptApp.EventType.CLOCK) {
      if (interval) {
        trigger = ScriptApp.newTrigger(name)
          .timeBased()
          .everyMinutes(interval)
          .create();
      } else {
        trigger = ScriptApp.newTrigger(name)
          .timeBased()
          .everyDays(1)
          .atHour(hour)
          .create();
      }
    } else if (type === ScriptApp.EventType.ON_FORM_SUBMIT) {
      trigger = ScriptApp.newTrigger(name)
        .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
        .onFormSubmit()
        .create();
    } else if (type === ScriptApp.EventType.ON_EDIT) {
      trigger = ScriptApp.newTrigger(name)
        .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
        .onEdit()
        .create();
    } else if (type === ScriptApp.EventType.WEEK_DAY) {
      trigger = ScriptApp.newTrigger(name)
        .timeBased()
        .onWeekDay(day)
        .atHour(hour)
        .create();
    }
  });

  // Display success message
  Browser.msgBox('DME Triggers were created successfully');
}
