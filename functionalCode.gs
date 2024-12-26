function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  
 // if (sheetName == "Layout Form" && e.range.getColumn() == 22 && e.range.getRow() > 1) {   //Index is 1 based....this
  if (sheetName == "Layout Form" && e.range.getColumn() == 23 && e.range.getRow() > 1) {   //Index is 1 based.....this
    var value = e.value;
    var row = e.range.getRow();
    
    
    if (value == "Done") {
      //Dropdown -  Done,Incomplete Input,Inquiry Dead,Hold
      var timestamp = new Date();
     // sheet.getRange(row, 23).setValue(timestamp);
      sheet.getRange(row, 24).setValue(timestamp); //...........this
      sheet.getRange(row, 24).setNumberFormat("dd/MM/yyyy HH:mm:ss");   //.....this
    }
  }
}


/////////////////////////////////////////////////////////////////////////////////////////////


function registerNewEditResponseURLTrigger() {
  // check if an existing trigger is set
  var existingTriggerId = PropertiesService.getUserProperties().getProperty('onFormSubmitTriggerID')
  if (existingTriggerId) {
    var foundExistingTrigger = false
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
      if (trigger.getUniqueId() === existingTriggerId) {
        foundExistingTrigger = true
      }
    })
    if (foundExistingTrigger) {
      return
    }
  }
var trigger = ScriptApp.newTrigger('onFormSubmitEvent')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onFormSubmit()
    .create()
PropertiesService.getUserProperties().setProperty('onFormSubmitTriggerID', trigger.getUniqueId())
}
function getTimestampColumn(sheet) {
  for (var i = 1; i <= sheet.getLastColumn(); i += 1) {
    if (sheet.getRange(1, i).getValue() === 'Timestamp') {
      return i
    }
  }
  return 1
}
function getFormResponseEditUrlColumn(sheet) {
  var form = FormApp.openByUrl(sheet.getFormUrl())
  for (var i = 1; i <= sheet.getLastColumn(); i += 1) {
    if (sheet.getRange(1, i).getValue() === 'Form Response Edit URL') {
      return i
    }
  }
  // get the last column at which the url can be placed.
  return Math.max(sheet.getLastColumn() + 1, form.getItems().length + 2)
}
/**
 * params: { sheet, form, formResponse, row }
 */
function addEditResponseURLToSheet(params) {
  if (!params.col) {
    params.col = getFormResponseEditUrlColumn(params.sheet)
  }
  var formResponseEditUrlRange = params.sheet.getRange(params.row, params.col)
  formResponseEditUrlRange.setValue(params.formResponse.getEditResponseUrl())
}
function onOpen() {
  var menu = [{ name: 'Add Form Edit Response URLs', functionName: 'setupFormEditResponseURLs' }]
  SpreadsheetApp.getActive().addMenu('Forms', menu)
}


/////////////////////////////////////////////////////////////////////////////////////////////



function setupFormEditResponseURLs() {
  var sheet = SpreadsheetApp.getActiveSheet()
  var spreadsheet = SpreadsheetApp.getActive()
  var formURL = sheet.getFormUrl()
  if (!formURL) {
    SpreadsheetApp.getUi().alert('No Google Form associated with this sheet. Please connect it from your Form.')
    return
  }
  var form = FormApp.openByUrl(formURL)
// setup the header if not existed
  var headerFormEditResponse = sheet.getRange(1, getFormResponseEditUrlColumn(sheet))
  var title = headerFormEditResponse.getValue()
  if (!title) {
    headerFormEditResponse.setValue('Form Response Edit URL')
  }
var timestampColumn = getTimestampColumn(sheet)
  var editResponseUrlColumn = getFormResponseEditUrlColumn(sheet)
  
  var timestampRange = sheet.getRange(2, timestampColumn, sheet.getLastRow() - 1, 1)
  var editResponseUrlRange = sheet.getRange(2, editResponseUrlColumn, sheet.getLastRow() - 1, 1)
  if (editResponseUrlRange) {
    var editResponseUrlValues = editResponseUrlRange.getValues()
    var timestampValues = timestampRange.getValues()
    for (var i = 0; i < editResponseUrlValues.length; i += 1) {
      var editResponseUrlValue = editResponseUrlValues[i][0]
      var timestampValue = timestampValues[i][0]
      if (editResponseUrlValue === '') {
        var timestamp = new Date(timestampValue)
        if (timestamp) {
          var formResponse = form.getResponses(timestamp)[0]
          editResponseUrlValues[i][0] = formResponse.getEditResponseUrl()
          var row = i + 2
          if (row % 10 === 0) {
            spreadsheet.toast('processing rows ' + row + ' to ' + (row + 10))
            editResponseUrlRange.setValues(editResponseUrlValues)
            SpreadsheetApp.flush()
          }
        }
      }
    }
    
    editResponseUrlRange.setValues(editResponseUrlValues)
    SpreadsheetApp.flush()
  }
registerNewEditResponseURLTrigger()
  SpreadsheetApp.getUi().alert('You are all set! Please check the Form Response Edit URL column in this sheet. Future responses will automatically sync the form response edit url.')
}
function onFormSubmitEvent(e) {
  var sheet = e.range.getSheet()
  var form = FormApp.openByUrl(sheet.getFormUrl())
  var formResponse = form.getResponses().pop()
  addEditResponseURLToSheet({
    sheet: sheet,
    form: form,
    formResponse: formResponse,
    row: e.range.getRow(),
  })
}

/////////////////////////////////////////////////////////////////////////////////////////////

function sendIncompleteMail(e) {
  var sheetName = "Layout Form";
  var columnName = "Status";
  var emailColumnNumber = 2;
  var incompleteInputValue = "Incomplete Input";
  var inquiryDeadValue = "Inquiry Dead";
  var sentStatusValue = "Sent";

  var salesExecutiveNameColumn = 2;
  var designRemarksColumn = 26;
  var assignedToColumn = 21;
  var formResponseEditURLColumn = 31;
  var inquiryNoColumn = 3;
  var hyperlinkColumn = 31;      //This needs to be checked
  var sentStatusColumn = 27;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var dataRange = sheet.getRange(2, getColumnIndexByName(sheet, columnName), sheet.getLastRow() - 1, 1);
  var dataValues = dataRange.getValues();
  var emails = [];

  for (var i = 0; i < dataValues.length; i++) {
    var statusValue = dataValues[i][0];
    var sentStatus = sheet.getRange(i + 2, sentStatusColumn).getValue();
    if (statusValue === incompleteInputValue && sentStatus !== sentStatusValue) {
      var rowData = sheet.getRange(i + 2, 1, 1, sheet.getLastColumn()).getValues()[0];
      var salesExecutiveName = rowData[salesExecutiveNameColumn - 1];
      var designRemarks = rowData[designRemarksColumn - 1];
      var assignedTo = rowData[assignedToColumn - 1];
      var formResponseEditURL = rowData[formResponseEditURLColumn - 1]; // Fetch formResponseEditURL for each row
      var inquiryNo = rowData[inquiryNoColumn - 1];
      var email = rowData[emailColumnNumber - 1];
      var hyperlink = rowData[hyperlinkColumn - 1];

      var emailBody = generateIncompleteEmailBody(salesExecutiveName, designRemarks, assignedTo, formResponseEditURL, hyperlink);
      var subject = "Layout info required for IQ No. " + inquiryNo;

      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: emailBody,
      });

      sheet.getRange(i + 2, sentStatusColumn).setValue(sentStatusValue);
    } else if (statusValue === inquiryDeadValue && sentStatus !== sentStatusValue) {
      var rowData = sheet.getRange(i + 2, 1, 1, sheet.getLastColumn()).getValues()[0];
      var salesExecutiveName = rowData[salesExecutiveNameColumn - 1];
      var assignedTo = rowData[assignedToColumn - 1];
      var inquiryNo = rowData[inquiryNoColumn - 1];
      var email = rowData[emailColumnNumber - 1];
      var formResponseEditURL = rowData[formResponseEditURLColumn - 1]; // Fetch formResponseEditURL for each row
      var hyperlink = rowData[hyperlinkColumn - 1];

      var emailBody = generateInquiryDeadEmailBody(salesExecutiveName, assignedTo, inquiryNo, hyperlink, formResponseEditURL);
      var subject = "Inquiry Dead for IQ No. " + inquiryNo;

      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: emailBody,
      });

      sheet.getRange(i + 2, sentStatusColumn).setValue(sentStatusValue);
    }
  }
}


function generateIncompleteEmailBody(salesExecutiveName, designRemarks, assignedTo, formResponseEditURL, hyperlink) {
  var body = "Hi <b>" + salesExecutiveName + "</b>,<br><br>" +
    "The Layout request filled by you has some details missing.<br>" +
    "Following are the remarks:<br><br>" +
    "<b>" + designRemarks + "</b><br><br>" +
    "To provide the missing information in the same form, please " +
    "<b><a href='" + formResponseEditURL + "'>CLICK HERE</a></b>.<br><br>" +
    "Best Regards,<br>" +
    assignedTo + ".";

    //If ticket is not updated in 10 days then it will be considered as Inquiry dead.

  return body;
}

function generateInquiryDeadEmailBody(salesExecutiveName, assignedTo, inquiryNo, hyperlink, formResponseEditURL) {
  var body = "Hi <b>" + salesExecutiveName + "</b>,<br><br>" +
    "The inquiry with IQ No. " + inquiryNo + " is considered dead.<br>" +
    "To keep this Inquiry alive and provide the missing information, please " +
    "<b><a href='" + formResponseEditURL + "'>CLICK HERE</a></b>.<br><br>" +
    "Best Regards,<br>" +
    assignedTo + ".";

  return body;
}


////////////////////////////////////////////////////////////////////////////////////////


function getColumnIndexByName(sheet, columnName) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === columnName) {
      return i + 1;
    }
  }
  throw new Error("Column '" + columnName + "' not found in the sheet.");
}







// Not in use
function originalTimestamp() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Layout Form');
  var lastRow = sheet.getLastRow();
  var columnAD = sheet.getRange('AD2:AD' + lastRow);
  var columnA = sheet.getRange('A2:A' + lastRow);
  
  var columnADValues = columnAD.getValues();
  var columnAValues = columnA.getValues();
  
  for (var i = 0; i < columnADValues.length; i++) {
    if (columnADValues[i][0] === '') {
      columnADValues[i][0] = columnAValues[i][0];
    }
  }
  
  columnAD.setValues(columnADValues);
  columnAD.copyTo(columnAD, {contentsOnly: true});
}



function splitHyperlinks() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Layout Form');
  var lastRow = sheet.getLastRow();
  
  for (var i = 1; i <= lastRow; i++) {
    var valueInN = sheet.getRange(i, 14).getValue(); // Checking if Column N has data
    var valueInJ = sheet.getRange(i, 10).getValue(); // Getting value from Column J
    
    if (valueInN === "" && valueInJ !== "") {
      var hyperlinks = valueInJ.split(',');
      for (var j = 0; j < hyperlinks.length; j++) {
        sheet.getRange(i, 14 + j).setFormula('=HYPERLINK("' + hyperlinks[j] + '", "Click Here")');
      }
    }
  }
}


/////////////////////////////////////////////////////////////////////////////////////////////

function checkTimestamp() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Layout Form");
  var data = sheet.getDataRange().getValues();

  var timestampColumn = 1;
  var inquiryNoColumn = 3;
  var emailColumn = 33; // Column AG
  var assignedToColumn = 21;
  var sentTimestampColumn = 34; // Column AH

  var updatedRows = [];

  var timestampFormat = "dd/MM/yyyy HH:mm:ss";
  var timestampFormatCell = sheet.getRange(1, sentTimestampColumn);
  timestampFormatCell.setNumberFormat(timestampFormat);

  for (var i = 1; i < data.length; i++) { // Start from row 2 (excluding header row)
    var timestamp = data[i][timestampColumn - 1];
    var inquiryNo = data[i][inquiryNoColumn - 1];
    var email = data[i][emailColumn - 1];
    var assignedTo = data[i][assignedToColumn - 1];
    var sentTimestamp = data[i][sentTimestampColumn - 1];
    var agTimestamp = data[i][sentTimestampColumn - 1];

    if (agTimestamp === "") {
      sheet.getRange(i + 1, timestampColumn).copyTo(sheet.getRange(i + 1, sentTimestampColumn), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      sheet.getRange(i + 1, sentTimestampColumn).setNumberFormat(timestampFormat);
      agTimestamp = data[i][sentTimestampColumn - 1]; // Update the agTimestamp variable with the newly copied value
    }

    if (agTimestamp !== "" && new Date(agTimestamp) < new Date(timestamp)) {
      var subject = "Inquiry No. " + inquiryNo + " Updated";
      var message = "Dear " + assignedTo + ",\n\n" +
        "The form response for Inquiry No. " + inquiryNo + " has been updated in the sheet.\n\n" +
        "Please do the needful.\n\n" +
        "Regards,\n" +
        "Parag Mistry";

      MailApp.sendEmail(email, subject, message);

      sheet.getRange(i + 1, sentTimestampColumn).setValue(new Date());
      sheet.getRange(i + 1, sentTimestampColumn).setNumberFormat(timestampFormat);

      updatedRows.push(i + 1);
    }
  }

  if (updatedRows.length > 0) {
    Logger.log("Emails sent for rows: " + updatedRows.join(", "));
  }
}



/////////////////////////////////////////////////////////////////////////////////////////////



function archiveData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName("Layout Form");
  var archiveSheet = spreadsheet.getSheetByName("Archive");

  var sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  var archiveHeaders = archiveSheet.getRange(1, 1, 1, archiveSheet.getLastColumn()).getValues()[0];

  var lastHeader = sourceHeaders[sourceHeaders.length - 1];
  if (lastHeader !== "Archive Helper") {
    var formulas = [
      '=ArrayFormula(IF(ROW(A:A)=1,"Year",IF(A1:A="",,TEXT(A1:A,"yyy")*1)))',
      '=ArrayFormula(IF(ROW(A:A)=1,"Qtr",IF(A1:A="",,IF(CEILING(TEXT(A1:A,"m")*1,3)/3=1,"Q4",IF(CEILING(TEXT(A1:A,"m")*1,3)/3=2,"Q1",IF(CEILING(TEXT(A1:A,"m")*1,3)/3=3,"Q2","Q3"))))))',
      '=ArrayFormula(IF(ROW(A:A)=1,"Year - Qtr",IF(AL1:AL="",,AL1:AL&" - "&AM1:AM)))',
      '={"Archive Helper";"Yes"}'
    ];

    var headerRow = 1;
    var headerRange = sourceSheet.getRange(headerRow, sourceHeaders.length + 1, 1, formulas.length);
    headerRange.setFormulas([formulas]);

    // Update sourceHeaders array with the new formulas
    sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
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
    var lastColumnValue = archiveData[row][archiveData[row].length - 1];
    if (lastColumnValue === "Yes") {
      archiveSheet.deleteRow(row + 1);
    }
  }

  var sourceData = sourceSheet.getDataRange().getValues();
  var hasDataInColumnAE = false; // Flag to track if there is any data in column AE
  for (var row = sourceData.length - 1; row > 0; row--) {
    var conditionValue = sourceData[row][21];   //V or Status column
    if (conditionValue === "Inquiry Dead" || conditionValue === "Done") {
      //Incomplete Inquiry must be archived after 7 days
      archiveSheet.appendRow(sourceData[row]);
      sourceSheet.deleteRow(row + 1);
      hasDataInColumnAE = true; // Set the flag to true if there is data
    }
  }

  var lastNonBlankRow = archiveSheet.getLastRow() + 1;
  if (!hasDataInColumnAE) {
    lastNonBlankRow = Math.max(lastNonBlankRow, 2); // Start from row 2 if there is no data in column AE
  }

  var formulaCell = archiveSheet.getRange(lastNonBlankRow, 1);
  formulaCell.setFormula('=IFNA(FILTER(INDIRECT("\'Layout Form\'!A2:"&LEFT(ADDRESS(1,COUNTA(\'Layout Form\'!1:1),4),LEN(ADDRESS(1,COUNTA(\'Layout Form\'!1:1),4))-1)),INDIRECT("\'Layout Form\'!A2:A")<>""))');
}








