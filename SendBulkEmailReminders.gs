function SendBulkEmailReminders() {
  //This is stopped as of now
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName("Data");
  var lastRow = dataSheet.getLastRow();
  var dataRange = dataSheet.getRange("A1:AC" + lastRow);
  var dataValues = dataRange.getValues();
  var headers = dataValues[0];
  
  // Calculate yesterday's date
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  var formattedDate = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "dd-MM-yyyy");
  
  // Group data by sales executives
  var groupedData = {};
  for (var i = 1; i < dataValues.length; i++) {
    var row = dataValues[i];
    var salesExecEmail = row[6]; // Assuming sales executives' emails are in column G
    var execName = row[5]; // Assuming sales executives' names are in column F
    var condition1 = row[17]; // Archive Condition)
    var condition2 = row[2]instanceof Date ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), "dd-MM-yyy") : ''; //Timestamp
    var condition3 = row[8]; // Assuming column D is 4th column
    var condition4 = row[9]; // Assuming column I is 9th column
    var condition5 = row[10]instanceof Date ? Utilities.formatDate(row[10], Session.getScriptTimeZone(), "dd-MM-yyy") : ''; //Date of Visit
    var condition6 = row[11]instanceof Date ? Utilities.formatDate(row[11], Session.getScriptTimeZone(), "dd-MM-yyy") : ''; //Next follow up
    var condition7 = row[12]; // Assuming column U is 21st column
    var condition8 = row[14]; // Assuming column Z is 26th column
    var condition9 = '<a href="' + row[27] + '">Edit Form</a>'; // Form Edit
    
    if (condition1 == "" ) {
      if (!groupedData[salesExecEmail]) {
        groupedData[salesExecEmail] = { name: execName, data: [] };
      }
      groupedData[salesExecEmail].data.push([condition2, condition3, condition4, condition5,condition6,condition7,condition8,condition9]);
    }
  }
  
  // Sending reminder emails
  for (var execEmail in groupedData) {
    var execName = groupedData[execEmail].name;
    var execData = groupedData[execEmail].data;
    var emailBody = "<html><body><p>Dear " + execName + ",</p>";
    emailBody += "<p>Here are your pending leads for " + formattedDate + ":</p>";
    emailBody += "<table border='1'><tr>";
    
    
    emailBody += "<th>" + headers[2] + "</th><th>" + headers[8] + "</th><th>" + headers[9] + "</th><th>" + headers[10] + "</th><th>" + headers[11] + "</th><th>" + headers[12] + "</th><th>" + headers[14] + "</th><th>" + headers[27] + "</th></tr>";

    
    
    for (var j = 0; j < execData.length; j++) {
      var rowData = execData[j];
      emailBody += "<tr>";
      for (var k = 0; k < rowData.length; k++) {
        emailBody += "<td>" + rowData[k] + "</td>";
      }
      emailBody += "</tr>";
    }
    emailBody += "</table></body></html>";
    
    // Send email
    MailApp.sendEmail({
      to: execEmail,
    //  cc: "rahulsolanki0045@gmail.com", // Add the email address(es) you want to include in CC
      subject: "Pending Leads till " + execName + " For - " + formattedDate,
      htmlBody: emailBody
});

  }
}
