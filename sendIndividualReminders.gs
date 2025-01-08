function sendIndividualReminders() {
  
  // All index are zero based

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName("Data");
  var lastRow = dataSheet.getLastRow();
  var dataRange = dataSheet.getRange("A2:K" + lastRow); // Adjusted range to include headers and data rows
  var dataValues = dataRange.getValues();
  
  var headers = dataSheet.getRange("A2:K2").getValues()[0]; // Fetch headers from the second row

  // Calculate current time
  var formattedTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm");

  // Loop through the data rows
  for (var i = 1; i < dataValues.length; i++) {
    var row = dataValues[i];
    var delegateeEmail = row[1]; // Column B - Delegatee Email Id
    var percentageStatus = row[10]; // Column K - % Delegatee Status

    // Check if the percentage is 0%
    if (percentageStatus === 0 || percentageStatus === "0%") {
      // Prepare email body
      var emailBody = "<html><body>";
      emailBody += "<p>Dear " + row[2] + ",</p>"; // Emp Name (Column C)
      emailBody += "<p>This is a reminder regarding your meeting task:</p>";
      emailBody += "<table border='1'><tr>";

      // Add headers
      headers.forEach(function(header) {
        emailBody += "<th>" + header + "</th>";
      });
      emailBody += "</tr>";

      // Add row data
      emailBody += "<tr>";
      row.forEach(function(cell) {
        emailBody += "<td>" + cell + "</td>";
      });
      emailBody += "</tr>";
      emailBody += "</table>";
      emailBody += "<p>Please take the necessary action.</p>";
      emailBody += "</body></html>";

      // Send email
      MailApp.sendEmail({
        to: delegateeEmail,
        //cc: "umang.panchal@greatescape.co.in,manish.singh@greatescape.co.in", // Add CC recipients if required
        subject: "Meeting FMS: Action Required",
        htmlBody: emailBody
      });
    }
  }
}
