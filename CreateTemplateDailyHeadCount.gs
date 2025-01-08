function appendDailyData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName("Template"); // Replace with your template sheet name
  const mainSheet = ss.getSheetByName("Daily Data"); // Replace with your main data sheet name
  
  const today = new Date(); // Automatically use today's date
  const formattedDate = Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
  
  // Get the template rows
  const templateData = templateSheet.getRange(4, 1, 19, templateSheet.getLastColumn()).getValues();
  
  // Replace the placeholder date
  const updatedData = templateData.map(row => {
    row[0] = formattedDate; // Assuming Date is in the first column
    return row;
  });
  
  // Append updated rows to the main sheet
  mainSheet.getRange(mainSheet.getLastRow() + 1, 1, updatedData.length, updatedData[0].length).setValues(updatedData);
}
