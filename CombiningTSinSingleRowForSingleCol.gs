function CombiningTSinSingleRow() {
  //This is being used for Sales CRM project for finding days between sales Stage

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Fetch source data from the "Inquiries" sheet (read-only)
  const logSheetUrl = "https://docs.google.com/spreadsheets/d/1T06aDlRWzga74ifp34OQnJE4Ti2DiNMrsl_eWVwb1Nk/edit?gid=891129541#gid=891129541";
  const inquiriesSheet = SpreadsheetApp.openByUrl(logSheetUrl).getSheetByName("Inquiries");
  const inquiriesData = inquiriesSheet.getDataRange().getValues(); // Source data is fetched but not altered

  // Get data from the current sheet (Column A values)
  const lastRow = sheet.getLastRow();
  const columnAValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat(); // Read column A, excluding header

  // Process the data and format timestamps
  const output = columnAValues.map(value => {
    if (!value) {
      return [""]; // Handle empty rows
    }

    // Filter the source data for matching rows
    const filtered = inquiriesData
      .filter(row => row[0] === value && row[10] !== "Fresh Inquiry")
      .map(row => {
        // Format the timestamp (row[1]) to "DD/MM/YYYY HH:MM:SS"
        const date = typeof row[1] === "number" ? new Date(row[1]) : new Date(row[1]);
        return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      });

    // Concatenate all matches into a single string separated by commas
    return filtered.length > 0 ? [filtered.join(", ")] : [""];
  });

  // Write the processed data to the active sheet (starting from D2)
  const outputRange = sheet.getRange(2, 4, output.length, 1); // Single column
  outputRange.clear(); // Clear existing data
  outputRange.setNumberFormat("@STRING@"); // Ensure plain text format
  outputRange.setValues(output); // Write new data
}
