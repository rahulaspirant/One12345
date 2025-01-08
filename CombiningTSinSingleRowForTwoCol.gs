function CombiningTSAndTextInSingleRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('File');
  
  // Fetch source data from the "Inquiries" sheet (read-only)
  const logSheetUrl = "https://docs.google.com/spreadsheets/d/1T06aDlRWzga74ifp34OQnJE4Ti2DiNMrsl_eWVwb1Nk/edit?gid=891129541#gid=891129541";
  const inquiriesSheet = SpreadsheetApp.openByUrl(logSheetUrl).getSheetByName("Inquiries");
  const inquiriesData = inquiriesSheet.getDataRange().getValues(); // Source data is fetched but not altered

  // Get data from the current sheet (Column A values)
  const lastRow = sheet.getLastRow();
  const columnAValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat(); // Read column A, excluding header

  // Process the data to create separate outputs for timestamps and text
  const outputTimestamps = [];
  const outputTexts = [];

  columnAValues.forEach(value => {
    if (!value) {
      // Handle empty rows
      outputTimestamps.push([""]);
      outputTexts.push([""]);
      return;
    }

    // Filter the source data for matching rows
    const filteredData = inquiriesData.filter(row => row[0] === value);

    // Format timestamps from column B (row[1])
    const formattedTimestamps = filteredData.map(row => {
      const date = typeof row[1] === "number" ? new Date(row[1]) : new Date(row[1]);
      return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    });

    // Collect text from column K (row[10])
    const combinedText = filteredData.map(row => row[10]);

    // Push the results to respective outputs
    outputTimestamps.push([formattedTimestamps.join(", ") || ""]);
    outputTexts.push([combinedText.join(", ") || ""]);
  });

  // Write timestamps to column D (starting from D2)
  const timestampRange = sheet.getRange(2, 4, outputTimestamps.length, 1); // Single column
  timestampRange.clear(); // Clear existing data
  timestampRange.setNumberFormat("@STRING@"); // Ensure plain text format
  timestampRange.setValues(outputTimestamps);

  // Write combined text to column E (starting from E2)
  const textRange = sheet.getRange(2, 5, outputTexts.length, 1); // Single column
  textRange.clear(); // Clear existing data
  textRange.setNumberFormat("@STRING@"); // Ensure plain text format
  textRange.setValues(outputTexts);
}
