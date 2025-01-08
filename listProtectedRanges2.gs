function listProtectedRanges2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName("Protection Audit") || ss.insertSheet("Protection Audit");
  
  // Clear the audit sheet for a fresh start
  auditSheet.clear();
  auditSheet.appendRow(["Sheet Name", "Protected Range", "Editors", "Warning"]);

  // Initialize data structure for tracking overlaps
  const rangesData = [];

  // Loop through each sheet to get protected ranges
  ss.getSheets().forEach(sheet => {
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protections.forEach(protection => {
      try {
        const range = protection.getRange();
        const editors = protection.getEditors().map(editor => editor.getEmail()).join(", ");
        const rangeA1 = range.getA1Notation();
        const sheetName = sheet.getName();

        // Add to ranges data for overlap detection
        rangesData.push({ sheetName, rangeA1, editors, protection });

        // Append data to audit sheet
        auditSheet.appendRow([sheetName, rangeA1, editors, ""]);
      } catch (e) {
        // Log the error to the audit sheet or Logger
        auditSheet.appendRow([sheet.getName(), "N/A", "N/A", "Permission Error"]);
        Logger.log("Error accessing range or editors: " + e.message);
      }
    });
  });

  // Check for overlaps
  for (let i = 0; i < rangesData.length; i++) {
    for (let j = i + 1; j < rangesData.length; j++) {
      if (rangesData[i].sheetName === rangesData[j].sheetName &&
          rangesOverlap(rangesData[i].protection.getRange(), rangesData[j].protection.getRange())) {
        
        // Mark overlap warnings in the audit sheet
        auditSheet.getRange(i + 2, 4).setValue("Overlap with " + rangesData[j].rangeA1);
        auditSheet.getRange(j + 2, 4).setValue("Overlap with " + rangesData[i].rangeA1);
      }
    }
  }

  SpreadsheetApp.flush();
}
