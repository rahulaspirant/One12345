function prefillColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet(); // Use your target sheet name
  const range = sheet.getRange(4, 6, 19, 2); // Adjust to match your "Department" and "Resp." columns
  
  const departments = [
    "LRTM Molding - APS", "Mold Maint. - APS", "Finishing - APS", "Steel Painting - APS",
    "HL - APS", "Post Production", "DND", "Re-Surfacing", "CNC", "T1-G1", "Acetone M/C",
    "Launcher", "Buffer", "Maintenance", "House Keeping", "Rokda", "Store", "Security",
    "Ramdas and Team (Special Approval)"
  ];
  
  const responsiblePersons = [
    "Person 1", "Person 2", "Person 3", "Person 4", // Update with actual names
    "Person 5", "Person 6", "Person 7", "Person 8", "Person 9", "Person 10",
    "Person 11", "Person 12", "Person 13", "Person 14", "Person 15", "Person 16",
    "Person 17", "Person 18", "Person 19"
  ];
  
  // Combine departments and responsible persons
  const data = departments.map((dept, i) => [dept, responsiblePersons[i]]);
  
  // Fill the range
  range.setValues(data);
}
