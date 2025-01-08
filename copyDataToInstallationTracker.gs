function copyDataToInstallationTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const plannerSheet = ss.getSheetByName('Planner Form');
  const trackerSheet = ss.getSheetByName('Installation Tracker');
  
  const plannerData = plannerSheet.getDataRange().getValues();
  let trackerLastRow = trackerSheet.getLastRow();
  
  // Define column indices for data extraction and insertion
  const columnsToCopy = [2, 3, 4, 5, 6, 9, 12, 13, 14, 15]; // B, C, D, E, F, I, L, M, N, O (1-based index)
  const columnsToPaste = [3, 4, 5, 6, 7, 8, 9, 10, 16, 17]; // C, D, E, F, G, H, I, J, P, Q (1-based index)
  
  plannerData.forEach(function(row, index) {
    // Skip the header row
    if (index === 0) return;
    
    const copiedStatus = row[17 - 1]; // Column Q (16th column, 0-based index)
    
    // Check if 'Copied' is present in column Q (status column)
    if (copiedStatus === 'Copied') {
      return; // Skip this row if already copied
    }
    
    // Check if the row is eligible for copying (based on Column Q)
    if (row[16 - 1] !== 'Copied') {
      // Extract data from Planner Form row
      const rowDataToCopy = columnsToCopy.map(colIndex => row[colIndex - 1]);
      
      // Determine the next available row in Installation Tracker
      const nextRow = trackerLastRow + 1;
      
      // Paste the data into Installation Tracker
      columnsToPaste.forEach(function(colIndex, idx) {
        trackerSheet.getRange(nextRow, colIndex).setValue(rowDataToCopy[idx]);
      });
      
      // Mark 'Copied' in Planner Form column Q (16th column, 0-based index)
      plannerSheet.getRange(index + 1, 17).setValue('Copied');
      
      // Update trackerLastRow to reflect the new last row after data insertion
      trackerLastRow++;
    }
  });
}
