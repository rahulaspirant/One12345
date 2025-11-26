/**
 * MAIN FUNCTION
 * This is the entry point. It orchestrates the flow but delegates the work.
 */
function runMouldMasterSync() {
  // --- 1. CONFIGURATION ---
  const config = {
    sourceUrl: 'https://docs.google.com/spreadsheets/d/132P2Aegycp4bmptgmXCENHrBgvLQTCcWB5fSDbKwTS0/edit?gid=0#gid=0',
    targetUrl: 'https://docs.google.com/spreadsheets/d/132P2Aegycp4bmptgmXCENHrBgvLQTCcWB5fSDbKwTS0/edit?gid=656840740#gid=656840740',
    sourceSheetName: 'Source',
    targetSheetName: 'Target',
    headersToExtract: ["Mould Number", "Mould Name", "Mould Type", "Status as of 24-25"],
    formatDate: false
  };

  // --- 2. GET SHEETS ---
  const sourceSheet = getSheet(config.sourceUrl, config.sourceSheetName);
  const targetSheet = getSheet(config.targetUrl, config.targetSheetName);

  // --- 3. GET SOURCE DATA & HEADERS ---
  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getDisplayValues()[0];
  const sourceData = sourceSheet.getRange(1, 1, sourceSheet.getLastRow() - 2, sourceSheet.getLastColumn()).getDisplayValues();
  
  // Create a map (e.g., {"Mould Number": 0, "Status": 4}) to find columns by name easily
  const headerMap = createHeaderIndexMap(sourceHeaders);

  // --- 4. GET EXISTING IDs (To prevent duplicates) ---
  const existingMouldNums = getExistingMouldNumbers(targetSheet, 1, 4); // 1 = Column A, 4 = Start Row

  // --- 5. PROCESS & FILTER DATA ---
  // This function handles the complex logic of filtering "Available" and matching columns
  const newRows = processAndFilterData(
    sourceData, 
    headerMap, 
    existingMouldNums, 
    config.headersToExtract, 
    config.formatDate
  );

  // --- 6. WRITE DATA ---
  if (newRows.length > 0) {
    appendDataToSheet(targetSheet, newRows, 1); // 1 is the starting column in target
    Logger.log(`Success: Added ${newRows.length} new rows.`);
  } else {
    Logger.log("No new data to add.");
  }
}


/**
 * HELPER 1: Opens a specific sheet from a URL.
 * Why modular? Reused twice (source and target), keeps the main function clean.
 */
function getSheet(url, sheetName) {
  const ss = SpreadsheetApp.openByUrl(url);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found in spreadsheet.`);
  return sheet;
}


/**
 * HELPER 2: Creates a dictionary of Header Name -> Column Index.
 * Why modular? Essential for robust code. If columns move in the source, this adapts automatically.
 */
function createHeaderIndexMap(headers) {
  const map = {};
  headers.forEach((header, index) => {
    map[header.trim()] = index;
  });
  return map;
}


/**
 * HELPER 3: Fetches existing Mould Numbers from the target to create a "Block List".
 * Why modular? Separates the logic of "reading target" from "processing source".
 */
function getExistingMouldNumbers(sheet, columnIndex, startRow) {
  const lastRow = sheet.getLastRow();
  const existingSet = new Set();
  
  if (lastRow >= startRow) {
    const data = sheet.getRange(startRow, columnIndex, lastRow - startRow + 1).getValues();
    data.forEach(([val]) => {
      if (val && String(val).trim() !== "") {
        existingSet.add(String(val).trim());
      }
    });
  }
  return existingSet;
}


/**
 * HELPER 4: The Logic Core. Filters "Available", checks duplicates, and maps columns.
 * Why modular? This is the most complex part. Isolating it makes it easier to debug logic errors.
 */
function processAndFilterData(data, headerMap, existingIds, headersToExtract, formatDate) {
  const processedRows = [];

  // Indices of the columns we actually want to copy
  const targetIndices = headersToExtract.map(name => {
    if (!(name in headerMap)) throw new Error(`Header "${name}" missing in source.`);
    return headerMap[name];
  });

  const statusIndex = headerMap["Status as of 24-25"];
  const mouldNumIndex = headerMap["Mould Number"];

  // Loop through every row in source
  for (let row of data) {
    const status = row[statusIndex];
    const mouldNum = String(row[mouldNumIndex]).trim();

    // 1. Business Logic Filter: Must be "Available" and have a Mould Number
    if (status !== "Available" || mouldNum === "") continue;

    // 2. Duplicate Check: Skip if already in target
    if (existingIds.has(mouldNum)) continue;

    // 3. Build the new row based ONLY on headersToExtract
    const newRow = targetIndices.map((colIndex, i) => {
      let value = row[colIndex];
      const headerName = headersToExtract[i];

      // Date Formatting Logic (Only runs if enabled and matches specific header)
      if (formatDate && headerName === "Final Actual Delivery Date" && value) {
        value = formatMyDate(value);
      }
      return value;
    });

    processedRows.push(newRow);
  }

  return processedRows;
}


/**
 * HELPER 5: Writes data to the bottom of the sheet.
 * Why modular? Handles the math of calculating ranges so you don't have to.
 */
function appendDataToSheet(sheet, data, startColumn) {
  const lastRow = sheet.getLastRow();
  // setRange(row, col, numRows, numCols)
  sheet.getRange(lastRow + 1, startColumn, data.length, data[0].length).setValues(data);
}


/**
 * HELPER 6: Simple Date Formatter.
 * Why modular? Keeps the messy Date logic out of the main loop.
 */
function formatMyDate(dateValue) {
  const date = new Date(dateValue);
  if (isNaN(date)) return dateValue; // Return original if not a valid date
  
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}-${month}-${year}`;
}
