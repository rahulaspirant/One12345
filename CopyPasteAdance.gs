function fetchData() {
//Single flow
  // CONFIGURATION
  const copyMode = "Selected"; // Use "All" to copy all columns, or "Selected" to copy only specific columns
  const filterEnabled = true; // Set to false to disable data filtering
  const formatDateColumns = false; // ✅ Set to true to enable date column formatting

  const headersToExtract = [
    "Mould Number",
    "Mould Name",
    "Mould Type",
    "Status as of 24-25",
  ];

  const sourceSpreadsheetUrl = 'https://docs.google.com/spreadsheets/d/1t8uhbsbyB8_gcdNOp6w1sQYGEH4870y_H1lFmyJh5kk/edit?gid=952746983#gid=952746983';
  const targetSpreadsheetUrl = 'https://docs.google.com/spreadsheets/d/104f7pi1F2a-ndTyrfDUYvMWmS8CHAv-UpcdECwPBQRA/edit?gid=0#gid=0';
  const sourceSheetName = 'Mould Master File';
  const targetSheetName = 'Sheet1';
  const targetStartColumn = 1;
  const startRow = 4; // Target data starts from row 4

  // Open sheets
  const sourceSheet = SpreadsheetApp.openByUrl(sourceSpreadsheetUrl).getSheetByName(sourceSheetName);
  const targetSheet = SpreadsheetApp.openByUrl(targetSpreadsheetUrl).getSheetByName(targetSheetName);

  // Get source headers and data
  const headers = sourceSheet.getRange(2, 1, 1, sourceSheet.getLastColumn()).getDisplayValues()[0];
  const lastSourceRow = sourceSheet.getLastRow();
  const data = sourceSheet.getRange(3, 1, lastSourceRow - 2, sourceSheet.getLastColumn()).getDisplayValues();

  // Create header map
  const headerIndexMap = {};
  headers.forEach((header, index) => {
    headerIndexMap[header.trim()] = index;
  });

  // Determine columns to extract
  const columnIndices = (copyMode === "Selected")
    ? headersToExtract.map(h => {
        if (!(h in headerIndexMap)) throw new Error(`Header not found: ${h}`);
        return headerIndexMap[h];
      })
    : headers.map((_, i) => i);

  const targetHeaders = (copyMode === "Selected") ? headersToExtract : headers;

  // Apply optional filter
  const filteredData = filterEnabled
    ? data.filter(row => {
        const StatusCol = row[headerIndexMap["Status as of 24-25"]];
        const MouldNum = row[headerIndexMap["Mould Number"]];
        return StatusCol == "Available" && MouldNum !== "";
      })
    : data;

  // STEP 1: Get existing Mould Numbers from target sheet
  const targetLastRow = targetSheet.getLastRow();
  const targetMouldNumColIndex = headersToExtract.indexOf("Mould Number") + 1;
  const existingMouldNums = new Set();

  if (targetLastRow >= startRow) {
    const existingValues = targetSheet.getRange(startRow, targetMouldNumColIndex, targetLastRow - startRow + 1).getValues();
    existingValues.forEach(([val]) => {
      if (val && val.toString().trim() !== "") {
        existingMouldNums.add(val.toString().trim());
      }
    });
  }

  // STEP 2: Filter out duplicates and format data
  const selectedData = filteredData
    .filter(row => {
      const mouldNumber = row[headerIndexMap["Mould Number"]].toString().trim();
      return !existingMouldNums.has(mouldNumber);
    })
    .map(row =>
      columnIndices.map((colIdx, i) => {
        const value = row[colIdx];
        const header = targetHeaders[i];
        if (
          formatDateColumns &&
          ["Final Actual Delivery Date"].includes(header) &&
          value
        ) {
          const date = new Date(value);
          if (!isNaN(date)) {
            const day = String(date.getDate()).padStart(2, '0');
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const year = date.getFullYear();
            return `${day}-${month}-${year}`;
          }
        }
        return value;
      })
    );

  // STEP 3: Paste to target only if new rows found
  if (selectedData.length > 0) {
    targetSheet.getRange(targetLastRow + 1, targetStartColumn, selectedData.length, selectedData[0].length).setValues(selectedData);
  }
}



/* =========================================================
   YOUR MODULAR CODE (Updated with Logging)
   ========================================================= */

function runMouldMasterSync() {
//modular

  const config = {
    sourceUrl: 'https://docs.google.com/spreadsheets/d/1ZPOAaV93ZUWrBxYHUrNjz6_Hm2dLKTEZ76-bTMQKSVA/edit?gid=0#gid=0',
    targetUrl: 'https://docs.google.com/spreadsheets/d/1ZPOAaV93ZUWrBxYHUrNjz6_Hm2dLKTEZ76-bTMQKSVA/edit?gid=1068030377#gid=1068030377',
    sourceSheetName: 'Source',
    targetSheetName: 'Target',
    headersToExtract: ["Student Name","Gender","Major","Marks","Joining Date"],
    formatDate: false
  };

  // --- GET SHEETS ---
  const sourceSheet = getSheet(config.sourceUrl, config.sourceSheetName);
  const targetSheet = getSheet(config.targetUrl, config.targetSheetName);

  // --- GET SOURCE DATA ---
  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getDisplayValues()[0];
  const sourceData = sourceSheet.getRange(1, 1, sourceSheet.getLastRow() - 2, sourceSheet.getLastColumn()).getDisplayValues();
  
  

  const headerMap = createHeaderIndexMap(sourceHeaders);

  // --- GET EXISTING ---
  const existingMouldNums = getExistingMouldNumbers(targetSheet, 1, 4); 

  // --- PROCESS ---
  const newRows = processAndFilterData(
    sourceData, 
    headerMap, 
    existingMouldNums, 
    config.headersToExtract, 
    config.formatDate
  );

  // --- WRITE ---
  if (newRows.length > 0) {
    appendDataToSheet(targetSheet, newRows, 1);
    
  } else {
    
  }
}

// --- HELPER FUNCTIONS WITH LOGS ---

function getSheet(url, sheetName) {
  
  const ss = SpreadsheetApp.openByUrl(url);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);
  return sheet;
}

function createHeaderIndexMap(headers) {
  const map = {};
  headers.forEach((header, index) => {
    map[header.trim()] = index;
  });
  // Log just a sample to keep it clean
  
  return map;
}

function getExistingMouldNumbers(sheet, columnIndex, startRow) {
  const lastRow = sheet.getLastRow();
  const existingSet = new Set();
  
  if (lastRow >= startRow) {
    const data = sheet.getRange(startRow, columnIndex, lastRow - startRow + 1).getValues();
    data.forEach(([val]) => {
      if (val && String(val).trim() !== "") existingSet.add(String(val).trim());
    });
  }
  
  return existingSet;
}

function processAndFilterData(data, headerMap, existingIds, headersToExtract, formatDate) {
  const processedRows = [];
  let skippedAvailable = 0;
  let skippedDuplicate = 0;

  const targetIndices = headersToExtract.map(name => headerMap[name]);
  const statusIndex = headerMap["Gender"];
  const mouldNumIndex = headerMap["Student Name"];

  
  for (let row of data) {
    const status = row[statusIndex];
    const mouldNum = String(row[mouldNumIndex]).trim();

    if (status !== "Male" || mouldNum === "") {
      skippedAvailable++;
      continue;
    }

    if (existingIds.has(mouldNum)) {
      skippedDuplicate++;
      continue;
    }

    const newRow = targetIndices.map((colIndex, i) => {
      let value = row[colIndex];
      const headerName = headersToExtract[i];
      if (formatDate && headerName === "Final Actual Delivery Date" && value) {
        value = formatMyDate(value);
      }
      return value;
    });

    processedRows.push(newRow);
  }

 

  return processedRows;
}

function appendDataToSheet(sheet, data, startColumn) {
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, startColumn, data.length, data[0].length).setValues(data);
}

function formatMyDate(dateValue) {
  // Simple check
  const date = new Date(dateValue);
  if (isNaN(date)) return dateValue; 
  return `${date.getDate()}-${date.getMonth() + 1}-${date.getFullYear()}`;
}
