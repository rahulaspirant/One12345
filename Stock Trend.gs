/***************** SIMPLE CONFIG ************************/

// Sheet names
const SETUP_SHEET_NAME = "SETUP";
const DATA_SHEET_NAME  = "DATA";

// SETUP sheet layout
const SETUP_FIRST_DATA_ROW = 6; // Your data starts at row 6

// Column numbers in SETUP sheet (1-based)
const COL_NAME       = 1;
const COL_MARKET     = 2;
const COL_SYMBOL     = 3;
const COL_DAYS_DROP  = 4;
const COL_DAYS_GAIN  = 5;
const COL_CNS_GAIN   = 6;
const COL_TOTAL_DAYS = 7;
const COL_CHECK      = 8;
const COL_LINK       = 9;

// Cells in SETUP that store parameters
const CELL_DAYS_BACK     = "B1"; // how many days back
const CELL_PERCENT_GAIN  = "B2"; // gain threshold
const CELL_PERCENT_DOWN  = "B3"; // drop threshold

/***************** SMALL HELPERS ************************/

function getSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    setup: ss.getSheetByName(SETUP_SHEET_NAME),
    data: ss.getSheetByName(DATA_SHEET_NAME)
  };
}

// Set all checkboxes in SETUP (from first data row downwards) to true/false
function setAllCheckboxes_(value) {
  var sheets = getSheets_();
  var setupSheet = sheets.setup;
  var lastRow = setupSheet.getLastRow();
  if (lastRow < SETUP_FIRST_DATA_ROW) return; // nothing to do
  Logger.log(sheets);

  var numRows = lastRow - SETUP_FIRST_DATA_ROW + 1;
  // Fill an array with the same value (true or false)
  var values = [];
  for (var i = 0; i < numRows; i++) {
    values.push([value]);
  }

  setupSheet
    .getRange(SETUP_FIRST_DATA_ROW, COL_CHECK, numRows, 1)
    .setValues(values);
}

/***************** MAIN FUNCTIONS ************************/

function start() {
  var sheets = getSheets_();
  var setupSheet = sheets.setup;
  var lastRow = setupSheet.getLastRow();

  for (var row = SETUP_FIRST_DATA_ROW; row <= lastRow; row++) {
    var checkbox = setupSheet.getRange(row, COL_CHECK).getValue();

    if (checkbox === true) {
      // Clear DAYS DROP, DAYS GAIN, CNS GAIN DAY, TOTAL DAYS
      setupSheet
        .getRange(row, COL_DAYS_DROP, 1, 4) // from D to G in that row
        .clearContent();

      var symbol = setupSheet.getRange(row, COL_SYMBOL).getValue();
      var market = setupSheet.getRange(row, COL_MARKET).getValue();

      getData(symbol, market, row);
      setLink(symbol, market, row);

      // Uncheck after processing
      setupSheet.getRange(row, COL_CHECK).setValue(false);
    }
  }
}

function checkAll() {
  setAllCheckboxes_(true);
}

function uncheckAll() {
  setAllCheckboxes_(false);
}

function setLink(symbol, market, row) {
  var sheets = getSheets_();
  var setupSheet = sheets.setup;

  var richValue = SpreadsheetApp.newRichTextValue()
    .setText(symbol)
    .setLinkUrl("https://www.google.com/finance/quote/" + symbol + ":" + market)
    .build();

  setupSheet.getRange(row, COL_LINK).setRichTextValue(richValue);
}

function getData(symbol, market, row) {
  var sheets = getSheets_();
  var dataSheet = sheets.data;
  var setupSheet = sheets.setup;

  // Read parameters from SETUP
  var daysBack    = setupSheet.getRange(CELL_DAYS_BACK).getValue();
  var percentGain = setupSheet.getRange(CELL_PERCENT_GAIN).getValue();
  var percentDown = setupSheet.getRange(CELL_PERCENT_DOWN).getValue();

  // Clear old helper columns (C and D) in DATA
  var maxRows = dataSheet.getMaxRows();
  dataSheet.getRange(1, 3, maxRows, 2).clearContent(); // C:D

  // Put GOOGLEFINANCE formula
  var cell = dataSheet.getRange("A1");
  cell.setFormula(
    '=GOOGLEFINANCE("' + market + ':' + symbol + '","price",TODAY()-' + daysBack + ',TODAY())'
  );

  // Put symbol header in C1
  dataSheet.getRange("C1").setValue(symbol);

  // Wait for data to populate (optional: sometimes needed)
  SpreadsheetApp.flush();

  var datalastRow = dataSheet.getLastRow();

  // Set difference formula in column C
  for (var i = 3; i <= datalastRow; i++) {
    var pastRow = i - 1;
    var diffCell = dataSheet.getRange(i, 3);
    diffCell.setFormula("=B" + i + "-B" + pastRow);
  }

  // Set percent change formula in column D
  for (var j = 3; j <= datalastRow; j++) {
    var pastRow2 = j - 1;
    var pctCell = dataSheet.getRange(j, 4);
    pctCell.setFormula("=(C" + j + "/B" + pastRow2 + ")*100");
  }

  SpreadsheetApp.flush(); // make sure formulas calculate

  var percentGainTotal = 0;
  var percentDownTotal = 0;
  var consecutiveDaysGain = 0;

  // Count gain / drop days
  for (var r = 3; r <= datalastRow; r++) {
    var currentValue = dataSheet.getRange(r, 4).getValue(); // column D
    if (currentValue > percentGain) {
      percentGainTotal++;
    }
    if (currentValue < percentDown * -1) {
      percentDownTotal++;
    }
  }

  // Count consecutive gain days from the bottom
  for (var r2 = datalastRow; r2 >= 3; r2--) {
    var currentValue2 = dataSheet.getRange(r2, 4).getValue();
    if (currentValue2 > 0) {
      consecutiveDaysGain++;
    } else {
      break;
    }
  }

  // Write back to SETUP
  setupSheet.getRange(row, COL_DAYS_GAIN).setValue(percentGainTotal);
  setupSheet.getRange(row, COL_DAYS_DROP).setValue(percentDownTotal);
  setupSheet.getRange(row, COL_CNS_GAIN).setValue(consecutiveDaysGain);
  setupSheet.getRange(row, COL_TOTAL_DAYS).setValue(datalastRow - 1);
}
