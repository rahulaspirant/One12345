function start() {
  
  var ss= SpreadsheetApp.getActiveSpreadsheet();                     // Get active spreadsheet
  var setupSheet = ss.getSheetByName("SETUP");                       // Reference SETUP sheet
  var lastRow = setupSheet.getLastRow();                             // Find last row with data
  

  for(var y = 6; y <= lastRow; y++)                                  // Loop from row 6 to last row
  {
    var checkbox = setupSheet.getRange(y,8).getValue();              // Read checkbox (column H)
    
    if(checkbox == true)                                             // If checkbox is checked
    {
      setupSheet.getRange("D"+y+":G"+y).clear();                     // Clear previous result columns
      var symbol = setupSheet.getRange(y,3).getValue();              // Get stock symbol (column C)
      var market = setupSheet.getRange(y,2).getValue();              // Get market (column B)
      getData(symbol, market, y);                                    // Fetch data and calculate stats
      setLink(symbol, market, y);                                    // Set hyperlink on column I
      setupSheet.getRange(y,8).setValue(false);                      // Uncheck checkbox after processing
    }
  }
}

function checkAll() {
  
  var ss= SpreadsheetApp.getActiveSpreadsheet();                     // Get active spreadsheet
  var setupSheet = ss.getSheetByName("SETUP");                       // Reference SETUP sheet
  var lastRow = setupSheet.getLastRow();                             // Find last row
  
  for(var y = 6; y <= lastRow; y++)                                  // Loop through rows
  {
    setupSheet.getRange(y,8).setValue(true);                         // Set checkbox to TRUE
  }

}

function uncheckAll() {
  
  var ss= SpreadsheetApp.getActiveSpreadsheet();                     // Get active spreadsheet
  var setupSheet = ss.getSheetByName("SETUP");                       // Reference SETUP sheet
  var lastRow = setupSheet.getLastRow();                             // Find last row
  
  for(var y = 6; y <= lastRow; y++)                                  // Loop through rows
  {
    setupSheet.getRange(y,8).setValue(false);                        // Set checkbox to FALSE
  }

}

function setLink(symbol, market, row)
{
  var ss= SpreadsheetApp.getActiveSpreadsheet();                     // Get active spreadsheet
  var setupSheet = ss.getSheetByName("SETUP");                       // Reference SETUP sheet
  
  // Create rich-text value with hyperlink
  var richValue = SpreadsheetApp.newRichTextValue()
  .setText(symbol)
  .setLinkUrl("https://www.google.com/finance/quote/"+symbol+":"+market)
  .build();
  
  setupSheet.getRange(row, 9).setRichTextValue(richValue);           // Set link in column I
}

function getData(symbol, market, row) {

  var ss= SpreadsheetApp.getActiveSpreadsheet();                     // Get active spreadsheet
  var dataSheet = ss.getSheetByName("DATA");                         // Sheet used for temporary calculations
  var setupSheet = ss.getSheetByName("SETUP");                       // Reference SETUP sheet
  
  var percentGain = setupSheet.getRange(2,2).getValue();             // Threshold gain %
  var percentDown = setupSheet.getRange(3,2).getValue();             // Threshold down %
  var daysBack = setupSheet.getRange(1,2).getValue();                // Number of days to fetch data
  
  dataSheet.getRange("C1:D1000").clear();                            // Clear previous calculations

  // Place GOOGLEFINANCE formula in A1 to fetch historical prices
  var cell = dataSheet.getRange("A1");
  cell.setFormula("=GOOGLEFINANCE(\""+market+":"+symbol+"\",\"price\",TODAY()-"+daysBack+",TODAY())");
  
  dataSheet.getRange("C1").setValue(symbol);                         // Store symbol in temp sheet

  var datalastRow = dataSheet.getLastRow();                          // Last row of fetched data
  

  // Column C: Difference between today and previous day
  for(var i = 3; i <= datalastRow; i++)
  {
    var pastRow = i - 1;
    var cell = dataSheet.getRange(i,3);
    cell.setFormula("=B"+i+"-B"+pastRow+"");                         // Price change
  }

  // Column D: Percentage change
  for(var i = 3; i <= datalastRow; i++)
  {
    var pastRow = i - 1;
    var cell = dataSheet.getRange(i,4);
    cell.setFormula("=(C"+i+"/B"+pastRow+")*100");                   // % change formula
  }

  var percentGainTotal = 0;                                          // Count of gains crossing threshold
  var percentDownTotal = 0;                                          // Count of drops crossing threshold
  var consecutiveDaysGain = 0;                                       // Count of continuous positive % days

  // Count how many % changes exceed thresholds
  for(var i = 3; i <= datalastRow; i++)
  {
    var currentValue = dataSheet.getRange(i,4).getValue();
    if(currentValue > percentGain)
    {
      percentGainTotal++;                                            // Track gain days
    }
    if(currentValue < percentDown * -1)
    {
      percentDownTotal++;                                            // Track drop days
    }
  }

  // Count consecutive positive days starting from most recent day
  for(var i = datalastRow; i >= 3; i--)
  {
    var currentValue = dataSheet.getRange(i,4).getValue();
    if(currentValue > 0)
    {
      consecutiveDaysGain++;                                         // Increase streak count
    }
    else
    {
      break;                                                         // Stop when a negative % found
    }
  }

  // Output calculated results back to SETUP sheet
  setupSheet.getRange(row,5).setValue(percentGainTotal);             // Total gain days
  setupSheet.getRange(row,4).setValue(percentDownTotal);             // Total down days
  setupSheet.getRange(row,6).setValue(consecutiveDaysGain);          // Consecutive gains
  setupSheet.getRange(row,7).setValue(datalastRow-1);                // Total days of data (excluding header)

}
