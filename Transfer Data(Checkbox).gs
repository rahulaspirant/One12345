function onEdit(e) {
  //Transfer Data(Checkbox)
  //https://stackoverflow.com/questions/79024480/creating-a-project-management-file-in-google-sheets-need-help-adding-row-to-the
  
  const src = e.source.getActiveSheet();
  const r = e.range;

  if (src.getName() == "Tasks" && r.columnStart == 7 && r.rowStart != 1) {
    if (r.getValue()) {
      const dest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Complete");
      dest.appendRow(src.getRange(r.rowStart, 1, 1, 6).getValues()[0]);
      src.deleteRow(r.rowStart);
      dest.getRange(dest.getLastRow(), 7).insertCheckboxes();
      dest.getRange(dest.getLastRow(), 7).check();
    }
  } else if (src.getName() == "Complete" && r.columnStart == 7 && r.rowStart != 1) {
    if (!r.getValue()) {
      const ret = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tasks");
      ret.appendRow(src.getRange(r.rowStart, 1, 1, 6).getValues()[0]);
      src.deleteRow(r.rowStart);
      ret.getRange(ret.getLastRow(), 7).insertCheckboxes();
      ret.getRange(ret.getLastRow(), 7).uncheck();
    }
  }
}

//Ver 2
function onEdit(e) {
  const src = e.source.getActiveSheet();
  const r = e.range;
  if (src.getName() != "Tasks" || r.columnStart != 7 || r.rowStart == 1) return;
  const dest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Complete");
  dest.insertRows(2);
  src.getRange(r.rowStart,1,1,7).moveTo(dest.getRange(2,1,1,7));
  src.deleteRow(r.rowStart);
 }
