var config = {
  '1jHpFdNeADrczvbcIm0mDbIOu-7M9dd5WW98X21XEr5Q': {   //Design & Dev. Input form
    sheetName: 'Dropdown',
    dropdowns: [
      { range: 'L2:L', itemIndex: 0 },
      //{ range: 'B2:B', itemIndex: 1 }
    ]
   },
  '1dyY6qfc0JN1NmLraC_dSUHGPPLPj2l11MzZGbUhoyEY': {  //PT - Design Review-1(Review record)
    sheetName: 'Dropdown',
    dropdowns: [
      { range: 'L2:L', itemIndex: 0 },
      //{ range: 'D2:D', itemIndex: 2 }
    ]
  },////////////////////////////////////////////////////////
  
   '1Nd1z13VVcyav7Hv9ba7ys3fia9OCYw9ECIxCmD3vbHA': {  //PT - Design Review-2(Review record)
    sheetName: 'Dropdown',
    dropdowns: [
      { range: 'L2:L', itemIndex: 0 },
      //{ range: 'D2:D', itemIndex: 2 }
    ]
  }
  ,
  
   '1H7rdwyBf9jeUEIetE33tyEwRSRwMsAKcuw8s2dxr7vk': {  //PT - Design Validation
    sheetName: 'Dropdown',
    dropdowns: [
      { range: 'L2:L', itemIndex: 0 },
      //{ range: 'D2:D', itemIndex: 2 }
    ]
  }
 ,
    '1_NafHcZMMSz0wkzS7pGhkONdcYGeGPZ-TJ3cixmgPdo': {  //PT - Design Completion Certificate
    sheetName: 'Dropdown',
    dropdowns: [
      { range: 'L2:L', itemIndex: 0 },
      //{ range: 'D2:D', itemIndex: 2 }
    ]
  }
  
  //Add more forms as needed
};



function updateFormMultipleQ() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  for (var formId in config) {
    if (config.hasOwnProperty(formId)) {
      var formConfig = config[formId];
      var form = FormApp.openById(formId);
      var sheet = spreadsheet.getSheetByName(formConfig.sheetName);
      
      if (!sheet) {
        Logger.log('Sheet ' + formConfig.sheetName + ' not found for form ' + formId);
        continue;
      }

      var formItems = form.getItems(FormApp.ItemType.LIST);

      formConfig.dropdowns.forEach(function(dropdown) {
        var data = sheet.getRange(dropdown.range).getValues();
        var items = data.flat().filter(String); // Flatten and filter empty strings

        if (dropdown.itemIndex < formItems.length) {
          var listItem = formItems[dropdown.itemIndex].asListItem();
          listItem.setChoiceValues(items);
          Logger.log('Dropdown at index ' + dropdown.itemIndex + ' updated for form ' + formId);
        } else {
          Logger.log('Item index ' + dropdown.itemIndex + ' out of bounds for form ' + formId);
        }
      });
    }
  }
}
