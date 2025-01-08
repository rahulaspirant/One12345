// Construct PDF name
    const pdfName = companyName +".pdf";

    // Construct URL for exporting PDF
    const url = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?exportFormat=pdf&format=pdf&gid=${sheet.getSheetId()}&gridrange=${range.getA1Notation()}`;

    // Fetch PDF
    const options = {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true // Add this line
    };
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      throw new Error(`Failed to fetch PDF. Response code: ${response.getResponseCode()}`);
    }
    const pdfBlob = response.getBlob();

    // Create PDF file in the output folder
    const newFile = outputFolder.createFile(pdfBlob).setName(pdfName);
    const pdfUrl = newFile.getUrl();
