function linkImagesToProducts() {
  //https://drive.google.com/drive/folders/1aN0DnS5gOSmiAX4oct51cFAqiDGlr240?usp=drive_link
  //This will upload images from googledrive to googlesheet
  
  const folderId = '1aN0DnS5gOSmiAX4oct51cFAqiDGlr240'; // Folder ID of the images
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();

  // Set the dynamic columns
  const productNameCol = 2;  // Column number for Product Name (e.g., 2 for column B)
  const imageLinkCol = 8;    // Column number for Image Links (e.g., 9 for column I)

  // Access your Google Sheet and fetch the product names dynamically
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Product Master');
  const data = sheet.getRange(2, productNameCol, sheet.getLastRow() - 1, 1).getValues(); 
  // Fetch only the Product Name column dynamically based on the productNameCol variable

  let imageUrlList = [];

  // Create a map of all available images by name for faster lookups
  let imageMap = {};
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName(); // e.g., ProductA.jpg or ProductA_Image.jpg
    const fileUrl = file.getUrl();
    imageMap[fileName] = fileUrl; // Map file names to their URLs
  }

  // Iterate through the products and check for image availability
  for (let i = 0; i < data.length; i++) {
    const productName = data[i][0]; // Product Name from the specified column
    let imageUrl = "No Image Available"; // Default value if no image found

    // Search through imageMap for any file that contains the product name
    for (let fileName in imageMap) {
      if (fileName.includes(productName)) {
        imageUrl = imageMap[fileName]; // If a match is found, use the image URL
        break; // Stop searching once a match is found
      }
    }

    // Add the image URL or "No Image Available" to the list
    imageUrlList.push([imageUrl]);
  }

  // Write the URLs back to the Google Sheet in the dynamic column for Image Links
  sheet.getRange(2, imageLinkCol, imageUrlList.length, 1).setValues(imageUrlList);
}
