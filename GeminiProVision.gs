const properties = PropertiesService.getScriptProperties().getProperties();
const geminiApiKey = properties['GOOGLE_API_KEY'];
const geminiProVisionEndpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${geminiApiKey}`;
const GeminiFormat = 'Output shall be provided in English script containing all values as described in Scope section for all line-items in the given image. NA to be used for denoting missing values.';

function callGeminiProVision(prompt, image, temperature=0, topK=2) {
const imageData = Utilities.base64Encode(image.getAs('image/png').getBytes());

const payload = {
"contents": [
{
"parts": [
{
"text": prompt
},
{
"inlineData": {
"mimeType": "image/png",
"data": imageData
}
}
]
}
],
"generationConfig": {
"temperature": temperature,
"topK": topK,
"responseMimeType": "application/json"
},
};

const options = {
'method' : 'post',
'contentType': 'application/json',
'payload': JSON.stringify(payload)
};

const response = UrlFetchApp.fetch(geminiProVisionEndpoint, options);
const data = JSON.parse(response);
const content = data["candidates"][0]["content"]["parts"][0]["text"];
return content;
}


function listImageFileBlobs(folderId) {
var fileBlobs = [];
var folder = DriveApp.getFolderById(folderId);

// Get all files in the current folder
var files = folder.getFiles();
while (files.hasNext()) {
var file = files.next();
var mimeType = file.getMimeType();

// Check if the file is a JPG or PNG
if (mimeType === 'image/jpeg' || mimeType === 'image/png') {
fileBlobs.push([file,file.getBlob()]);
}
}

// Get all subfolders and process them recursively
var subFolders = folder.getFolders();
while (subFolders.hasNext()) {
var subfolder = subFolders.next();
var subfolderfileBlobs = listImageFileBlobs(subfolder); // Recursive call
fileBlobs = fileBlobs.concat(subfolderfileBlobs);
}

return fileBlobs;
}


function writeArrtoSheet(ImgArr, file, datasheet,headers, indexval){

//file Url in col 1
datasheet.getRange(indexval,1).setValue(file.getUrl());

// data pushed in col 2
const rowsNum = ImgArr.length;
const range = datasheet.getRange(indexval,2,rowsNum,headers.length);
range.setValues(ImgArr);

// return updated index
return indexval+rowsNum;

}


function readFiles(fileBlobs,datasheet,indexrange,indexval,GeminiRole,field_def,headers, destfolderId) {

// Get the destination folder
var destfolder = DriveApp.getFolderById(destfolderId);

// Loop through each file ID
fileBlobs.forEach(function(fileBlob) {
try {

// Move the file to the destination folder
var file=fileBlob[0];
var image=fileBlob[1];

indexval = writeArrtoSheet(GetImgArr(image,GeminiRole,field_def,headers),file,datasheet,headers,indexval);
file.moveTo(destfolder);
indexrange.setValue(indexval);

} catch (e) {
// Log any errors
console.log('Error related to ' + file.getName() + ': ' + e.toString());
}
});
}


function getHeaders(datasheet) {
// Get the range of the first row (headers) in the datasheet
var headerRange = datasheet.getRange(1, 2, 1, datasheet.getLastColumn() - 1);

// Get the values of the header range as a 2D array
var headerValues = headerRange.getValues();

// Flatten the 2D array to a 1D array
var headers = headerValues[0];

return headers;
}


function getGoogleSheetById(fileId) {

const spreadsheet = SpreadsheetApp.openById(fileId);
const datasheet = spreadsheet.getSheetByName('Data');
const indexrange = spreadsheet.getSheetByName('Index').getRange('A1');
const GeminiRole = spreadsheet.getSheetByName('Role').getRange('A1').getValue();
const field_def = spreadsheet.getSheetByName('Definitions').getDataRange().getValues();
return [datasheet,getHeaders(datasheet),indexrange,indexrange.getValue(),GeminiRole,field_def];

}


function csvStringToArray(csvString) {
return csvString.split('\n').slice(1, -1).map(line => line.split(','));
}


function GenGeminiDefString(field_def) {

// Ignore the first row by slicing the array from index 1
var slicedArray = field_def.slice(1);

// Map each row to the desired format
var formattedStrings = slicedArray.map(function(row) {
return row[0] + " (" + row[1] + ")";
});

// Join the formatted strings with a newline character
var strfield = formattedStrings.join("\n");

return strfield;
}


function jsonToArray(jsonObj, headers) {
// Automatically get the first (and only) key in the JSON object
const dataKey = Object.keys(jsonObj)[0];

// Extract the data array from the JSON object
const data = jsonObj[dataKey];

// Initialize the result array
const result = [];

// Loop through each entry in the data array
data.forEach(entry => {
// Initialize a row array
const row = [];

// Loop through each header
headers.forEach(header => {
// Check if the header exists in the entry
if (entry.hasOwnProperty(header)) {
// Push the value to the row array
row.push(entry[header]);
} else {
// Push "NA" if the header is not found
row.push("NA");
}
});

// Add the row to the result array
result.push(row);
});

return result;
}


function GenGeminiPrompt(GeminiRole,field_def){
return 'Role & Objective:\n' + GeminiRole + '\n\nScope:\nYou will exclusively capture values corresponding to below Standard Fields in the given image (Definition and data type in parenthesis).\n' + GenGeminiDefString(field_def) + '\n\nFormat:\n' + GeminiFormat;
}


function GetImgArr(image,GeminiRole,field_def,headers) {
const prompt = GenGeminiPrompt(GeminiRole,field_def);
const GemJSON = callGeminiProVision(prompt, image);
return jsonToArray(JSON.parse(GemJSON),headers);
}


function execBatch(databaseId,toreadfolderId,readfolderId){

const [datasheet,headers,indexrange,indexval,GeminiRole,field_def]=getGoogleSheetById(databaseId);
const fileBlobs=listImageFileBlobs(toreadfolderId);
readFiles(fileBlobs,datasheet,indexrange,indexval,GeminiRole,field_def,headers,readfolderId);

}
