function mySleep (sec) // <------------------------------------------------------- THIS IS IMPORTANT!!! 
{// <----------------------------------------------------------------------------- there are multiple calls to this function throughout the code 
SpreadsheetApp.flush();// <------------------------------------------------------- in order to make sure that the script completes the action before moving on to the next one
Utilities.sleep(sec*1000);
SpreadsheetApp.flush();
}

function clearFolder() {
var folder_id ='YOUR_DRIVE_FOLDER_ID';
var folder = DriveApp.getFolderById(folder_id);
const files = folder.getFiles();
while (files.hasNext()){
var file = files.next();
file.setTrashed(true); 
}
const folders = folder.getFolders();
while (folders.hasNext()){
var folder = folders.next();
folder.setTrashed(true); 
}
}

function sheetToCsv()
{
var tz = "GMT+3"; // <------------------------------------------------------------- Timezone (LT) (for the temporary folder name and the file names)
var tz2 = "GMT-4"; // <------------------------------------------------------------ Timezone (EST) (for the part where it adds the date to the tracker sheet)
var d_format = "yyyy-MM-dd"; // <-------------------------------------------------- Timestamp Format (for the temporary folder name and the file names)
var d2_format = "yyyy MM dd HHmm"; // <-------------------------------------------- new Timestamp format (for the part where it adds the date to the tracker sheet)
var date = Utilities.formatDate(new Date(), tz2, d_format);
var date2 = Utilities.formatDate(new Date(), tz, d2_format);
var ssID = SpreadsheetApp.getActiveSpreadsheet().getId();// <---------------------- gets the ID of the active sheet
var folder_id ='YOUR_DRIVE_FOLDER_ID'; // <---------------------------------------- EDIT THIS IF THE ROOT FOLDER CHANGES

//--------------------------------------------------------------------------------- Creates a new folder inside the root folder with the date as the folder name

var parentFolder = DriveApp.getFolderById(folder_id); // <------------------------- makes a new temporary folder in the root folder
var newFolderID = parentFolder.createFolder(date).getId(); // <-------------------- collects the ID of the temporary folder

//--------------------------------------------------------------------------------- variables for the sheet review

var ss = SpreadsheetApp.getActiveSpreadsheet();
var trackerSheetName = ss.getSheetName();
var trackerSheet = ss.getSheetByName(trackerSheetName);
var tsys = 0; 
var fdNorth = 0;
var fdOmaha = 0;
var heartland = 0;
var elavon = 0;
var chase = 0;
var vantiv = 0;
var procRow = 3;
var checkRow = 4;
var checkCol = 15;
sec = 2;
var requestData = {"method": "GET", "headers":{"Authorization":"Bearer "+ScriptApp.getOAuthToken()}};
//--------------------------------------------------------------------------------- Review part for the sheets that need csv files to be generated for (it does not generate all of them)

for (i=0; i<7; i++){
var newCol = checkCol + i;
var processor = trackerSheet.getRange(procRow, newCol).getValue();
var checkValue = trackerSheet.getRange(checkRow, newCol).getValue();
if(processor == "Template_F" && checkValue != ""){
tsys++;
}
else if(processor == "Template_C" && checkValue != ""){
fdNorth++;
}
else if(processor == "Template_D" && checkValue != ""){
fdOmaha++;
}
else if(processor == "Template_E" && checkValue != ""){
heartland++;
}
else if(processor == "Template_B" && checkValue != ""){
elavon++;
}
else if(processor == "Template_A" && checkValue != ""){
chase++;
}
else if(processor == "Template_G" && checkValue != ""){
vantiv++;
}
//Utilities.sleep(200);
}
mySleep(1);

//------------------------------------------------------------------------------------- Based on variable values, this is where it starts generating csv files for specific sheets into the new temporary folder in drive
//------------------------------------------------------------------------------------- +Trims WhiteSpace
//------------------------------------------------------------------------------------- +Removes any comma's in the cells
//------------------------------------------------------------------------------------- +Copies the Client_ID to the tracker sheet
//------------------------------------------------------------------------------------- +Deletes info from the template sheet once the file is generated
//------------------------------------------------------------------------------------- +Adds the client_id from the template sheet to the tracker sheet and sets the submission date
//------------------------------------------------------------------------------------- +Shows a prompt once the file is generated (must click "OK" to continue)


if (tsys > 0){
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template_F"); // <------------- This part for TSYS sheet
var sheetNameId = sheet.getSheetId().toString();
var range = sheet.getDataRange();
range.trimWhitespace(); // <--------------------------------------------------------------------- "trims" any WhiteSpace from the tempoate sheet cells
var data = range.getValues();
for (var row = 0; row < data.length; row++) { // <----------------------------------------------- "removes" any comma's left in the cells and rewrites the cell info
for (var col = 0; col < data[row].length; col++) {
data[row][col] = (data[row][col]).toString().replace(/,/g, '');
}
}
range.setValues(data);
mySleep(1);

params= ssID+"/export?gid="+sheetNameId +"&format=csv"; // <------------------------------------- generates the CSVfile of this specific sheet to the temporary drive folder
var url = "https://docs.google.com/spreadsheets/d/"+ params;
var result = UrlFetchApp.fetch(url, requestData); 
var resource = {
title: "TSYS Bulk Board Template "+date2+".csv",
mimeType: "application/vnd.csv",
parents: [{ id: newFolderID }]
};
var fileJson = Drive.Files.insert(resource,result);

var aVals = trackerSheet.getRange("A1:A").getValues(); // <-------------------------------------- checks for the first empty cell in the "A" column in the tracker sheet
var aLast = aVals.filter(String).length+1;
var l = 1;
//Logger.log("aLast - "+aLast);
while (sheet.getLastRow()>1){ // <--------------------------------------------------------------- cleanup part for copying the client_id from template to tracker + deleting rows with data from template sheet
var lr = sheet.getLastRow();
var lc = sheet.getLastColumn();
var source = sheet.getRange (lr,1,1,1);
var destination = trackerSheet.getRange(aLast+l,1,1,1);
source.copyTo(destination, {contentsOnly:true});
trackerSheet.getRange(aLast+l,3,1,1).setValue(date);
var clear = sheet.getRange(lr,1,1,lc);
clear.clearContent();
l++;
//Utilities.sleep(500);
};
//Browser.msgBox("TSYS .csv file has been generated.\\nTemplate sheet was cleared out."); // <--- prompt to click when the cleanup is completed (comment this if it gets annoying to keep confirming for every template)
mySleep(1);
}

if (fdNorth > 0){
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template_C"); // <------------- This part for First Data North sheet (full function comments above in the "TSYS part of the code")
var sheetNameId = sheet.getSheetId().toString();
var range = sheet.getDataRange();
range.trimWhitespace();
var data = range.getValues();
for (var row = 0; row < data.length; row++) {
for (var col = 0; col < data[row].length; col++) {
data[row][col] = (data[row][col]).toString().replace(/,/g, '');
}
}
range.setValues(data);
mySleep(1);

params= ssID+"/export?gid="+sheetNameId +"&format=csv";
var url = "https://docs.google.com/spreadsheets/d/"+ params;
var result = UrlFetchApp.fetch(url, requestData); 
var resource = {
title: "FirstDataNashvillePTS (North BE) "+date2+".csv",
mimeType: "application/vnd.csv",
parents: [{ id: newFolderID }]
};
var fileJson = Drive.Files.insert(resource,result);

var aVals = trackerSheet.getRange("A1:A").getValues();
var aLast = aVals.filter(String).length+1;
var l = 1;
//Logger.log("aLast - "+aLast);
while (sheet.getLastRow()>1){
var lr = sheet.getLastRow();
var lc = sheet.getLastColumn();
var source = sheet.getRange (lr,2,1,1);
var destination = trackerSheet.getRange(aLast+l,1,1,1);
source.copyTo(destination, {contentsOnly:true});
trackerSheet.getRange(aLast+l,3,1,1).setValue(date);
var clear = sheet.getRange(lr,1,1,lc);
clear.clearContent();
l++;
//Utilities.sleep(500);
};
//Browser.msgBox("First Data North .csv file has been generated.\\nTemplate sheet was cleared out.");
mySleep(1);
}

if (fdOmaha > 0){
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template_D"); // <------------- This part for First Data Omaha sheet (full function comments above in the "TSYS part of the code")
var sheetNameId = sheet.getSheetId().toString();
var range = sheet.getDataRange();
range.trimWhitespace();
var data = range.getValues();
for (var row = 0; row < data.length; row++) {
for (var col = 0; col < data[row].length; col++) {
data[row][col] = (data[row][col]).toString().replace(/,/g, '');
}
}
range.setValues(data);
mySleep(1);

params= ssID+"/export?gid="+sheetNameId +"&format=csv";
var url = "https://docs.google.com/spreadsheets/d/"+ params;
var result = UrlFetchApp.fetch(url, requestData); 
var resource = {
title: "FirstDataNashvilleHDC (Omaha) "+date2+".csv",
mimeType: "application/vnd.csv",
parents: [{ id: newFolderID }]
};
var fileJson = Drive.Files.insert(resource,result);

var aVals = trackerSheet.getRange("A1:A").getValues();
var aLast = aVals.filter(String).length+1;
var l = 1;
//Logger.log("aLast - "+aLast);
while (sheet.getLastRow()>1){
var lr = sheet.getLastRow();
var lc = sheet.getLastColumn();
var source = sheet.getRange (lr,1,1,1);
var destination = trackerSheet.getRange(aLast+l,1,1,1);
source.copyTo(destination, {contentsOnly:true});
trackerSheet.getRange(aLast+l,3,1,1).setValue(date);
var clear = sheet.getRange(lr,1,1,lc);
clear.clearContent();
l++;
//Utilities.sleep(500);
};
//Browser.msgBox("First Data Omaha .csv file has been generated.\\nTemplate sheet was cleared out.");
mySleep(1);
}

if (heartland > 0){
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template_E"); // <------------ This part for Heartland sheet (full function comments above in the "TSYS part of the code")
var sheetNameId = sheet.getSheetId().toString();
var range = sheet.getDataRange();
range.trimWhitespace();
var data = range.getValues();
for (var row = 0; row < data.length; row++) {
for (var col = 0; col < data[row].length; col++) {
data[row][col] = (data[row][col]).toString().replace(/,/g, '');
}
}
range.setValues(data);
mySleep(1);

params= ssID+"/export?gid="+sheetNameId +"&format=csv";
var url = "https://docs.google.com/spreadsheets/d/"+ params;
var result = UrlFetchApp.fetch(url, requestData); 
var resource = {
title: "Heartland Bulk Board Template "+date2+".csv",
mimeType: "application/vnd.csv",
parents: [{ id: newFolderID }]
};
var fileJson = Drive.Files.insert(resource,result);

var aVals = trackerSheet.getRange("A1:A").getValues();
var aLast = aVals.filter(String).length+1;
var l = 1;
//Logger.log("aLast - "+aLast);
while (sheet.getLastRow()>1){
var lr = sheet.getLastRow();
var lc = sheet.getLastColumn();
var source = sheet.getRange (lr,1,1,1);
var destination = trackerSheet.getRange(aLast+l,1,1,1);
source.copyTo(destination, {contentsOnly:true});
trackerSheet.getRange(aLast+l,3,1,1).setValue(date);
var clear = sheet.getRange(lr,1,1,lc);
clear.clearContent();
l++;
//Utilities.sleep(500);
};
//Browser.msgBox("Heartland .csv file has been generated.\\nTemplate sheet was cleared out.");
mySleep(1);
}

if (elavon > 0){
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template_B"); // <------------ This part for Elavon sheet (full function comments above in the "TSYS part of the code")
var sheetNameId = sheet.getSheetId().toString();
var range = sheet.getDataRange();
range.trimWhitespace();
var data = range.getValues();
for (var row = 0; row < data.length; row++) {
for (var col = 0; col < data[row].length; col++) {
data[row][col] = (data[row][col]).toString().replace(/,/g, '');
}
}
range.setValues(data);
mySleep(1);

params= ssID+"/export?gid="+sheetNameId +"&format=csv";
var url = "https://docs.google.com/spreadsheets/d/"+ params;
var result = UrlFetchApp.fetch(url, requestData); 
var resource = {
title: "Elavon Bulk Board Template "+date2+".csv",
mimeType: "application/vnd.csv",
parents: [{ id: newFolderID }]
};
var fileJson = Drive.Files.insert(resource,result);

var aVals = trackerSheet.getRange("A1:A").getValues();
var aLast = aVals.filter(String).length+1;
var l = 1;
//Logger.log("aLast - "+aLast);
while (sheet.getLastRow()>1){
var lr = sheet.getLastRow();
var lc = sheet.getLastColumn();
var source = sheet.getRange (lr,1,1,1);
var destination = trackerSheet.getRange(aLast+l,1,1,1);
source.copyTo(destination, {contentsOnly:true});
trackerSheet.getRange(aLast+l,3,1,1).setValue(date);
var clear = sheet.getRange(lr,1,1,lc);
clear.clearContent();
l++;
//Utilities.sleep(500);
};
//Browser.msgBox("Elavon .csv file has been generated.\\nTemplate sheet was cleared out.");
mySleep(1);
}

if (chase > 0){
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template_A"); // <------------ This part for Chase Paymentech sheet (full function comments above in the "TSYS part of the code")
var sheetNameId = sheet.getSheetId().toString();
var range = sheet.getDataRange();
range.trimWhitespace();
var data = range.getValues();
for (var row = 0; row < data.length; row++) {
for (var col = 0; col < data[row].length; col++) {
data[row][col] = (data[row][col]).toString().replace(/,/g, '');
}
}
range.setValues(data);
mySleep(1);

params= ssID+"/export?gid="+sheetNameId +"&format=csv";
var url = "https://docs.google.com/spreadsheets/d/"+ params;
var result = UrlFetchApp.fetch(url, requestData); 
var resource = {
title: "ChasePaymentech "+date2+".csv",
mimeType: "application/vnd.csv",
parents: [{ id: newFolderID }]
};
var fileJson = Drive.Files.insert(resource,result);

var aVals = trackerSheet.getRange("A1:A").getValues();
var aLast = aVals.filter(String).length+1;
var l = 1;
//Logger.log("aLast - "+aLast);
while (sheet.getLastRow()>1){
var lr = sheet.getLastRow();
var lc = sheet.getLastColumn();
var source = sheet.getRange (lr,1,1,1);
var destination = trackerSheet.getRange(aLast+l,1,1,1);
source.copyTo(destination, {contentsOnly:true});
trackerSheet.getRange(aLast+l,3,1,1).setValue(date);
var clear = sheet.getRange(lr,1,1,lc);
clear.clearContent();
l++;
//Utilities.sleep(500);
};
//Browser.msgBox("Chase Paymentech .csv file has been generated.\\nTemplate sheet was cleared out.");
mySleep(1);
}
if (vantiv > 0){
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template_G"); // <------------ This part for WorldPay (Vantiv) sheet (full function comments above in the "TSYS part of the code")
//Logger.log(sheet.getName());
var sheetNameId = sheet.getSheetId().toString();
var range = sheet.getDataRange();
range.trimWhitespace();
var data = range.getValues();
for (var row = 0; row < data.length; row++) {
for (var col = 0; col < data[row].length; col++) {
data[row][col] = (data[row][col]).toString().replace(/,/g, '');
}
}
range.setValues(data);
mySleep(1);

params= ssID+"/export?gid="+sheetNameId +"&format=csv";
var url = "https://docs.google.com/spreadsheets/d/"+ params;
var result = UrlFetchApp.fetch(url, requestData); 
var resource = {
title: "Vantiv Bulk Boarding Template "+date2+".csv",
mimeType: "application/vnd.csv",
parents: [{ id: newFolderID }]
};
var fileJson = Drive.Files.insert(resource,result);

var aVals = trackerSheet.getRange("A1:A").getValues();
var aLast = aVals.filter(String).length+1;
var l = 1;
//Logger.log("aLast - "+aLast);
while (sheet.getLastRow()>1){
var lr = sheet.getLastRow();
var lc = sheet.getLastColumn();
var source = sheet.getRange (lr,1,1,1);
var destination = trackerSheet.getRange(aLast+l,1,1,1);
source.copyTo(destination, {contentsOnly:true});
trackerSheet.getRange(aLast+l,3,1,1).setValue(date);
var clear = sheet.getRange(lr,1,1,lc);
clear.clearContent();
l++;
//Utilities.sleep(500);
};
//Browser.msgBox("Vantiv (WorldPay) .csv file has been generated.\\nTemplate sheet was cleared out.");
mySleep(1);
}

//-------------------------------------------------------------------------------------------------- CREATES A TEMPORARY FOLDER FOR DOWNLOAD LINK

var zipFolder = DriveApp.getFolderById(newFolderID);
var files = zipFolder.getFiles();
var blobs = [];
while (files.hasNext()) {
blobs.push(files.next().getBlob());
};
var zipBlob = Utilities.zip(blobs, zipFolder.getName() + ".zip");
var fileId = DriveApp.createFile(zipBlob).getId();
var URL = "https://drive.google.com/uc?export=download&id=" + fileId;
var URL2 = "https://drive.google.com/drive/folders/YOUR_DRIVE_FOLDER_ID";

//------------------------------------------------------------------------------------------------ SHOWS A PROMPT WITH THE FOLDER DOWNLOAD LINK

var htmlOutput = HtmlService
.createHtmlOutput('<p><a href="'+URL+'" target="_blank">Download the CSV files</a></p><hr><p>Or click below if the above fails</p><p><a href="'+URL2+'" target="_blank">Visit the drive folder (opens a new window)</a></p><p color="red">Don\'t forget to delete the dated folder after it\'s downloaded.</p>')
.setSandboxMode(HtmlService.SandboxMode.IFRAME)
.setWidth(380)
.setHeight(220);
/*var htmlOutput2 = HtmlService
.createHtmlOutput('<p><a href="'+URL2+'" target="_blank">Visit the drive folder (opens a new window)</a></p><p color="red">Don\'t forget to delete the dated folder after it\'s downloaded.</p>')
.setSandboxMode(HtmlService.SandboxMode.IFRAME)
.setWidth(380)
.setHeight(120);
var ui = SpreadsheetApp.getUi();
var response = ui.alert('Answer:','Are you the owner of the FPBB_CSV drive folder?', ui.ButtonSet.YES_NO);
if (response == ui.Button.YES){*/
SpreadsheetApp.getUi().showModelessDialog(htmlOutput, 'Download the CSV files');
/* DriveApp.getFolderById(newFolderID).setTrashed(true);
}
else{
SpreadsheetApp.getUi().showModelessDialog(htmlOutput2, 'Redirect to google drive for CSV files');
}*/

Logger.log(URL);
//DriveApp.getFolderById(newFolderID).setTrashed(true); // <------------------------------------ DELETES (sends to "trash") THE TEMPORARY FOLDER THAT WAS GENERATED AT THE START OF THIS SCRIPT

} 