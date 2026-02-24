function onOpen() {
var ui = SpreadsheetApp.getUi();

ui.createMenu('Sheet Workflow Tools')
.addItem('Custom STM + DBM list insta-sort', 'menuItemMainSort')
.addSeparator()
.addSubMenu(ui.createMenu('In_Progress sheet actions')
.addItem('Move entries to Recovered', 'menuItemRecovered')
.addItem('Move entries to Voids', 'menuItemVoids2')
.addSeparator()
.addItem('Update Cancellation status', 'menuItemSFStatus'))
.addSeparator()
.addSubMenu(ui.createMenu('Generate outputs')
.addItem('Generate EMAIL output', 'menuItemEmailTemplate')
.addSeparator()
.addItem('Generate RETRY output', 'menuItemRetry')
.addSeparator()
.addItem('Generate Status_Lookup output', 'menuItemStmStatus'))
.addToUi();

}

function menuItemSFStatus() {
var ui = SpreadsheetApp.getUi();

var result = ui.alert(
"Please Read and Confirm",
"This script will review all the records in the 'In_Progress' sheet and update their cancellation statuses\n\nPlease confirm the uploaded status sheet is called 'Status_Lookup'.",
ui.ButtonSet.YES_NO,
);

if (result == ui.Button.YES) {

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sfSheet = ss.getSheetByName("Status_Lookup");
var proSheet = ss.getSheetByName("In_Progress");

var row = 3;
var count = 0;
var actionCount = 0;

var tz = ss.getSpreadsheetTimeZone();
var dt = new Date();
var date = Utilities.formatDate(dt,tz, "yyyy-MM-dd");

var proLR = proSheet.getDataRange().getNumRows();
var recordCount = proLR - row + 1;
var cidCount = 0;

proSheet.getRange(3,14,proLR-2).clearNote();

for (i = 0; i < recordCount; i=i+cidCount) {
var index = row+i;
cidCount = proSheet.getRange(index,7).getValue();
proSheet.getRange(index,14,cidCount,1).setNote("Checked on: "+date);

if (proSheet.getRange(index,13).getValue() == "No" ||
proSheet.getRange(index,13).getValue() == "" ||
proSheet.getRange(index,13).getValue() == "CHECK" ||
proSheet.getRange(index,13).getValue() == "Silent") {

var checkClientID = proSheet.getRange(index,6).getValue();
var textFinder = sfSheet.createTextFinder(checkClientID).matchEntireCell(true).findNext();

if (textFinder == null) {continue;}

var indexSF = textFinder.getRow();


if (sfSheet.getRange(indexSF,2).getDisplayValue() == "Yes") {

if (proSheet.getRange(index,13).getValue() == "Silent"){
continue;
}

proSheet.getRange(index,13, cidCount).setValue("CHECK");
proSheet.getRange(index,13, cidCount).setNote(sfSheet.getRange(indexSF,4).getDisplayValue());
date2 = Utilities.formatDate(new Date(sfSheet.getRange(indexSF,3).getDisplayValue()),tz,"yyyy-MM-dd");
proSheet.getRange(index,14, cidCount).setValue(date2);

} else {

if (proSheet.getRange(index,13).getValue() == "No"){
continue;
}

proSheet.getRange(index,13, cidCount).setValue("No");
proSheet.getRange(index,14, cidCount).setValue('=if(OR(M'+index+'="No",M'+index+'=""),"","DATE?")');
}
}
}

} else {
ui.alert("Script cancelled.");
}
}

function menuItemRecovered() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var row = 3;
var copySheet = ss.getSheetByName('In_Progress');
var pasteSheet = ss.getSheetByName('Recovered');
var count1 = 0; var count2 = 0;
var lr = copySheet.getDataRange().getNumRows()-1;
for (i = 0; i < lr; i++){
var isAction = copySheet.getRange(row+i,19).getValue();
count1++;
Logger.log(isAction);
if (isAction == "Recovered"){
var index = pasteSheet.getDataRange().getNumRows();
var source = copySheet.getRange(row+i,1,1,5);
pasteSheet.insertRowAfter(index);
var destination = pasteSheet.getRange(index+1,1,1,5);
source.copyTo(destination, {contentsOnly:true});
count2++;
copySheet.deleteRows(row+i,1);
i--;
Logger.log(count2);
}
}
Logger.log(count2);
if (count2 == 0){
Logger.log(count2);
SpreadsheetApp.getUi()
.alert('Checked '+count1+' data entries, found no entries marked as "Recovered".');
}
else if (count2 > 0){
Logger.log(count2);
SpreadsheetApp.getUi()
.alert('Checked '+count1+' data entries, moved '+count2+' to Recovered sheet.');
}
}

function menuItemVoids() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var row = 3;
var copySheet = ss.getSheetByName('In_Progress');
var pasteSheet = ss.getSheetByName('Voids');
var count1 = 0; var count2 = 0;
var lr = copySheet.getDataRange().getNumRows()-1;
for (i = 0; i < lr; i++){
var isAction = copySheet.getRange(row+i,19).getValue();
count1++;
Logger.log(isAction);
if (isAction == "Void"){
var index = pasteSheet.getDataRange().getNumRows();
var source = copySheet.getRange(row+i,1,1,5);
var sourceNote = copySheet.getRange(row+i,25).getValue();
pasteSheet.insertRowAfter(index);
var destination = pasteSheet.getRange(index+1,1,1,5);
source.copyTo(destination, {contentsOnly:true});
pasteSheet.getRange(index+1,6).setValue(sourceNote);
count2++;
copySheet.deleteRows(row+i,1);
i--;
Logger.log(count2);
}
}
Logger.log(count2);
if (count2 == 0){
Logger.log(count2);
SpreadsheetApp.getUi()
.alert('Checked '+count1+' data entries, found no entries marked as "Void".');
}
else if (count2 > 0){
Logger.log(count2);
SpreadsheetApp.getUi()
.alert('Checked '+count1+' data entries, moved '+count2+' to Voids sheet.');
}
}

function menuItemVoids2() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var copySheet = ss.getSheetByName('In_Progress');
var pasteSheet = ss.getSheetByName('Voids');
var row = 3; var count1 = 0; var count2 = 0;

var lr = copySheet.getDataRange().getNumRows();
var index = pasteSheet.getDataRange().getNumRows(); 

var data1 = []; var data2 = []; var rowIndex = [];
var isAction = "";

for (i = 0; i < lr-1; i++){
isAction = copySheet.getRange(row+i,19).getValue();
count1++;
if (isAction == "Void"){
data1[count2] = copySheet.getRange(row+i,1,1,5).getDisplayValues();
data2[count2] = copySheet.getRange(row+i,25).getValue();
rowIndex[count2] = row+i;
count2++;
}
}

pasteSheet.insertRowsAfter(index,count2);

for (j = 0; j < count2; j++){
pasteSheet.getRange(index+j+1,1,1,5).setValues(data1[j]);
pasteSheet.getRange(index+j+1,6).setValue(data2[j]);
}

var rowsToDelete = rowIndex
.map(row => Math.floor(row))
.filter(row => row >= 1 && row <= copySheet.getMaxRows())
.sort((a, b) => b - a);

rowsToDelete.forEach(row => {
copySheet.deleteRow(row);
})

if (count2 == 0){
SpreadsheetApp.getUi()
.alert('Checked '+count1+' data entries, found no entries marked as "Void".');
}
else if (count2 > 0){
SpreadsheetApp.getUi()
.alert('Checked '+count1+' data entries, moved '+count2+' to Voids sheet.');
}
}

function menuItemRetry() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var row = 3;
var copySheet = ss.getSheetByName('In_Progress');
var pasteSheet = ss.getSheetByName('Retry_Template');
var tz = ss.getSpreadsheetTimeZone();

var count1 = 0; var count2 = 0; var count3 = 0;
var lr = copySheet.getDataRange().getNumRows()-2;

var output = [];
var temp1 = "";
var temp2 = "";
var temp3 = "";

for (i = 0; i < lr; i++){
var isAction = copySheet.getRange(row+i,19).getValue();
count1++;
if (isAction == "Retry"){
temp1 = copySheet.getRange(row+i,2,1,1).getValue();
var dt = Utilities.formatDate(new Date(copySheet.getRange(row+i,4,1,1).getValue()),tz,"MM/yy");
temp2 = 'RA fees '+dt;
var dt2 = new Date();
var date = Utilities.formatDate(dt2,tz, "yyyyMMdd");
var checkEnt1 = copySheet.getRange(row+i,2,1,1).getValue();
var checkEnt2 = copySheet.getRange(row+i+1,2,1,1).getValue();
if (checkEnt1 == checkEnt2){
count3++;
dt2.setDate(dt2.getDate()+count3);
var date2 = Utilities.formatDate(dt2,tz, "yyyyMMdd");
temp3 = date2;
}
else {
temp3 = date;
count3 = 0;
}
copySheet.getRange(row+i,19).setValue("Pending (retry)");
output[count2] = '{"id": "'+temp1+'","name": "'+temp2+'","date": "'+temp3+'"}';
count2++;
Logger.log(output);
}
}
if (count2 == 0){
SpreadsheetApp.getUi()
.alert('Checked '+count1+' data entries, found no entries marked as "Retry".');
}
else if (count2 > 0){
SpreadsheetApp.getUi()
.alert("["+output+"]");
}
}

function menuItemMainSort() { // Upload_Queue_B + Upload_Queue_A -> In_Review -> In_Progress (full auto)
const DELIMITER = ";"; // <---- sets the default delimiter to semicolon
var ss = SpreadsheetApp.getActiveSpreadsheet(); // <---- grabs the current spreadsheet
var row = 3; // <---- sets default Row index for review (skips first two rows)
var stmSheet = ss.getSheetByName('Upload_Queue_B'); // <---- grabs sheet named "Upload_Queue_B"
var dbmSheet = ss.getSheetByName('Upload_Queue_A'); // <---- grabs sheet named "Upload_Queue_A"
var revSheet = ss.getSheetByName('In_Review'); // <---- grabs sheet named "In_Review"
var proSheet = ss.getSheetByName('In_Progress'); // <---- grabs sheet named "In_Progress"

var lastRowSTM = stmSheet.getLastRow(); // <---- collects last row index from "Upload_Queue_B" sheet
var lastRowDBM = dbmSheet.getLastRow(); // <---- collects last row index from "Upload_Queue_A" sheet

var pb = 0; // <---- default value for "Progress Bar" (row "2" in every relevant sheet that this script interracts with)
// 0.############### ------ basic number format that that we get from API


proSheet.getRange(1,2).setValue(pb); // <---- sets the default vlaue of the progress bar




//------------------------------------------------------------------------------ "STM Upload" sheet formatting

if (stmSheet.getRange("B3:K").isBlank()){ // <---- this section is fast enough not to need an "else" option :)))

stmSheet.getRange(3,1,lastRowSTM-2,1).splitTextToColumns(DELIMITER); // <---- splits the raw data in Upload_Queue_B sheet to separate columns based on the set delimiter
SpreadsheetApp.flush();
Utilities.sleep(100);

stmSheet.getRange('I:I').createTextFinder("'").useRegularExpression(true).replaceAllWith(""); // <---- removes the leftover apostrophe from the email column after splitting

stmSheet.insertColumnsAfter(11,4); // <---- adds two temporary columns for data re-formatting and adds relevant formulas
stmSheet.getRange(3,12).setFormula('=left(D3,10)');
stmSheet.getRange(3,13).setFormula('=E3/100');
stmSheet.getRange(3,14).setFormula('=J3/100');
stmSheet.getRange(3,15).setFormula('=left(K3,10)');

SpreadsheetApp.flush();
Utilities.sleep(100);

stmSheet.getRange(3,12,1,4).autoFill(stmSheet.getRange(3,12,lastRowSTM-2,4), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES); // <--- "autofills" formulas to all cells below

SpreadsheetApp.flush();
Utilities.sleep(100);


var sourceSTM = stmSheet.getRange(3,12,lastRowSTM-2,2);
var destinationSTM = stmSheet.getRange(3,4,lastRowSTM-2,2);
sourceSTM.copyTo(destinationSTM, {contentsOnly:true}); // copies data from the first 2 temporary columns to the main data columns and overwrites it

var sourceSTM2 = stmSheet.getRange(3,14,lastRowSTM-2,2);
var destinationSTM2 = stmSheet.getRange(3,10,lastRowSTM-2,2);
sourceSTM2.copyTo(destinationSTM2, {contentsOnly:true}); // copies data from the last 2 temporary columns to the main data columns and overwrites it

SpreadsheetApp.flush();
Utilities.sleep(100);


stmSheet.deleteColumns(12,4); // <---- deletes the temporary columns
stmSheet.getRange(3,5,lastRowSTM-2,1).setNumberFormat('_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'); // <---- sets the "amount" column to "accounting" format
stmSheet.getRange(3,10,lastRowSTM-2,1).setNumberFormat('_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'); // <---- sets the "Current Balance" column to "accounting" format

stmSheet.getRange(3,1,lastRowSTM-2,11).sort({column: 4, ascending: true}); // <---- sorts the entire table based on "date" column (ascending)

SpreadsheetApp.flush();
Utilities.sleep(100);

var dateCheck = stmSheet.getRange(3,4).getDisplayValue(); // <---- section to check if the date is the same in every row
var dateFinder = stmSheet.createTextFinder(dateCheck);
var dateIndex = dateFinder.findAll();
var rowDate = dateIndex[dateIndex.length - 1].getRow();

if (rowDate != lastRowSTM){ stmSheet.deleteRows(rowDate+1,lastRowSTM-rowDate); } // <---- deletes all rows after the last one with the matching date (COMMENT THIS IF AN ALL-TIME REPORT IS BEING RUN)
proSheet.getRange(1,2).setValue(0.01);

}

//--------------------------------------------------------------

if (dbmSheet.getRange("B3:D").isBlank()){

dbmSheet.getRange(3,1,lastRowDBM-2,1).splitTextToColumns(DELIMITER);

SpreadsheetApp.flush();
Utilities.sleep(100);

dbmSheet.getRange('E:E').setNumberFormat('@');
dbmSheet.getRange('E:E').createTextFinder("'").useRegularExpression(true).replaceAllWith("");

dbmSheet.insertColumnAfter(5);
dbmSheet.getRange(3,6).setFormula('=left(B3,3)');

SpreadsheetApp.flush();
Utilities.sleep(100);

dbmSheet.getRange(3,6,1,1).autoFill(dbmSheet.getRange(3,6,lastRowDBM-2,1), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

SpreadsheetApp.flush();
Utilities.sleep(100);


var sourceDBM = dbmSheet.getRange(3,6,lastRowDBM-2,1);
var destinationDBM = dbmSheet.getRange(3,2,lastRowDBM-2,1);
sourceDBM.copyTo(destinationDBM, {contentsOnly:true});

SpreadsheetApp.flush();
Utilities.sleep(100);


dbmSheet.deleteColumn(6);
proSheet.getRange(1,2).setValue(0.02);

} 

//---------------------------------------------------------------

var source = [];
var destination = [];

var lastRowREV = revSheet.getLastRow();

if (revSheet.getDataRange().getNumRows() === 3){

source = stmSheet.getRange(3,1,lastRowSTM-2,11);
destination = revSheet.getRange(3,1,lastRowREV-2,11);
source.copyTo(destination, {contentsOnly:true});

lastRowREV = revSheet.getLastRow();

revSheet.getRange(3,12,1,6).autoFill(revSheet.getRange(3,12,lastRowREV-2,6), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);


SpreadsheetApp.flush();
Utilities.sleep(200);


revSheet.getRange(3,4,lastRowREV-2,1).setNumberFormat('yyyy"-"mm"-"dd');
revSheet.getRange(3,11,lastRowREV-2,1).setNumberFormat('yyyy"-"mm"-"dd');

SpreadsheetApp.flush();
Utilities.sleep(200);
}

lastRowREV = revSheet.getLastRow();

proSheet.getRange(1,2).setValue(0.03);

//----------------------------------------------------------------


var count = 0;
var actionCount = 0;

var isEmptyRev = revSheet.getRange(row,2).getValue();
var isEmptySTM = stmSheet.getRange(row,2).getValue();
var isEmptyDBM = dbmSheet.getRange(row,2).getValue();

var rowCountRev = revSheet.getDataRange().getNumRows();
var rowCountSTM = stmSheet.getDataRange().getNumRows();
var rowCountDBM = dbmSheet.getDataRange().getNumRows();

var checkEntityID = "";
var checkInProgress = "";
var checkSTM = "";
var checkVoids = "";

var textFinder = "";
var searchRow = 0;
var index = 0;

revSheet.getRange(3,1,rowCountRev-2,16).setFontColor("black");

for (i = 0; i < rowCountRev-2; i++){
checkInProgress = revSheet.getRange(row+i,15).getValue();
checkSTM = revSheet.getRange(row+i,17).getValue();
pb = 3+Math.floor(i/(lastRowREV-2)*97);
proSheet.getRange(1,2).setValue(pb/100);

if (isEmptyRev == "" && rowCountRev < 3){
SpreadsheetApp.getUi().alert('Nothing to sort, check "Upload_Queue_B" and "Upload_Queue_A" sheets for data');
break;
}

if (checkInProgress != "-" && checkSTM == "-"){

checkEntityID = revSheet.getRange(row+i,2).getValue();
checkVoids = revSheet.getRange(row+i,14).getValue();

textFinder = proSheet.createTextFinder(checkEntityID);
searchRow = textFinder.findNext().getRow();
count = proSheet.getRange(searchRow,7).getValue()-1;
index = searchRow+count;

dataCopy(revSheet, proSheet, index, row, i);

if (checkVoids != "-"){
proSheet.getRange(index+1,2).setNote('This entity was found in Voids list');
}

if (revSheet.getRange(row+i,8).getValue().toString().startsWith("gccoffee")) {
proSheet.getRange(index+1,6).setNote("Special request: SEND THE TEMPLATE NOTIFICATION TO example@email.com REGARDLESS OF THE CODE -- they will need to arrange it on their end");
}

}
else if (checkInProgress == "-" && checkSTM == "-"){

checkVoids = revSheet.getRange(row+i,14).getValue();
index = proSheet.getDataRange().getNumRows();

dataCopy(revSheet, proSheet, index, row, i);

if (checkVoids != "-"){
proSheet.getRange(index+1,2).setNote('This entity was found in Voids list');
}
if (revSheet.getRange(row+i,8).getValue().toString().startsWith("gccoffee")) {
proSheet.getRange(index+1,6).setNote("Special request: SEND THE TEMPLATE NOTIFICATION TO example@email.com REGARDLESS OF THE CODE -- they will need to arrange it on their end");
}

}
else { revSheet.getRange(row+i,1,1,17).setFontColor("green"); } 
if (i%20 === 0) { 
SpreadsheetApp.flush();
Utilities.sleep(200);
}


}

proSheet.getRange(1,2).setValue(1);

//-------------------------------------------------------------
while (rowCountSTM > 1){
rowCountSTM = stmSheet.getDataRange().getNumRows();
if (rowCountSTM === 3 && isEmptySTM != ""){
stmSheet.getRange(3,1,1,11).clearContent(); 
break;
}
else if (rowCountSTM === 3 && isEmptySTM == ""){
break;
}
else{
stmSheet.deleteRows(3,rowCountSTM-3);
}
}

//-------------------------------------------------------------
while (rowCountDBM > 1){
rowCountDBM = dbmSheet.getDataRange().getNumRows();
if (rowCountDBM === 3 && isEmptyDBM != ""){
dbmSheet.getRange(3,1,1,5).clearContent(); 
break;
}
else if (rowCountDBM === 3 && isEmptyDBM == ""){
break;
}
else{
dbmSheet.deleteRows(3,rowCountDBM-3);
}
}

//-------------------------------------------------------------
while (rowCountRev > 1){
rowCountRev = revSheet.getDataRange().getNumRows();
if (rowCountRev === 3 && isEmptyRev != ""){
revSheet.getRange(3,1,1,11).clearContent();
revSheet.getRange(3,1,1,17).setFontColor("black");
actionCount++; 
break;
}
else if (rowCountRev === 3 && isEmptyRev == ""){
break;
}
else{
revSheet.deleteRows(3,rowCountRev-3);
actionCount++;
}
}
if (actionCount != 0){
SpreadsheetApp.getUi().alert('New data sorting completed!');
}

}

function dataCopy(revSheet, proSheet, index, row, i) {

var source = [];
var destination = [];

var source2 = [];
var destination2 = [];

var link = "";

source = revSheet.getRange(row+i,1,1,5);
proSheet.insertRowAfter(index);
destination = proSheet.getRange(index+1,1,1,5);
source.copyTo(destination, {contentsOnly:true});

proSheet.getRange(index+1,19).setValue('Review (new)');

source2 = proSheet.getRange(index,7,1,8);
destination2 = proSheet.getRange(index+1,7,1,8);
source2.copyTo(destination2);

proSheet.getRange(index+1,6).setValue(revSheet.getRange(row+i,8).getValue());
proSheet.getRange(index+1,9).setValue(revSheet.getRange(row+i,12).getValue());
proSheet.getRange(index+1,15).setValue(revSheet.getRange(row+i,6).getValue());
proSheet.getRange(index+1,16).setValue(revSheet.getRange(row+i,9).getValue());
proSheet.getRange(index+1,18).setValue(revSheet.getRange(row+i,13).getValue());

var dCheck1 = new Date(revSheet.getRange(row+i,11).getValue());
var dCheck2 = new Date();
dCheck2.setHours(dCheck2.getHours() - 72);


if (+dCheck1 < +dCheck2 && revSheet.getRange(row+i,10).getValue() > 0.02) {
proSheet.getRange(index+1,11).setValue("SUS");
proSheet.getRange(index+1,12).setValue(revSheet.getRange(row+i,10).getDisplayValue());


} else {
proSheet.getRange(index+1,11).setValue("No");
proSheet.getRange(index+1,12).setFormula('=if(K'+(index+1)+'="Yes","AMOUNT?","-")')
}


link = revSheet.getRange(row+i,7).getValue();
proSheet.getRange(index+1,17).setFormula('=HYPERLINK("https://example-portal.com/'+link+'/test","Withdrawals")');

}

function menuItemEmailTemplate() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var proSheet = ss.getSheetByName('In_Progress');
var emSheet = ss.getSheetByName('Email_Template_DO_NOT_EDIT');
var legSheet = ss.getSheetByName('Legend_DO_NOT_EDIT');

if (proSheet.getCurrentCell().getColumn() != 6 || proSheet.getCurrentCell().getRow() < 3){
SpreadsheetApp.getUi()
.alert('Make sure you have something selected in the "In_Progress" sheet "SF Client_ID" column to use this script');
return;
}

emSheet.getRange(2,2).setValue(proSheet.getCurrentCell().getDisplayValue());

var clientId = emSheet.getRange(2,2).getDisplayValue();
var email = emSheet.getRange(4,2).getDisplayValue();
var entity = emSheet.getRange(19,5).getDisplayValue()
var code = emSheet.getRange(16,5).getDisplayValue();
var textFinder = legSheet.createTextFinder(code);
var searchCode = textFinder.findNext().getRow();
var codeDesc = legSheet.getRange(searchCode,2).getDisplayValue();
var revelUrl = emSheet.getRange(12,5).getDisplayValue();
var revelEst = emSheet.getRange(13,5).getDisplayValue();
var dba = emSheet.getRange(14,5).getDisplayValue();
var months = emSheet.getRange(15,5).getDisplayValue();
var totalUnpaid = emSheet.getRange(17,5).getDisplayValue();
var totalHeld = emSheet.getRange(18,5).getDisplayValue();
var lastFourBan = "******"+emSheet.getRange(20,5).getDisplayValue();

var howTo1 = legSheet.getRange(searchCode,5).getDisplayValue();
var howTo2 = legSheet.getRange(searchCode,6).getDisplayValue();
var howTo3 = legSheet.getRange(searchCode,7).getDisplayValue();
var howTo4 = legSheet.getRange(searchCode,8).getDisplayValue();

var msgUnpaid = "";
var msgHowTo = "";

if (code == "R01" || code == "R09"){
msgHowTo = '<br><b>'+howTo1+'</b><br>'
} else if (code == "R02" || code == "R03" || code == "R04") {
msgHowTo = '<br>'+howTo1+'<br><br><a href="'+howTo2+'">'+howTo2+'</a><br><br>'+howTo3+'<br>'
} else if (code == "R07" || code == "R08") {
msgHowTo = '<br>'+howTo1+'<br><br>'+howTo2+'<br>'
} else if (code == "R16") {
msgHowTo = '<br>'+howTo1+'<br><br><a href="'+howTo2+'">'+howTo2+'</a><br>'
} else if (code == "R10" || code == "R29") {
msgHowTo = '<br>'+howTo1+'<br><br>'+howTo2+'<br><a href="'+howTo3+'">'+howTo3+'</a><br><br>'+howTo4+'<br>'
}

if (emSheet.getRange(18,5).getValue() > 0){
msgUnpaid = "<br>Please also be aware that due to this return code, the system automatically puts your daily payouts on hold and as a result, the account currently has"+totalHeld+" held in account balance."
} else {msgUnpaid = "";}



var htmlOutput = HtmlService
.createHtmlOutput('<p><b><i>SOME_BASIC_HTML_TEMPLATE_HERE</i></b> '+clientId+'<br><b><i>USING_THE_VARIABLES_LISTED</i></b> '+entity+'</p><hr>')
.setSandboxMode(HtmlService.SandboxMode.IFRAME)
.setWidth(980)
.setHeight(1000);

SpreadsheetApp.getUi().showModelessDialog(htmlOutput, 'Case Email Output');
}

function menuItemStmStatus() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var proSheet = ss.getSheetByName('In_Progress');
var row = 3;
var rows = proSheet.getLastRow()-2;
var range = proSheet.getRange(row,1,rows,1).getValues();

for (i = 0; i < rows; i++){
temp = range[i];
range[i] = '{"id": "'+temp+'"}';
}

Browser.msgBox("test","["+range+"]\\n",Browser.Buttons.OK);

}