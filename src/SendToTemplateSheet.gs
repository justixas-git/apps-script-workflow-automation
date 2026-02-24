function onOpen() {
var ui = SpreadsheetApp.getUi();

ui.createMenu('Template Workflow Tools')
.addSubMenu(ui.createMenu('Submitting new entries (boarding team)')
.addItem('Send selected row(s) to template sheet(s)', 'copyInfo'))
.addSeparator()
.addSubMenu(ui.createMenu('Admin Actions')
.addItem('Generate CSV files for FTP', 'sheetToCsv')
.addItem('FPBB_CSV folder cleanup', 'clearFolder'))
.addToUi();

}

function copyInfo() {
var tz = "GMT-4";
var ts_format = "yyyy-MM-dd"; // Timestamp Format
var date = Utilities.formatDate(new Date(), tz, ts_format);
var ss = SpreadsheetApp.getActiveSpreadsheet();
var copySheetName = ss.getSheetName();
var copySheet = ss.getSheetByName(copySheetName);
var currentCell = SpreadsheetApp.getCurrentCell();
var currentRow = currentCell.getRow();
var currentColumn = currentCell.getColumn();
var tsys = 0; tsysCID = [];
var fdNorth = 0; fdNorthCID = [];
var fdOmaha = 0; fdOmahaCID = [];
var elavon = 0; elavonCID = [];
var chase = 0; chaseCID = [];
var heartland = 0; heartlandCID = [];
var vantiv = 0; vantivCID = [];
var found;
var counter = 0;
var counterValue = ss.getRange("H1").getValue();
function isInArray(value, array) {
return array.indexOf(value) > -1;
}
function flatten(arrayOfArrays){
return [].concat.apply([], arrayOfArrays);
}

for (i=0; i<ss.getActiveRange().getNumRows();i++){
var newRow = currentRow + i;
var processor = currentCell.offset(i,-1).getValue();
var checkCID = copySheet.getRange(newRow,currentColumn).getValue();
if (currentCell.offset(i,-1).getValue() == "Template_F" && 
currentRow > 9 && 
currentColumn == 8 && 
currentCell.offset(i,-5).getValue() == "" && 
currentCell.getValue() != "" && 
currentCell.offset(i,-7).getValue() != "Testing"){
found = 0;
var pasteSheet = ss.getSheetByName("Template_F");
var lr = pasteSheet.getDataRange().getLastRow();
var checkRange = pasteSheet.getRange(2,1,lr).getValues();
var checkValue = copySheet.getRange(newRow,currentColumn).getValue();
if (isInArray(checkValue, flatten(checkRange)) == true){
found++;
Browser.msgBox(checkValue+" is already in the "+processor+" sheet.");
}
if (found == 0){
// get source range
var source = copySheet.getRange(newRow,currentColumn,1,37);
// get destination range
var destination = pasteSheet.getRange(pasteSheet.getLastRow()+1,1,1,37);
// copy values to destination range
source.copyTo(destination, {contentsOnly:true});
currentCell.offset(i,-2).setValue(date);
tsys++;
tsysCID[tsys-1] = " "+checkValue;
counter++;
}

}else if (currentCell.offset(i,-1).getValue() == "Template_C" && 
currentRow > 9 && 
currentColumn == 8 && 
currentCell.offset(i,-5).getValue() == "" && 
currentCell.getValue() != "" && 
currentCell.offset(i,-7).getValue() != "Testing"){
found = 0;
var pasteSheet = ss.getSheetByName("Template_C");
var lr = pasteSheet.getDataRange().getLastRow();
var checkRange = pasteSheet.getRange(2,2,lr).getValues();
var checkValue = copySheet.getRange(newRow,currentColumn).getValue();
if (isInArray(checkValue, flatten(checkRange)) == true){
found++;
Browser.msgBox(checkValue+" is already in the "+processor+" sheet.");
}
if (found == 0){
//var source = copySheet.getRange(newRow,currentColumn,1,35);
//var destination = pasteSheet.getRange(pasteSheet.getLastRow()+1,2,1,35);
//source.copyTo(destination, {contentsOnly:true});
pasteRow = pasteSheet.getLastRow()+1;
var source1 = copySheet.getRange(newRow,currentColumn,1,7);
var destination1 = pasteSheet.getRange(pasteRow,2,1,7);
source1.copyTo(destination1, {contentsOnly:true});
var source2 = copySheet.getRange(newRow,currentColumn+7,1,2);
var destination2 = pasteSheet.getRange(pasteRow,10,1,2);
source2.copyTo(destination2, {contentsOnly:true});
var source3 = copySheet.getRange(newRow,currentColumn+9,1,20);
var destination3 = pasteSheet.getRange(pasteRow,13,1,20);
source3.copyTo(destination3, {contentsOnly:true});
var source4 = copySheet.getRange(newRow,currentColumn+32,1,3);
var destination4 = pasteSheet.getRange(pasteRow,47,1,3);
source4.copyTo(destination4, {contentsOnly:true});
if (pasteSheet.getRange(pasteRow,26).getValue() == "no") pasteSheet.getRange(pasteRow,26).setValue("false");
if (pasteSheet.getRange(pasteRow,27).getValue() == "yes") pasteSheet.getRange(pasteRow,27).setValue("true");

if (pasteSheet.getRange(pasteRow,32).getValue() == "0") {
pasteSheet.getRange(pasteRow,33).setValue("None");
pasteSheet.getRange(pasteRow,32).setValue("");
}
else {
pasteSheet.getRange(pasteRow,33).setValue("HPC");
pasteSheet.getRange(pasteRow,34).setValue("https://example-checkout.com");
pasteSheet.getRange(pasteRow,32).setValue("");
}

pasteSheet.getRange(pasteRow,1).setValue("REVEL1");
pasteSheet.getRange(pasteRow,45).setValue("false");
pasteSheet.getRange(pasteRow,46).setValue("false");
pasteSheet.getRange(pasteRow,50).setValue("984");
pasteSheet.getRange(pasteRow,51).setValue("Revel");

currentCell.offset(i,-2).setValue(date);
fdNorth++;
fdNorthCID[fdNorth-1] = " "+checkValue;
counter++;
}

}else if (currentCell.offset(i,-1).getValue() == "Template_D" && 
currentRow > 9 && 
currentColumn == 8 && 
currentCell.offset(i,-5).getValue() == "" && 
currentCell.getValue() != "" && 
currentCell.offset(i,-7).getValue() != "Testing"){
found = 0;
var pasteSheet = ss.getSheetByName("Template_D");
var lr = pasteSheet.getDataRange().getLastRow();
var checkRange = pasteSheet.getRange(2,1,lr).getValues();
var checkValue = copySheet.getRange(newRow,currentColumn).getValue();
if (isInArray(checkValue, flatten(checkRange)) == true){
found++;
Browser.msgBox(checkValue+" is already in the "+processor+" sheet.");
}
if (found == 0){
var source = copySheet.getRange(newRow,currentColumn,1,29);
var destination = pasteSheet.getRange(pasteSheet.getLastRow()+1,1,1,29);
source.copyTo(destination, {contentsOnly:true});
currentCell.offset(i,-2).setValue(date);
fdOmaha++;
fdOmahaCID[fdOmaha-1] = " "+checkValue;
counter++;
}

}else if (currentCell.offset(i,-1).getValue() == "Template_B" && 
currentRow > 9 && 
currentColumn == 8 && 
currentCell.offset(i,-5).getValue() == "" && 
currentCell.getValue() != "" && 
currentCell.offset(i,-7).getValue() != "Testing"){
found = 0;
var pasteSheet = ss.getSheetByName("Template_B");
var lr = pasteSheet.getDataRange().getLastRow();
var checkRange = pasteSheet.getRange(2,1,lr).getValues();
var checkValue = copySheet.getRange(newRow,currentColumn).getValue();
if (isInArray(checkValue, flatten(checkRange)) == true){
found++;
Browser.msgBox(checkValue+" is already in the "+processor+" sheet.");
}
if (found == 0){
var source = copySheet.getRange(newRow,currentColumn,1,29);
var destination = pasteSheet.getRange(pasteSheet.getLastRow()+1,1,1,29);
source.copyTo(destination, {contentsOnly:true});
currentCell.offset(i,-2).setValue(date);
elavon++;
elavonCID[elavon-1] = " "+checkValue;
counter++;
}

}else if (currentCell.offset(i,-1).getValue() == "Template_E" && 
currentRow > 9 && 
currentColumn == 8 && 
currentCell.offset(i,-5).getValue() == "" && 
currentCell.getValue() != "" && 
currentCell.offset(i,-7).getValue() != "Testing"){
found = 0;
var pasteSheet = ss.getSheetByName("Template_E");
var lr = pasteSheet.getDataRange().getLastRow();
var checkRange = pasteSheet.getRange(2,1,lr).getValues();
var checkValue = copySheet.getRange(newRow,currentColumn).getValue();
if (isInArray(checkValue, flatten(checkRange)) == true){
found++;
Browser.msgBox(checkValue+" is already in the "+processor+" sheet.");
}
if (found == 0){
var source = copySheet.getRange(newRow,currentColumn,1,28);
var destination = pasteSheet.getRange(pasteSheet.getLastRow()+1,1,1,28);
source.copyTo(destination, {contentsOnly:true});
currentCell.offset(i,-2).setValue(date);
heartland++;
heartlandCID[heartland-1] = " "+checkValue;
counter++;
}

}else if (currentCell.offset(i,-1).getValue() == "Template_A" && 
currentRow > 9 && 
currentColumn == 8 && 
currentCell.offset(i,-5).getValue() == "" && 
currentCell.getValue() != "" && 
currentCell.offset(i,-7).getValue() != "Testing"){
found = 0;
var pasteSheet = ss.getSheetByName("Template_A");
var lr = pasteSheet.getDataRange().getLastRow();
var checkRange = pasteSheet.getRange(2,1,lr).getValues();
var checkValue = copySheet.getRange(newRow,currentColumn).getValue();
if (isInArray(checkValue, flatten(checkRange)) == true){
found++;
Browser.msgBox(checkValue+" is already in the "+processor+" sheet.");
}
if (found == 0){
var source = copySheet.getRange(newRow,currentColumn,1,30);
var destination = pasteSheet.getRange(pasteSheet.getLastRow()+1,1,1,30);
source.copyTo(destination, {contentsOnly:true});
currentCell.offset(i,-2).setValue(date);
chase++;
chaseCID[chase-1] = " "+checkValue;
counter++;
}

}else if (currentCell.offset(i,-1).getValue() == "Template_G" && 
currentRow > 9 && 
currentColumn == 8 && 
currentCell.offset(i,-5).getValue() == "" && 
currentCell.getValue() != "" && 
currentCell.offset(i,-7).getValue() != "Testing"){
found = 0;
var pasteSheet = ss.getSheetByName("Template_G");
var lr = pasteSheet.getDataRange().getLastRow();
var checkRange = pasteSheet.getRange(2,1,lr).getValues();
var checkValue = copySheet.getRange(newRow,currentColumn).getValue();
if (isInArray(checkValue, flatten(checkRange)) == true){
found++;
Browser.msgBox(checkValue+" is already in the "+processor+" sheet.");
}
if (found == 0){
var source = copySheet.getRange(newRow,currentColumn,1,33);
var destination = pasteSheet.getRange(pasteSheet.getLastRow()+1,1,1,33);
source.copyTo(destination, {contentsOnly:true});
currentCell.offset(i,-2).setValue(date);
vantiv++;
vantivCID[vantiv-1] = " "+checkValue;
counter++;
}

}else if(currentRow < 9){
Browser.msgBox("Good try, but this does not work on the first 9 rows of the sheet :)) \\nTry again.");
break;
}else if(currentCell.getValue() == "" && currentRow > 9 && currentColumn == 8){
Browser.msgBox("Sorry, looks like you have nothing entered in H"+newRow+". :(");
}else if(currentCell.offset(i,-1).getValue() == "" && currentRow > 9 && currentColumn == 8 && currentCell.offset(i,-5).getValue() == "" && currentCell.getValue() != ""){
Browser.msgBox("No processor selected in column G for "+checkCID+" in H"+newRow+".");
}else if(currentCell.offset(i,-1).getValue() != "" && currentRow > 9 && currentColumn == 8 && currentCell.offset(i,-5).getValue() == "" && currentCell.offset(i,-7).getValue() == "Testing"){
Browser.msgBox("Looks like "+checkCID+" in H"+newRow+" is already being tested.");
}else if(currentCell.offset(i,-1).getValue() != "" && currentRow > 9 && currentColumn == 8 && currentCell.offset(i,-5).getValue() != ""){
Browser.msgBox("Looks like "+checkCID+" in H"+newRow+" already has an active FreedomPay account (Column C): "+currentCell.offset(i,-5).getValue());
}
else{
Browser.msgBox("Make sure you have the desired client_id's selected in column H before clicking this button.");
break;
}
}
if (i>0){
Browser.msgBox("Submission results",
"| "+tsys+ " x ___TSYS_____ |:"+tsysCID+"\\n| "
+fdNorth+ " x _FD_NORTH_ |:"+fdNorthCID+"\\n| "
+fdOmaha+ " x _FD_OMAHA_ |:"+fdOmahaCID+"\\n| "
+heartland+" x _HEARTLAND |:"+heartlandCID+"\\n| "
+elavon+ " x __ELAVON___ |:"+elavonCID+"\\n| "
+chase+ " x __CHASE____ |:"+chaseCID+"\\n| "
+vantiv+ " x __VANTIV___ |:"+vantivCID, Browser.Buttons.OK);
}
if (counter>0){
var num1 = counter;
Logger.log(num1);
var num2 = parseInt(counterValue);
Logger.log(num2);
var num3 = num1+num2;
Logger.log(num3);
ss.getRange("H1").setValue(num3);
}


// clear source values
//source.clearContent();
}
