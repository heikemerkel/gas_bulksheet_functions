////////////////////////////////////////////////////////////////////////////
//// BELOW FUNCTIONS ARE FOR GENERAL OPERATIONS ////////////////////////////
////////////////////////////////////////////////////////////////////////////

function addNewColumnsG(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (i=1; i<sheets.length-2; i++){
    sheets[i].insertColumnAfter(16);
    sheets[i].getRange(7,17,1,1).setValue('Start date for projections')
  }
}

function deleteColumnsG(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (i=1; i<sheets.length-2; i++){
    sheets[i].deleteColumn(17);
  }
}

function generateSheetListG(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheet_names = []
  for (i=0; i< sheets.length; i++) {
    sheet_names.push([sheets[i].getName()])
  }
  //console.log("sheet_names: ", sheet_names);
  var toc_sheet = ss.getSheetByName('Tabs');
  var toc_range = toc_sheet.getRange(1,2,sheet_names.length, sheet_names[0].length)
  toc_range.setValues(sheet_names)
}

function copyPasteCellsG(){
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheets = ss.getSheets();
  for (i=2; i<sheets.length-2; i++){
    sheets[i].getRange(3,3,1,1).setValue("=Full!C3")
  }
}

function updateEncumFA(){ //updates F&A rates for encumbrances using the formula in column 14 of each tab
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();    
  for (i=3; i<sheets.length-3; i++){
    var textFinder = sheets[i].getRange("A9:A").createTextFinder("7-F&A").findAll();
    var rcs = textFinder[0].getRowIndex();
    console.log("rcs: ",rcs);
    sheets[i].getRange(rcs,14).copyTo(sheets[i].getRange(rcs,10), SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
false);
  }
}

function updateColor(){ //updates Grant Total color formats in each sheet (watch for Match Total!)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fullSheet = ss.getSheetByName('Full');
  var sheets = ss.getSheets();    
  for (i=3; i<sheets.length-3; i++){
    var textFinder = sheets[i].getRange("A9:A").createTextFinder("Grand TOTAL").findAll();
    var rcs = textFinder[0].getRowIndex();
    console.log("rcs: ",rcs);
    fullSheet.getRange(44,3).copyTo(sheets[i].getRange(rcs,3), SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);
    fullSheet.getRange(44,3).copyTo(sheets[i].getRange(rcs+1,12), SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);
  }
}

function updateR00(){ //updates formatting for R05 and PR date in row 2+3, col 3
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fullSheet = ss.getSheetByName('Full');
  var sheets = ss.getSheets();    
  for (i=2; i<sheets.length-3; i++){
    var exclude = sheets[i].getName().indexOf('*ST'); //if tab '*ST' is not found ==> value is -1, include it in the summation, this makes it possible to create subtask tabs 
                                                      //that don't get included in summation in sheet Full
    if (-1 == exclude) {
       fullSheet.getRange(2,3,2,1).copyTo(sheets[i].getRange(2,3,2,1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);
    }
  }
}

function insertRowsG(){ //insert more rows in tabs for more personnel and copy formulas along
 var ss = SpreadsheetApp.getActiveSpreadsheet(); 
 var numRows = ss.getSheetByName("Names").getRange("A125").getValue(); // value = number of rows
 var sheets = ss.getSheets();
  for (i=3; i<sheets.length-2; i++){ //start with sheet number 3
    sheets[i].insertRowsAfter(9, numRows)//(start at row 9, how many rows to add?)
    var lCol = sheets[i].getLastColumn(); //how many columns to copy
      for (j=1; j<=numRows; j++) { 
      sheets[i].getRange(9, 6, 1, lCol).copyTo(sheets[i].getRange(9+j,6,1,lCol), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
      //take formulas from row 9 and paste it in all just created rows, one at the time, only PASTE_FORMULA, don't transpose (=false)
      }
  }
}


