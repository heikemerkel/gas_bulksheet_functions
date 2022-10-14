////////////////////////////////////////////////////////////////////////////
//// BELOW FUNCTIONS ARE FOR GENERAL OPERATIONS ////////////////////////////
////////////////////////////////////////////////////////////////////////////

function addNewColumns(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (i=1; i<sheets.length-2; i++){
    sheets[i].insertColumnAfter(16);
    sheets[i].getRange(7,17,1,1).setValue('Start date for projections')
  }
}

function deleteColumns(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (i=1; i<sheets.length-2; i++){
    sheets[i].deleteColumn(17);
  }
}

function generateSheetList(){
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

function copyPasteCells(){
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheets = ss.getSheets();
  for (i=2; i<sheets.length-2; i++){
    sheets[i].getRange(3,3,1,1).setValue("=Full!C3")
  }
}

function insertRows(){ //insert more rows in tabs for more personnel and copy formulas along
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


