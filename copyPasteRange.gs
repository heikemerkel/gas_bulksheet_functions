function copyPasteCells(){ //copy and paste dates and R's from one sheet into all the other sheets, fixed formatting issues
 var ss = SpreadsheetApp.getActiveSpreadsheet(); 
//  var valuesToCopy = ss.getSheetByName("Solar Deployment").getRange("S7:BY8").getValues(); // value = number of rows
 var sheets = ss.getSheets();
  for (i=4; i<sheets.length-2; i++){ //start with sheet number 4
      ss.getSheetByName("UAA-Participant").getRange("S7:BY8").copyTo(sheets[i].getRange("S7:BY8"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      //take values, format, formulas from sheet Solar Deployment and paste it in all the other sheets, PASTE_NORMAL, don't transpose (=false)
  }
}

function deleteColumns(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (i=3; i<sheets.length-2; i++){
    var last = sheets[i].getLastColumn();
    sheets[i].deleteColumn(last);
    // sheets[i].deleteColumn(15);
  }
}

function addNewColumns(){ //add new column at the end with all the correct formatting (date+14, R+1)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (i=3; i<sheets.length-2; i++){
            sheets[i].insertColumnAfter(sheets[i].getMaxColumns());
            sheets[i].getRange(1,sheets[i].getLastColumn(),sheets[i].getLastRow(),1).copyTo(sheets[i].getRange(1,sheets[i].getLastColumn()+1,sheets[i].getLastRow(),1),{formatOnly:true , contentsOnly:true}); //copy formulas and formatting
            var date = new Date(sheets[i].getRange(7,sheets[i].getLastColumn()).getValue()); //format entry as date
            // console.log("date: ",date);
            new Date(date.setDate(date.getDate()+14)); //add 14 days to last date
            // console.log("weDate: ",date);
            sheets[i].getRange(7,sheets[i].getLastColumn()).setValue(date); //put new date in new column
            var r = sheets[i].getRange(8,sheets[i].getLastColumn()).getValue();
            console.log("r: ",r);
            if (r < 26) {sheets[i].getRange(8,sheets[i].getLastColumn()).setValue(r+1); console.log("r<26");} //increase R by 1 if less than R26
            else {sheets[i].getRange(8,sheets[i].getLastColumn()).setValue(1); console.log("r=1");}
  }
}