function updateEncumFA(){ //updates F&A rates for encumbrances using the formula in column 14 of each tab
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();    
  for (i=2; i<sheets.length-3; i++){
    var textFinder = sheets[i].getRange("A9:A").createTextFinder("7-F&A").findAll();
    var rcs = textFinder[0].getRowIndex();
    console.log("rcs: ",rcs);
    sheets[i].getRange(rcs,14).copyTo(sheets[i].getRange(rcs,10), SpreadsheetApp.CopyPasteType.PASTE_FORMULA,false);
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


function copyQueryFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fullSheet = ss.getSheetByName('Full');
  fullSheet.getRange(9,9).setFormula('=iferror(QUERY(\'*ST TOAD Labor\'!$A$1:$G,"select sum(G) where E=\'"&$A9&"\'and A=\'"&$A$2&"\' and not E=\'\' label sum(G)\'\'"),"")');
  fullSheet.getRange(9,8).setFormula('=iferror(QUERY(\'*ST TOAD Labor\'!$A$1:$G,"select sum(F) where E=\'"&$A9&"\'and A=\'"&$A$2&"\' and not E=\'\' label sum(F)\'\'"),"")');
  //setFormula('=query(arrayformula(Master!A:K), "SELECT B, C, D, E, F, G, H, I, J, K where A = \'"& \'Select Your Event\'!A3 &"\' Order By I, J",1)')
  var textFinder = fullSheet.getRange("A9:A").createTextFinder("Health Insurance (Acct 1949)").findAll();
  var rcs = textFinder[0].getRowIndex()-9;
  var fillDownRangeA = fullSheet.getRange(9,9,rcs);
  var fillDownRangeH = fullSheet.getRange(9,8,rcs);
  fullSheet.getRange(9,9).copyTo(fillDownRangeA);//for amount in column 9
  fullSheet.getRange(9,8).copyTo(fillDownRangeH);//for hours in column 8
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

function copyPasteCells(){ //copy and paste dates and R's from one sheet into all the other sheets, fixes formatting issues
 var ss = SpreadsheetApp.getActiveSpreadsheet(); 
//  var valuesToCopy = ss.getSheetByName("Solar Deployment").getRange("S7:BY8").getValues(); // value = number of rows
 var sheets = ss.getSheets();
  for (i=4; i<sheets.length-2; i++){ //start with sheet number 4
      ss.getSheetByName("T-0 ACEP").getRange("S7:BY8").copyTo(sheets[i].getRange("S7:BY8"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      //take values, format, formulas from sheet 'T-0 ACEP' and paste it in all the other sheets, PASTE_NORMAL, don't transpose (=false)
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

function deleteAddColumns() { //use this function to delete old date columns and add new ones at the end
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

    for (i=2; i<sheets.length-2; i++){ //go through all the sheets from sheet 3 (i=2) through the end minus 2 sheets (ACEP Emp & Hours per PR)
      //find the old columns with dates before the current payroll date = getRange(3,3), for every old column is finds and deletes, it will add a new one at the end.
        var lastPayDate = sheets[i].getRange(3,3).getValue(); 
        while ((lastPayDate >= sheets[i].getRange(7,19).getValue()) && (sheets[i].getRange(7,19).getValue() !='')) {
          // console.log("colDate inside: ",sheets[i].getRange(7,19).getValue());
            sheets[i].deleteColumn(19);
            sheets[i].insertColumnAfter(sheets[i].getMaxColumns());
            sheets[i].getRange(1,sheets[i].getLastColumn(),sheets[i].getLastRow(),1).copyTo(sheets[i].getRange(1,sheets[i].getLastColumn()+1,sheets[i].getLastRow(),1),{formatOnly:true , contentsOnly:true}); //copy formulas and formatting
            var date = new Date(sheets[i].getRange(7,sheets[i].getLastColumn()).getValue()); //format entry as date
            // console.log("date: ",date);
            new Date(date.setDate(date.getDate()+14)); //add 14 days to last date
            // console.log("weDate: ",date);
            sheets[i].getRange(7,sheets[i].getLastColumn()).setValue(date); //put new date in new column
            var r = sheets[i].getRange(8,sheets[i].getLastColumn()).getValue();
            if (r < 26) {sheets[i].getRange(8,sheets[i].getLastColumn()).setValue(r+1)} //increase R by 1 if less than R26
            else {sheets[i].getRange(8,sheets[i].getLastColumn()).setValue(1)}
        }
        if (sheets[i].getRange(7,19).getValue() !='') {sheets[i].getRange(6,19).setValue("Payroll projections (end date)")};
    }
}
