//////////////////////////////////////////////////////////////////////////////////////
//// BELOW FUNCTIONS UPDATE PERSONNEL AND NUMBERS IN TAB 'FULL' -- SIMPLE ////////////
//////////////////////////////////////////////////////////////////////////////////////

//TODO: put functions on a trigger to run in the middle of the night either every day or every week

function clearEntriesS(){ //to make sure old numbers don't stay in and only get partially overwritten by new entries
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fullSheet = ss.getSheetByName('Full');
  
  //Clear existing entries in labor:
    var textFinder = fullSheet.getRange("A9:A").createTextFinder("Health Insurance (Acct 1949)").findAll();
    //var rowcolumn = textFinder.findAll().map(r => ({row: r.getRow(), col: r.getColumn()}));
    var rcs = textFinder[0].getRowIndex()-9;
    fullSheet.getRange(9, 1, rcs, 3).clearContent(); //only clear values, not formatting
    // fullSheet.getRange(9, 8, rcs, 3).clearContent();
    fullSheet.getRange(9, 14, rcs).clearContent();
    
  //Clear existing entries in expenditures:
    // var textFinder = fullSheet.getRange("A9:A").createTextFinder("1-Personal Services").findAll(); //find the correct row to start
    // var exp = textFinder[0].getRowIndex();
    // //console.log("exp: ", exp);
    // fullSheet.getRange(exp, 3, 9, 3).clearContent(); //only clear values, not formatting  
}

function fillNamesInFullS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fullSheet = ss.getSheetByName('Full');
  var lCol = fullSheet.getLastColumn(); 
  var sheets = ss.getSheets();
    
    var counter = 9; //start filling names into Full at row 9 
    for (i=2; i<sheets.length-2; i++){ //go through all the sheets from sheet 3 (i=2) through the end minus 2 sheets (ACEP Emp & Hours per PR)
      var exclude = sheets[i].getName().indexOf('*ST'); //if tab '*ST' is not found ==> value is -1, include it in the summation, this makes it possible to create subtask tabs that don't get
                                                        // included in summation in sheet Full
        if (-1 == exclude) {
            for (j=9; j<500; j++){ //go through all the rows: row 9 through row 500 (break when the end is reached, so it doesn't actually go through 500 rows)
            var name = sheets[i].getRange(j,1).getValues(); //and get the value in row 9, column 1
            //console.log("name: ", name);
            if (name=="Health Insurance (Acct 1949)") {break;}  // 
            if (name!="") { //add the name to the list in Full after checking if it already exists
                var textFinder = fullSheet.getRange("A9:A").createTextFinder(name);  //check if the name already exists in sheet Full column A
                var occurrences = textFinder.findAll().map(x => x.getA1Notation());
                if (occurrences.length == 0) { //if the name does not exist ...
                  //Add rows in column A if needed and copy/paste the formula from the previous row:
                  if (fullSheet.getRange(counter,1).getValues()=="Health Insurance (Acct 1949)"){
                      fullSheet.insertRowsAfter(counter-1, 1);
                      fullSheet.getRange(counter-1,1,1,lCol).copyTo(fullSheet.getRange(counter,1,1,lCol),SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
                    }
                  fullSheet.getRange(counter,1).setValue(name); //add name to next empty row in column 1
                  counter+=1;  //increase the counter to put the next value in the next row
                  //console.log("counter: ", counter);
                }
            }
            }
        }
    }
}

function sumSheetS(column, name){ //helper function --> don't run this, it's getting called from another function
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var fullSheet = ss.getSheetByName('Full');
    var sheets = ss.getSheets();
    
    if (name != ""){
        var valued = 0; //reset value to 0
        for (k=2; k<sheets.length-2; k++){ //loop through all the sheet from sheet 3 (k=2) to sum up the values by name and column, first sheet is sheets[0]
            //console.log("k: ", k);
            var exclude = sheets[k].getName().indexOf('*ST'); //if tab '*ST' is not found ==> value is -1, include it in the summation, this makes it possible to create subtask tabs that don't get
                                                              // included in summation in sheet Full
            if (-1 == exclude) {
                var sheetsI = sheets[k].getRange("A9:A").getValues();
                //console.log("sheets: ", sheetsI);
                var textFinder = sheets[k].getRange("A9:A").createTextFinder(name[0]).matchEntireCell(true).findAll();
                //console.log("textfinder:", textFinder.length);
                
                if (textFinder.length !=0){
                  for (l=0;l<textFinder.length;l++){
                  //var same = textFinder[0].getValues();
                  //console.log("textFinder: ", same);
                  var row = textFinder[0].getRow();
                  //console.log("row: ", row);
                  addValue = sheets[k].getRange(row,column).getValue();
                  //console.log("addValue: ", addValue);
                  valued = +valued + +addValue; //the double plus forces numbers (it wants to turn numbers into strings at random occasions)
                  //console.log("valued: ", valued);
                  }
                }
            }    
         }
        if (valued != 0) { //write the value in sheet Full
            var textFinder = fullSheet.getRange("A9:A").createTextFinder(name[0]).matchEntireCell(true).findAll();
            if (textFinder.length !=0){
              var i = textFinder[0].getRow();
              var test = fullSheet.getRange(i,column).setValue(valued);}
         }
     }
}

//Helper functions for each column to prevent timeout issues --> don't run these, getting called from another function
function calculateLaborS(name){
    columnsOfInterest = [2,3,4,5,14]
    columnsOfInterest.map(function(column){return sumSheetS(column,name);})
}

function calculateExpS(name){
        columnsOfInterest = [14]
        columnsOfInterest.map(function(column){return sumSheetS(column,name);})
}

//fill in Numbers for Labor for each name and column (2,3,8,9,10,14) --> these are set up to run as a trigger in the middle of the night once a week
function fillNumbersInFullLaborS() { //sum all the labor numbers from each tab for each name in sheet 'Full'
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var fullSheet = ss.getSheetByName('Full');
    var textFinder1 = fullSheet.getRange("A9:A").createTextFinder("Health Insurance (Acct 1949)").findAll(); //find the correct end row
    var rowy = textFinder1[0].getRowIndex();
    //console.log("rowy: ", rowy);
    var namesFull = fullSheet.getRange(9,1,rowy-8,1).getValues(); //read all the names + Health Insurance into array
    console.log("namesFull: ", namesFull);
    namesFull.map(calculateLaborS);
}

function fillNumbersInFullExpS() { //sum all the exp numbers from each tab for each category in sheet 'Full'
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var fullSheet = ss.getSheetByName('Full');
    var sheets = ss.getSheets();
  
    var textFinder1 = fullSheet.getRange("A9:A").createTextFinder("Health Insurance (Acct 1949)").findAll(); //find the correct end row
    var rowy = textFinder1[0].getRowIndex();
    var expFull = fullSheet.getRange(rowy+2,1,9,1).getValues(); //read the categories from "1-Personal Services" to "7-F&A" into array
    console.log("expFull: ", expFull);
    expFull.map(calculateExpS);
}