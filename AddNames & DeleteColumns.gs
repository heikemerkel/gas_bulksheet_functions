/////////////////////////////////////////////////////////////////////////////
//// BELOW FUNCTIONS FILL IN MISSING NAMES FROM TOAD IN THE TABS ////////////
//// AND DELETES OLD LABOR PROJECTION COLUMNS ///////////////////////////////
/////////////////////////////////////////////////////////////////////////////

function fillNamesFromTOAD() {
  const alasql = AlaSQLGS.load(); //load sql library

  //pull values from TOAD labor GS
  var toad_labor_sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1SxvjuEMt-u9QucmX9axsRB6kBLUOopZGPWKbHt9Vmjo/edit#gid=2091278272'); //TOAD Labor GS
  var toad_labor_tab = toad_labor_sheet.getSheetByName('Sheet1');
  //var values_labor = [];
  var values_labor = toad_labor_tab.getRange("A3:G").getValues();
  //console.log("values labor: ",values_labor)

  //go through all sheets taking names that are in the sheet and look for missing ones in the TOAD query and add to end
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

    for (i=2; i<sheets.length-2; i++){ //go through all the sheets from sheet 3 (i=2) through the end minus 2 sheets (ACEP Emp & Hours per PR)
        var exclude = sheets[i].getName().indexOf('*ST'); //if tab '*ST' is not found ==> value is -1, include it in the summation, this makes it possible to create subtask tabs that don't get
                                                        // included in summation in sheet Full
        if (-1 == exclude) {
        //get the fund code
        var fund = sheets[i].getRange(3,1).getValue();
        //var actc = sheets[i].getRange(5,1).getValue(); //need to set activity code in a cell somewhere
        // console.log("fund: ", fund);
        // console.log("actc length: ", actc.length);
        // console.log("actc: ", actc);
        
        //get names from TOAD query for this fund
        //if (actc.length == 0) { //uncomment if using activity codes
          var query = "SELECT MATRIX [4] FROM ? WHERE [1] = "+fund
        //   }   //select 'name' where 'fund'
        // else {var query = "SELECT MATRIX [4] FROM ? WHERE [1] = "+fund+" AND [3] ="+actc}  //select 'name' where 'fund' and 'actv code'
        existNamesTOAD = alasql((query), [values_labor]);
        //console.log("query: ", query);
        existNamesTOADR = existNamesTOAD.reduce(function(a, b) {return a.concat(b);}, []);
        console.log("existNamesTOADR: ",existNamesTOADR);
        
        //get names in sheet[i]
        var textFinder = sheets[i].getRange("A9:A").createTextFinder("Health Insurance (Acct 1949)").findAll();
        var rcs = textFinder[0].getRowIndex();
        console.log("rcs: ",rcs);
        var existNamesSheet = sheets[i].getRange(9,1,rcs-9,1).getValues().reduce(function(a, b) {return a.concat(b);}, []); //turn into 1D array
        console.log("existNamesSheet: ",existNamesSheet);
        
        //find names that exist in TOAD that are missing in the sheet
        var missing = missed(existNamesSheet,existNamesTOADR);
        console.log("missing: ",missing);
        
        //add the missing names to the sheet
        for (j=0;j<missing.length;j++) { //loop through the missing names
          var lastRow = sheets[i].getRange(9,1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow(); 
          console.log("lastRow: ",lastRow); 
            if (lastRow > rcs) { //if last row with data is less than row Health Insurance (Acct 1949), insert a row and copy the formulas from row 9
              sheets[i].insertRowsAfter(rcs-1, 1);
              lastRow = sheets[i].getRange(9,1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow(); 
            }
            sheets[i].getRange(9,6,1,13).copyTo(sheets[i].getRange(lastRow+1,6,1,13),SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false); //copy formulas from row 9
            sheets[i].getRange(lastRow+1,1).setValue(missing[j]);
        }

        //delete old projection columns and add new columns to the end
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
}

function missed(existNamesSheet,existNamesTOADR){
  return existNamesTOADR.filter(function(e){return existNamesSheet.indexOf(e) === -1});
}