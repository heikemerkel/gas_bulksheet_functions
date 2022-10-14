////////////////////////////////////////////////////////////////////////////
//// BELOW FUNCTIONS CREATE THE ORIGINAL BUDGET SHEET WITH ALL THE TABS ////
////////////////////////////////////////////////////////////////////////////

function duplicateTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
 //Duplicate tabs from template 'Grant':
  var tabs = ss.getSheetByName('Tabs');

  var tabrowsG = tabs.getRange("B2:B").getValues();
  console.log("values: ",tabrowsG);
  var rowsG = tabrowsG.filter(String).length;
  console.log("rows: ",rowsG);
  var grantTabs = tabs.getSheetValues(2,2,rowsG,1); //(row, column, no of rows, no of columns)
  var grantFunds = tabs.getSheetValues(2,3,rowsG,1); 
  var grantOrgs = tabs.getSheetValues(2,4,rowsG,1);
  console.log("grantTabs: ",grantTabs)
  
  var templateSheetG = ss.getSheetByName('Grant');
 
  for (var i=0; i<grantTabs.length; i++){ 
      var newSheetG = templateSheetG.copyTo(ss);
      var protections = templateSheetG.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      
      // Duplicate Template
      newSheetG.setName(grantTabs[i]);  // sets tab name
      console.log("i: ",i);
      console.log("sheet name: ",grantTabs[i]);
      newSheetG.getRange("A3").setValue(grantFunds[i]);  //puts Fund in A3
      newSheetG.getRange("A4").setValue(grantOrgs[i]);   //puts Org in A4

      for (var j = 0; j < protections.length; j++) {
        var protection = protections[j];
        var protectedRange = protection.getRange().getA1Notation();
        var newProtection = newSheetG.getRange(protectedRange).protect();
      }

      //ss.setActiveSheet(newSheet);  
      var protection = newSheetG.protect().setDescription(grantTabs[i]);
      }

 //Duplicate tabs from template 'Match':
  var tabs = ss.getSheetByName('Tabs');
  var tabrowsM = tabs.getRange("E2:E").getValues();
  console.log("values: ",tabrowsM);
  var rowsM = tabrowsM.filter(String).length;
  console.log("rows: ",rowsM);
  if (rowsM!=0){
    var matchTabs = tabs.getSheetValues(2,5,rowsM,1); //(row, column, no of rows, no of columns)
    var matchFunds = tabs.getSheetValues(2,6,rowsM,1); 
    var matchOrgs = tabs.getSheetValues(2,7,rowsM,1);
    console.log("sheet name: ",matchTabs);
    var templateSheetM = ss.getSheetByName('Match');

    for (var i=0; i<matchTabs.length; i++){
        var newSheetM = templateSheetM.copyTo(ss);
        var protections = templateSheetM.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        
        // Duplicate Template
        newSheetM.setName(matchTabs[i]);  // sets tab name
        console.log("sheet name: ",matchTabs[i]);
        newSheetM.getRange("A3").setValue(matchFunds[i]);  //puts Fund in A3
        newSheetM.getRange("A4").setValue(matchOrgs[i]);   //puts Org in A4

        for (var j = 0; j < protections.length; j++) {
          var protection = protections[j];
          var protectedRange = protection.getRange().getA1Notation();
          var newProtection = newSheetM.getRange(protectedRange).protect();
        }

        //ss.setActiveSheet(newSheet);  
        var protection = newSheetM.protect().setDescription(matchTabs[i]);
        }
  }
  
   // delete helper sheets and clean up:            
   ss.deleteSheet(ss.getSheetByName("Grant"));
   ss.deleteSheet(ss.getSheetByName("Match"));
   var sheets = ss.getSheets();
    ss.setActiveSheet(ss.getSheetByName('ACEP Emp'));
    ss.moveActiveSheet(sheets.length); //move to end
    SpreadsheetApp.getActive().getSheetByName('ACEP Emp').hideSheet(); //and hide
    ss.setActiveSheet(ss.getSheetByName('Hours per PR PMEC'));
    ss.moveActiveSheet(sheets.length);
    SpreadsheetApp.getActive().getSheetByName('Hours per PR PMEC').hideSheet();

}

function fillInFromProposal(){ //insert all the numbers from the proposal budget into each tab

}

function cleanUp() { //delete REF# and other cells that aren't needed
  
}
