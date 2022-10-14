////////////////////////////////////////////////////////////////////////////
//// BELOW FUNCTIONS UPDATE TASK TABS FROM SUBTASK (*ST) TABS ////////////
////////////////////////////////////////////////////////////////////////////
function fillNumbersfromSubtask() {
  const alasql = AlaSQLGS.load(); //load sql library
  
  //Find all subtask sheets and their fund numbers/actv codes
  const searchText = "*ST";
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const subsheets = ss.filter(s => s.getSheetName().includes(searchText));
  console.log(subsheets.length);
  if (subsheets.length > 0) 
    var funds = [];
    var actvcodes = [];
    for (i=0; i<subsheets.length; i++){
      console.log(subsheets[i].getSheetName());
      funds[i] = subsheets[i].getRange(3,1).getValue();
      actvcodes[i] = subsheets[i].getRange(5,1).getValue();
      console.log("funds: ", funds);
      console.log("actv: ", actvcodes);
    } 

  //pull values from TOAD labor GS
  var toad_labor_sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1SxvjuEMt-u9QucmX9axsRB6kBLUOopZGPWKbHt9Vmjo/edit#gid=2091278272'); //TOAD Labor GS
  var toad_labor_tab = toad_labor_sheet.getSheetByName('Sheet1');
  var values_labor = toad_labor_tab.getRange("A3:G").getValues();
  console.log("report labor: ",values_labor)
/*
  //pull values from TOAD exp GS
  var toad_exp_sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1MSxh75rzvJqdzVZ72-7mcDbmTD2TGMGhGi865gZ-UW8/edit#gid=524646617'); //TOAD Exp GS
  var toad_exp_tab = toad_exp_sheet.getSheetByName('Sheet1');
  var values_exp = toad_exp_tab.getRange("A3:L").getValues();
  console.log("report exp: ",values_exp)
*/  
  
  // var ss = SpreadsheetApp.getActiveSpreadsheet();
  //   //var fullSheet = ss.getSheetByName('Full');
  //   var sheets = ss.getSheets();
  //   var subtasksheets = ss.getSheets().getName().indexOf('*ST');
  //   console.log("subtasksheets: ", subtasksheets);
  //const query = "SELECT MATRIX [0] FROM ? WHERE [0] LIKE ('" +name+"%')";
  //var res = alasql(AlaSQLGS.transformQueryColsNotation(query), [values]);
  //var res1 = alasql('SELECT [0] FROM ?',[values]);

  //find next empty row after manual update of personnel with budget and import missing personnel (look them up in column A and add if they don't exist) from TOAD query who have charged to the grant
  //read this into array: IFERROR(QUERY(IMPORTRANGE("GS ID","Sheet1!$A$3:$G"),"select Col5 where Col2="&$A3&" and Col3="&$A$4&"",0),"")

  const query = "SELECT [4] FROM ? WHERE [1] = "+funds[0],
  res = alasql((query), [values_labor]);
  console.log("query: ", query);
  console.log(res);
}
