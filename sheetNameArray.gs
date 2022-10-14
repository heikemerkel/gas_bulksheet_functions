function sheetURLArray() {
  var out = [];
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var j = 0;  
  for (var i=2; i<sheets.length-2; i++)  {
    var SS = SpreadsheetApp.getActiveSpreadsheet();
    var sheetsUrl = SS.getUrl();
    console.log("sheeturl: ", sheetsUrl);
    out[j] = SS.getUrl();
    out[j] += '#gid=';
    out[j] += sheets[i].getSheetId();
    j += 1;
  }  
    out.push();
    return out 
}

function sheetNameArray(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheet_names = []
  for (var i=2; i< sheets.length-2; i++) {
    sheet_names.push([sheets[i].getName()])
  }
  return sheet_names
}