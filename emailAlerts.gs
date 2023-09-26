function emailAlerts() {
  // check for total charges in tab 'Grant'
  var sheetName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Grant");
  var textFinder = sheetName.getRange("A9:A").createTextFinder("Grand TOTAL").findAll();
  //console.log("textfinder: ",textFinder);
  var rcs = textFinder[0].getRowIndex();
  //console.log("rcs: ",rcs);
  var totalCharges = sheetName.getRange(rcs,18).getValue();
  //console.log('totalCharges:',totalCharges);
    
  if (totalCharges < 0){
    //send email alert
    var message = 'The Total is now below the threshold.';//... tab ...
    var subject = '[Budget Sheets] LTER';
    MailApp.sendEmail('hmerkel@alaska.edu',subject, message);

    // For multiple email addresses, use below
    // var emailAddresses = "hmerkel@alaska.edu, ljwalls@alaska.edu";
    // // console.log("emails:", emailAddresses);
    // MailApp.sendEmail(emailAddresses, subject, message);
  }
}
