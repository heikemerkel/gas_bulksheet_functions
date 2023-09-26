//https://spreadsheet.dev/triggers-in-google-sheets
//https://stackoverflow.com/questions/33188514/programmatically-delete-a-google-apps-script-time-based-trigger-by-trigger-name


// // triggers to run on edit
// ScriptApp.newTrigger("fillNamesFromTOAD") // Run the sendEmailReport function.
//     .forSpreadsheet(SpreadsheetApp.getActive()) // Create the trigger in this spreadsheet.
//     .onEdit() // We want to set up an Edit trigger.
//     .create();

// // triggers to run between 5am-6am in the timezone of the script
// ScriptApp.newTrigger("myFunction")
//   .timeBased()
//   .atHour(5)
//   .everyDays(1) // Frequency is required if you are using atHour() or nearMinute()
//   .create();

// create triggers for simple budget sheet
function createTriggerS() {
  ScriptApp.newTrigger("fillNamesFromTOAD").timeBased().atHour(22).everyDays(1).create();
  ScriptApp.newTrigger("clearEntriesS").timeBased().atHour(0).everyDays(1).create();
  ScriptApp.newTrigger("fillNamesInFullS").timeBased().atHour(2).everyDays(1).create();
  ScriptApp.newTrigger("fillNumbersInFullLaborS").timeBased().atHour(4).everyDays(1).create();
  ScriptApp.newTrigger("fillNumbersInFullExpS").timeBased().atHour(6).everyDays(1).create();
  // ScriptApp.newTrigger("emailAlerts").forSpreadsheet(SpreadsheetApp.getActive()).timeBased().atHour(8).everyDays(1).create();
  // ScriptApp.newTrigger("").forSpreadsheet(SpreadsheetApp.getActive()).timeBased().atHour(10).everyDays(1).create();
}

// create triggers for extended budget sheets - this needs to be adjusted based on the specific budget sheet
function createTrigger() {
  ScriptApp.newTrigger("fillNamesFromTOAD").timeBased().atHour(22).everyDays(1).create();
  ScriptApp.newTrigger("clearEntries").timeBased().atHour(0).everyDays(1).create();
  ScriptApp.newTrigger("fillNamesInFull").timeBased().atHour(2).everyDays(1).create();
  ScriptApp.newTrigger("fillNumbersInFullLabor2").timeBased().atHour(4).everyDays(1).create();
  ScriptApp.newTrigger("fillNumbersInFullLabor14").timeBased().atHour(6).everyDays(1).create();
  ScriptApp.newTrigger("fillNumbersInFullExp").timeBased().atHour(8).everyDays(1).create();
  // ScriptApp.newTrigger("emailAlerts").forSpreadsheet(SpreadsheetApp.getActive()).timeBased().atHour(8).everyDays(1).create();
  // ScriptApp.newTrigger("").forSpreadsheet(SpreadsheetApp.getActive()).timeBased().atHour(10).everyDays(1).create();
}

function deleteTriggers(){
var triggers = ScriptApp.getProjectTriggers();
//console.log("triggers: ", triggers);
 for (var i = 0; i < triggers.length; i++) {
   ScriptApp.deleteTrigger(triggers[i]);
   //console.log(triggers[i]);
 }
}

function listTriggersByName(){
var triggers = ScriptApp.getProjectTriggers();
 for (var i = 0; i < triggers.length; i++) {
    Logger.log(triggers[i].getHandlerFunction())
  }
}














