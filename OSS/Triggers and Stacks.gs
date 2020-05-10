function createTrigger(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //Deletes all previous triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  // Create new triggers
  ScriptApp.newTrigger("formSubmitStack").forSpreadsheet(ss).onFormSubmit().create();
  // ScriptApp.newTrigger("Sort_Timestamp").forSpreadsheet(ss).onOpen().create();
  ScriptApp.newTrigger("changeStack").forSpreadsheet(ss).onChange().create();
  
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2");
  sheet.getRange("A3:AK").sort({column: 5, ascending: false});
  
  Browser.msgBox('Great Job!!! Now go to the Form Responses Sheet!!!');
  
}





//-------onChange Stack-------------------------
function changeStack()
{
  validateIR();
  sendEmail();
}


//--------onFormSubmit Stack
function formSubmitStack() //Calling functions from other script works
{
  Sort_Timestamp();
  newFormColor();
  notifyEPC();
}

