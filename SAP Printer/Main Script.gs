
//----------------- Function Sort Time Stamp ------------------------------------------------------------------------
function Sort_Timestamp()
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses");
  sheet.getRange("A3:CU").sort({column: 2, ascending: false});
  
}







//------------------Notify EPC function---------------------------------------
// --------Sends email notification to EPC email upon each form submission-------------------------

function notifyEPC()
{ 
  var bothEmails= "daniel_tambor@colpal.com"; //"esc_prod_control@colpal.com"+ "," +  "esc_operations@colpal.com";
  var html = HtmlService.createTemplateFromFile("Notify_EPC").evaluate().getContent();
  
  MailApp.sendEmail ( {
    to: bothEmails,
    subject: "New Printer Request Has Been Submitted",
    htmlBody: html,
    noReply: true 
  });
}









//-------------------Function Send Email on status change--------------------------------------

function sendEmail(){
  var row = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getCurrentCell().getRow(); // this is the active row
  var col = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getCurrentCell().getColumn(); // this is the active column 
  
  
  
  // Fetching value from confirmation fields
  
  var testDevConfirm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("P" + row ).getValue();
  var prodConfirm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("Q" + row ).getValue();
  var completeConfirm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("R" + row ).getValue();
  
  var status = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("D" + row ).getValue(); //gets status value from current row
  var email =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("C" + row ).getValue(); // gets email value from current row
  var cr= SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("G" + row ).getValue();
  Logger.log(status);
  Logger.log(email);
  Logger.log(cr);
  
  
  if  ( col == 4 && cr == "" && SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()== "Printers")
  
  {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("D" + row ).setValue("");
    Browser.msgBox('There is no CR#- email will not send until CR is created and status is changed');
  }
  
  
  
  
  
  
  //---------------Test and Dev In Progress Email ---------------------------------------------------------------------------
  
  else if (status == "Test and Dev In Progress" && col == 4 && testDevConfirm == "" && SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()== "Printers")
  { 
    var subject = ("SAP Printer Setup Needed Action; Test Your Printers. Printer now in Test and Dev.");
    var html = HtmlService.createTemplateFromFile("UpdateInProgress-TestandDev").evaluate().getContent();
    
    MailApp.sendEmail ( {
      to: email,
      subject: subject,
      htmlBody: html,
      //replyTo: "esc_prod_control@colpal.com" + "esc_operations@colpal.com",
      // name: "esc_operations@colpal.com" + " ESC_prod_control@colpal.com"
    });
    
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("P" + row ).setValue(" Test & Dev In Progress Email Sent");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("O" + row ).setValue(new Date()).setNumberFormat('M/D/YYYY hh:mm:ss') // Most Recent Change 
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("A" + row + ":R" + row).setBackgroundColor("Orange");
  }
  
  
  
  
  
  
  
  //-------------Production In Progress Email--------------------------------  
  else if (status == "Production In Progress" && col == 4 && prodConfirm == "" && SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()== "Printers")
  { 
    var subject = ("SAP Printer Setup Needed Action; Test Printers. Printer now in production.");
    var html = HtmlService.createTemplateFromFile("UpdateInProgress-Prod").evaluate().getContent();
    
    
    MailApp.sendEmail ( {
      to: email,
      subject: subject,
      htmlBody: html,
      //replyTo: "esc_prod_control@colpal.com" + "esc_operations@colpal.com", 
      //name: "esc_operations@colpal.com" + " ESC_prod_control@colpal.com" 
    });
    
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("Q" + row ).setValue("Production In Progress Email Sent");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("O" + row ).setValue(new Date()).setNumberFormat('M/D/YYYY hh:mm:ss')
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("A" + row + ":R" + row).setBackgroundColor("#ffd54d");
  }
  
  
  
  
  
  //--------------------Completed Email----------------------------
  else if (status == "Completed" && col == 4 && completeConfirm == "" && SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()== "Printers")
  { 
    
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Are you sure you want to set to \"Complete\"? An email will be sent and this record will be archived', ui.ButtonSet.YES_NO);
    
    if (response == ui.Button.YES)
    {
      var subject = ("SAP Printer Setup Completed");
      var html = HtmlService.createTemplateFromFile("Update-Completed").evaluate().getContent();
      
      
      try{
        MailApp.sendEmail ( {
          to: email,
          subject: subject,
          htmlBody: html,
          //replyTo: "esc_prod_control@colpal.com" + "esc_operations@colpal.com",
          //name: "esc_operations@colpal.com" + " ESC_prod_control@colpal.com" 
        });
        
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("R" + row ).setValue("Completion Email Sent");
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("O" + row ).setValue(new Date()).setNumberFormat('M/D/YYYY hh:mm:ss')
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("A" + row + ":S" + row).setBackgroundColor("#00ff0d");
      }
      
      catch (e) {}
      
      
      var responseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers");
      var archSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Completed");
      archSheet.insertRowsBefore(3,1);//create new row at top of archive sheet
      var rangeToCopy = responseSheet.getRange("A" + row + ":R" + row);
      var rangeToPaste = archSheet.getRange("A3:R3");
      rangeToCopy.copyTo(rangeToPaste);//past values into top row of Archive sheet
      responseSheet.deleteRow(row);
      try {SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Completed").getRange("S3").setValue(completeIRs.split("/").length);}
      catch(err){SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Completed").getRange("S3").setValue(1);}
      
    }
    
    if (response == ui.Button.NO)
    {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("D" + row).setValue("");
    }
  }
}










//------Validate CR-----------------------------------------------------------------------------------------------------------


function validateCR()
{
  
  var col = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getCurrentCell().getColumn();
  var cell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getCurrentCell().getValue();
  var patt = new RegExp(/[a-z]|\,|\.|\\|\(|\)|\:|\;|\'|\s|^\"/ig);      
  
  
  if (col == 7 && SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()== "Printers" && patt.test(cell)== true)
  {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getCurrentCell().setValue("");
    Browser.msgBox("Invalid characters in CR. Only use numbers and \"/\" symbol.");    
  }
}




//------Create output with printer Name and IP for email-----------------------------------------------------------------------------------------------------------
function outputCR()
{
  var row = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getCurrentCell().getRow(); // this is the active row
  var cr= SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("G" + row ).getValue();
  
  try{
    return "CR(s): "+ cr.split("/");
  }
  
  catch(e)
  {
    var cr= SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("G" + row ).getValue();
    return "CR(s): "+ cr;
  }
}



function outputPrinterName()
{
  var row = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getCurrentCell().getRow(); // this is the active row
  var printerName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("H" + row).getValue();
  
  return "Printer Name: " +printerName;
}





function outputIP()
{
  var row = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getCurrentCell().getRow(); // this is the active row
  var ip = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("I" + row ).getValue();
  
  return "IP Address: " + ip;
  
}

