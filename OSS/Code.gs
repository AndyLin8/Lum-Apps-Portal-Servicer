
//--------------------------------------------------------------

function Sort_Timestamp()
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2");
  sheet.getRange("A3:AK").sort({column: 5, ascending: false});
  
  
  /* //sets New SID to caps
  var sid = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("I3").getValue();
  var upperSid = sid.toUpperCase();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("I3").setValue(upperSid);
  */ 
}







//------------Functions for writing custom email message-----------------------------------------------

function writeEmail()
{
  var row = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getCurrentCell().getRow(); 
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Please type your message to the user.Click OK to send Email.', ui.ButtonSet.OK_CANCEL);
  var text = response.getResponseText();
  
  
  var confirmation = response.getSelectedButton();
  
  if (confirmation == "CANCEL")
  {
    return "NO";
  }
  
  else
  {
    PropertiesService.getUserProperties().setProperty('emailBody', text);
  }
}




function emailBody()
{
  var body = PropertiesService.getUserProperties().getProperty('emailBody'); //makeshift global variable
  //Logger.log(body);
  return body;
}







// --------Sends email notification to EPC email upon each form submission-------------------------

function notifyEPC()
{ 
  var row = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getCurrentCell().getRow();
  var email =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("C" + row ).getValue(); // gets email value from current row
  var bothEmails= "daniel_tambor@colpal.com";//"esc_prod_control@colpal.com"+ "," +  "esc_operations@colpal.com";
  var html = HtmlService.createHtmlOutputFromFile("Notify_EPC").getContent();
  
  
  MailApp.sendEmail ( {
    to: bothEmails,
    subject: "New OSS Request Has Been Submitted",
    htmlBody: html,
    noReply: true 
  });
}







//---Send emails on status change----------------------------------------------------------------

function sendEmail(){
  var bothEmails= "esc_prod_control@colpal.com"+ "," +  "esc_operations@colpal.com";
  var dbaEmail= "daniel_tambor@colpal.com";//"ESC_DBA_Group@colpal.com"; 
  var row = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getCurrentCell().getRow(); 
  var col = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getCurrentCell().getColumn();  
  
  var inProgConfirm =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("AH" + row ).getValue(); 
  var completeConfirm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("AI" + row ).getValue();
  var escalateConfirm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("AJ" + row ).getValue();
  
  var ir =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("B" + row ).getValue(); 
  var status = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("D" + row ).getValue(); 
  var email =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("C" + row ).getValue();
  
  if  ( col == 4 && ir == "" && SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()== "Form Responses 2")
  
  {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("D" + row ).setValue("");
    Browser.msgBox('There is no IR#- email will not send until IR is created and status is changed');
  } 
  
  
  
  
  
  //---------------In Progress ---------------------------------------------------------------------------
  else if ( ir != "" && status == "In Progress" && col == 4 && inProgConfirm == "" && SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()== "Form Responses 2")
  {
    var subject = ("IR#:"+ ir+ " " + "| OSS Status Update: Request Acknowledged");
    var html = HtmlService.createHtmlOutputFromFile("Request_Acknowledged").getContent();
    MailApp.sendEmail ( {
      to: email,
      subject: subject,
      htmlBody: html,
      noReply: true 
    });
    
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("AH" + row ).setValue("In Progress email sent"); 
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("AG" + row ).setValue(new Date()).setNumberFormat('M/D/YYYY hh:mm:ss')
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("A" + row + ":AK" + row ).setBackgroundColor("orange");
  }    
  
  
  
  
  
  
  
  //--------------- Completed ----------------------------------------------------------------------------------------------
  else if ( ir != "" && status == "Completed" && col == 4 && completeConfirm == "" && SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()== "Form Responses 2")
  { 
    
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Are you sure you want to set to \"Complete\"? An email will be sent and this record will be archived', ui.ButtonSet.YES_NO);
    
    if (response == ui.Button.YES)
    {
      
      
      var subject = ("IR#:" + ir+ " " + " | OSS Status Update: Request Completed");
      var html = HtmlService.createHtmlOutputFromFile("Request_Completed").getContent();
      
      try {  
        MailApp.sendEmail ( {
          to: email,
          subject: subject,
          htmlBody: html,
          noReply: true 
        });
        
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("AI"+row).setValue("Completed email sent");
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("A" + row + ":AK" + row ).setBackgroundColor("#00ff0d");
      }
      catch(e) {}
      
      
      
      var completeIRs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Completed").getRange("B3").getValue();
      var responseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2");
      var archSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Completed");
      archSheet.insertRowsBefore(3,1);//create new row at top of archive sheet
      var rangeToCopy = responseSheet.getRange("A" + row + ":AK" + row);
      var rangeToPaste = archSheet.getRange("A3:AK3");
      rangeToCopy.copyTo(rangeToPaste);//past values into top row of Archive sheet
      responseSheet.deleteRow(row);
      
      
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Completed").getRange("AG3").setValue(new Date()).setNumberFormat('M/D/YYYY hh:mm:ss')
      
      try {SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Completed").getRange("AK3").setValue(completeIRs.split("/").length);}
      catch(err){SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Completed").getRange("AK3").setValue(1);} 
      
    }
    
    else if (response == ui.Button.NO)
    {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("D" + row ).setValue("");
    }
  }
  
  
  
  
  
  
  //-----------Escalated to L3 --------------------------------------------------------------------------------------------------
  
  else if ( ir != "" && status == "Escalated to L3" && col == 4 && escalateConfirm == "" && SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()== "Form Responses 2")
  { 
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Are you sure you want to Escalate to L3? An email will be sent upon clicking \"Yes\"', ui.ButtonSet.YES_NO);
    
    if (response == ui.Button.YES)
    {
      var subject = ("IR#:" + ir+ " " + " | OSS Status Update: Escalated to L3");
      var html = HtmlService.createHtmlOutputFromFile("Request_Escalated").getContent();
      MailApp.sendEmail ( {
        to: dbaEmail,
        subject: subject,
        htmlBody: html,
        noReply: true 
      });
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("AJ" + row ).setValue("Escalated to L3 email sent"); 
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("AG" + row ).setValue(new Date()).setNumberFormat('M/D/YYYY hh:mm:ss');
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("A" + row + ":AK" + row ).setBackgroundColor("#ff4ac9");
    }
    
    else {SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("D" + row ).setValue("");}
  }   
  
  
  //-----------Need More Info --------------------------------------------------------------------------------------------------
  
  else if ( ir != "" && status == "Need More Info From User" && col == 4 && SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()== "Form Responses 2")
  {
    var confirm = writeEmail();
    
    if (confirm != "NO")
    {
      
      var subject = ("IR#:"+ ir+ " " + "| OSS Request Status Update: Need More Information");
      var html = HtmlService.createTemplateFromFile("need_info").evaluate().getContent();
      
      MailApp.sendEmail ( {
        to: email,
        subject: subject,
        htmlBody: html,
        noReply: false,
        replyTo: "esc_prod_control@colpal.com" + "esc_operations@colpal.com",
        name: "esc_prod_control@colpal.com" + "esc_operations@colpal.com"
      });
      
      
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("A" + row + ":AK" + row ).setBackgroundColor("#3d7bff");
      PropertiesService.getUserProperties().deleteProperty('emailBody');
    }
    
    else 
    {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("D" + row ).setValue("");
      PropertiesService.getUserProperties().deleteProperty('emailBody');
    }
  }
  
  
}// this bracket closes the whole function 




//-----------------------------------------------------------------------------

function validateIR()
{
  
  var col = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getCurrentCell().getColumn();
  
  
  var cell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getCurrentCell().getValue();
  var patt = new RegExp(/[a-z]|\,|\.|\\|\(|\)|\:|\;|\'|\s|^\"/ig);      
  
  
  if (col == 2 && SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()== "Form Responses 2" && patt.test(cell)== true)
  {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getCurrentCell().setValue("");
    Browser.msgBox("Invalid characters in IR. Only use numbers and \"/\" symbol.");    
  }
}

//-----------------------------------------------------------------------------------------------
function newFormColor()
{
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("A3:AK3").setBackgroundColor("Red");
  
}


