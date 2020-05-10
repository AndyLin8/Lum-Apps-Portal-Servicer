
//---------------------------------------------------------------------------------------------------

function Sort_Timestamp()
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2");
  sheet.getRange("A:AC").sort({column: 6, ascending: false});
  
}




//------------Functions for writing custom email message-----------------------------------------------



/*function writeEmail()
{
  var row = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getCurrentCell().getRow(); 
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Please type your message to the user.Click OK to send Email.', ui.ButtonSet.YES_NO);
  var text = response.getResponseText();
  
  
  var confirmation = response.getSelectedButton();
  
  if (confirmation == "NO")
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


*/





//-------------functions for notifying purna-----------------------------------------------------------
function verifyPurna() //checks to see if any of purna's sids are in the request
{

  
  var row = 3;
  var purnaSIDS = ["LA2","LM6","LAD","LAT","LAQ","LAP"];
  var sids1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("H"+row).getValue().toUpperCase();
  var sids2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("P"+row).getValue().toUpperCase();
  var allSids = sids1+", "+sids2;
  var sidHits = [];
  
  
  
  
  
  for (i = 0; i < purnaSIDS.length; i++)
  {
    if (allSids.indexOf(purnaSIDS[i])>=0)
    {
      sidHits.push(purnaSIDS[i]);
    }
  }
  
  return sidHits.join();
}





function purnaEmailString() //returns the string for output in the html template
{
  var row = 3;
  var hits = verifyPurna();
  var email =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("C" + row ).getValue();
  var emailString = email + " with SIDs " + hits;
  
  return emailString;
}




function notifyPurna() //sends email to purna
{
  
  var row = 3;
  
  
  if (verifyPurna() != "")
  {
    var html2 = HtmlService.createTemplateFromFile("notifyPurna").evaluate().getContent();
    
    
    var subject = ("Purna's New RFC Request Notification");
    MailApp.sendEmail ( {
      to: "daniel_tambor@colpal.com", ///change to purna's email
      subject: subject,
      htmlBody: html2,
      noReply: true 
    });
  }
}












// --------Sends email notification to EPC email upon each form submission-------------------------
function notifyEPC()
{ 
  //notifyPurna(); // notifies purna for any of his requested SIDs
  
  var bothEmails= "daniel_tambor@colpal.com"+ "," +  "tulsi_patel@colpal.com"; //change to prod control and operations emails
  var html = HtmlService.createHtmlOutputFromFile("Notify_EPC").getContent();
  
  
  MailApp.sendEmail ( {
    to: bothEmails,
    subject: "New RFC Request Has Been Submitted",
    htmlBody: html,
    noReply: true 
  });
  
}













//------Send emails on status change-------------------------------------------------------------

function sendEmail(){
  var dbaEmail= "daniel_tambor@colpal.com"//"ESC_DBA_Group@colpal.com"; //change to ESC_DBA_Group@colpal.com during deployment 
  var row = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getCurrentCell().getRow(); 
  var col = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getCurrentCell().getColumn();  // this is the active row and column 
  
  
  
  var inProgConfirm =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("Y" + row ).getValue(); 
  var completeConfirm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("Z" + row ).getValue();
  var escalateConfirm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("AB" + row ).getValue();
  
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
    var subject = ("IR#:"+ ir+ " " + "| RFC Request Status Update: Request Acknowledged");
    var html = HtmlService.createHtmlOutputFromFile("Request_Acknowledged").getContent();
    MailApp.sendEmail ( {
      to: email,
      subject: subject,
      htmlBody: html,
      noReply: true 
    });
    
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("Y" + row ).setValue("In Progress email sent"); 
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("X" + row ).setValue(new Date()).setNumberFormat('M/D/YYYY hh:mm:ss') //Most recent change
    
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("A" + row+ ":AC" +row ).setBackgroundColor("orange");
  }    
  
  
  
  
  
  
  
  //--------------- Completed ----------------------------------------------------------------------------------------------
  else if ( ir != "" && status == "Completed" && col == 4 && completeConfirm == "" && SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()== "Form Responses 2")
  { 
    
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Are you sure you want to set to \"Complete\"? An email will be sent and this record will be archived', ui.ButtonSet.YES_NO);
    
    if (response == ui.Button.YES)
    {
      
      var subject = ("IR#:"+ ir+ " " + "| RFC Request Status Update: Request Completed");
      var html = HtmlService.createHtmlOutputFromFile("Request_Completed").getContent();
      
      try {  
        MailApp.sendEmail ( {
          to: email,
          subject: subject,
          htmlBody: html,
          noReply: true 
        });
        
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("Z"+row).setValue("Completed email sent");
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("A" + row+ ":AC" +row ).setBackgroundColor("#00ff0d");
      }
      catch(e) {}
      
      
      var completeIRs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Completed").getRange("B3").getValue();
      var responseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2");
      var archSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Completed");
      archSheet.insertRowsBefore(3,1);//create new row at top of archive sheet
      var rangeToCopy = responseSheet.getRange("A" + row + ":AC" + row);
      var rangeToPaste = archSheet.getRange("A3:AC3");
      rangeToCopy.copyTo(rangeToPaste);//past values into top row of Archive sheet
      responseSheet.deleteRow(row);
      
      
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Completed").getRange("X3").setValue(new Date()).setNumberFormat('M/D/YYYY hh:mm:ss')
      
      try {SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Completed").getRange("AC3").setValue(completeIRs.split("/").length);}
      
      catch(err){SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Completed").getRange("AC3").setValue(1);}    
      
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
    
    if (response == ui.Button.YES){
      
      
      var subject = ("IR#:"+ ir+ " " + "| RFC Request Status Update: Escalated to L3");
      var html = HtmlService.createHtmlOutputFromFile("Request_Escalated").getContent();
      MailApp.sendEmail ( {
        to: dbaEmail,
        subject: subject,
        htmlBody: html,
        noReply: true 
      });
      
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("AB" + row ).setValue("Escalated to L3 email sent"); 
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("X" + row ).setValue(new Date()).setNumberFormat('M/D/YYYY hh:mm:ss');
      
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("A" + row+ ":AC" +row ).setBackgroundColor("#ff4ac9");
    }
    
    else if (response == ui.Button.NO)
    {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("D" + row ).setValue("");
    }
    
  }
  
  
  //-----------Need More info --------------------------------------------------------------------------------------------------
  
  
  else if ( ir != "" && status == "Need More Info From User" && col == 4 && SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()== "Form Responses 2")
  {
    Logger.log("Step: before getui");
    var ui = SpreadsheetApp.getUi();
    Logger.log("Step: after getui");
    ui.alert('Please type your message to the user.Click OK to send Email.', ui.ButtonSet.YES_NO)
    
    Browser.inputBox('Enter your name', Browser.Buttons.OK_CANCEL);
    var confirm = writeEmail();
    Logger.log("Step: after writeEmail");
    if (confirm != "NO")
    {
      
      var subject = ("IR#:"+ ir+ " " + "| RFC Request Status Update: Need More Information");
      var html = HtmlService.createTemplateFromFile("need_info").evaluate().getContent();
      
      MailApp.sendEmail ( {
        to: email,
        subject: subject,
        htmlBody: html,
        noReply: false,
        replyTo: "daniel_tambor@colpal.com",
        name: "daniel_tambor@colpal.com"
      });
      
      
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("A" + row + ":AC" + row ).setBackgroundColor("#3d7bff");
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("helper").getRange("A1").setValue("");
      //PropertiesService.getUserProperties().deleteProperty('emailBody');
    }
    
    else 
    {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("D" + row ).setValue("");
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("helper").getRange("A1").setValue("");
      //PropertiesService.getUserProperties().deleteProperty('emailBody');
    }
  }
  
  
  
  
  
  
}// this bracket closes the whole email function 








//----------------------------------------------------------------------------------------------
function validateIR()
{
  
  var col = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getCurrentCell().getColumn();
  
  
  var cell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getCurrentCell().getValue();
  var patt = new RegExp(/[a-z]|\,|\.|\\|\(|\)|\:|\;|\'|\s|^\"/ig);      
  
  if (col == 2 && SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()== "Form Responses 2" && patt.test(cell)== true)
  {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getCurrentCell().setValue("");
    Browser.msgBox("Invalid characters. Only use numbers and \"/\" symbol.");    
  }
}


//----------------------------------------------------------------------------------------------


function newFormColor()
{
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 2").getRange("A3:AC3").setBackgroundColor("Red");
  
}

