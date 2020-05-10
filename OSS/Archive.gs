//Outdated and not functional as of 8/15/19. Will update

///***Note: Only Artem Should Run Any of these functions as they will create Triggers. 


function archiveMetrics()
{
  /*This function starts off by deleting all triggers to prevent any accidental emails
  Then it creates a new sheet and copies the values from the metrics table to the new sheet.
  The function then goes to the original metrics sheet and replaces the year in all of the formulas to next year (ex. 2019 becomes 2020)
  Lastly, It renames metrics archive sheet with today's date and adds all triggers back in
  
  ***Note: This function should only be run at 23:59 EST as it will overwrite all of the metrics formulas
  */
  
  
  //------Delete Triggers----------------------------------------
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  
  
  //------sheet info------------------------------------------
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var metricSheet = sheet.getSheetByName("Metrics");
  
  
  
  //-------Date------------------------------------------------------
  
  var today = new Date(); // today's date
  var year = today.getFullYear();
  var nextYr = year+1;
  
  
  
  //-------Archive Metrics---------------------------------------------
  sheet.insertSheet("OSS Metrics Archive");
  var metArch = sheet.getSheetByName("OSS Metrics Archive");
  
  metricSheet.getRange("E1:I14").copyTo(metArch.getRange("A1:E14"), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  metricSheet.getRange("E1:I14").copyTo(metArch.getRange("A1:E14"), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  
  
  
  //---Change Metrics Year to next year-----------------------------
  
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Metrics").getRange("E1").setValue(nextYr);
  
  //--------Set total OSS to next year-----------------------------------
  
  /*January*/ metricSheet.getRange("F2").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 1, year('Form Responses 2'!E3:E), " +nextYr+"))");
  /*Feb*/ metricSheet.getRange("F3").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 2, year('Form Responses 2'!E3:E), " +nextYr+"))");
  /*Mar*/ metricSheet.getRange("F4").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 3, year('Form Responses 2'!E3:E), " +nextYr+"))");
  /*Apr*/ metricSheet.getRange("F5").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 4, year('Form Responses 2'!E3:E), " +nextYr+"))");
  /*May*/ metricSheet.getRange("F6").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 5, year('Form Responses 2'!E3:E), " +nextYr+"))");
  /*June*/ metricSheet.getRange("F7").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 6, year('Form Responses 2'!E3:E), " +nextYr+"))");
  /*July*/ metricSheet.getRange("F8").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 7, year('Form Responses 2'!E3:E), " +nextYr+"))");
  /*Aug*/ metricSheet.getRange("F9").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 8, year('Form Responses 2'!E3:E), " +nextYr+"))");
  /*Sept*/ metricSheet.getRange("F10").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 9, year('Form Responses 2'!E3:E), " +nextYr+"))");
  /*Oct*/ metricSheet.getRange("F11").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 10, year('Form Responses 2'!E3:E), " +nextYr+"))");
  /*Nov*/ metricSheet.getRange("F12").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 11, year('Form Responses 2'!E3:E), " +nextYr+"))");
  /*Dec*/ metricSheet.getRange("F13").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 12, year('Form Responses 2'!E3:E), " +nextYr+"))");
  
  
  
  
  //------- Set OSS Escalated to Next year----------------------------------
  
  /*January*/ metricSheet.getRange("G2").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 1, year('Form Responses 2'!E3:E), " +nextYr+ ", 'Form Responses 2'!AJ3:AJ, \"<>\"))");
  /*Feb*/ metricSheet.getRange("G3").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 2, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!AJ3:AJ, \"<>\"))");
  /*Mar*/ metricSheet.getRange("G4").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 3, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!AJ3:AJ, \"<>\"))");
  /*Apr*/ metricSheet.getRange("G5").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 4, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!AJ3:AJ, \"<>\"))");
  /*May*/ metricSheet.getRange("G6").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 5, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!AJ3:AJ, \"<>\"))");
  /*June*/ metricSheet.getRange("G7").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 6, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!AJ3:AJ, \"<>\"))");
  /*July*/ metricSheet.getRange("G8").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 7, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!AJ3:AJ, \"<>\"))");
  /*Aug*/ metricSheet.getRange("G9").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 8, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!AJ3:AJ, \"<>\"))");
  /*Sept*/ metricSheet.getRange("G10").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 9, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!AJ3:AJ, \"<>\"))");
  /*Oct*/ metricSheet.getRange("G11").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 10, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!AJ3:AJ, \"<>\"))");
  /*Nov*/ metricSheet.getRange("G12").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 11, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!AJ3:AJ, \"<>\"))");
  /*Dec*/ metricSheet.getRange("G13").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 12, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!AJ3:AJ, \"<>\"))");
  
  
  
  
  //------- Set Completed OSS to Next year-------------------------------
  
  /*January*/ metricSheet.getRange("H2").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 1, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!D3:D,\"Completed\"))");
  /*Feb*/ metricSheet.getRange("H3").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 2, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!D3:D,\"Completed\"))");
  /*Mar*/ metricSheet.getRange("H4").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 3, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!D3:D,\"Completed\"))");
  /*Apr*/ metricSheet.getRange("H5").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 4, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!D3:D,\"Completed\"))");
  /*May*/ metricSheet.getRange("H6").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 5, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!D3:D,\"Completed\"))");
  /*June*/ metricSheet.getRange("H7").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 6, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!D3:D,\"Completed\"))");
  /*July*/ metricSheet.getRange("H8").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 7, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!D3:D,\"Completed\"))");
  /*Aug*/ metricSheet.getRange("H9").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 8, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!D3:D,\"Completed\"))");
  /*Sept*/ metricSheet.getRange("H10").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 9, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!D3:D,\"Completed\"))");
  /*Oct*/ metricSheet.getRange("H11").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 10, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!D3:D,\"Completed\"))");
  /*Nov*/ metricSheet.getRange("H12").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 11, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!D3:D,\"Completed\"))");
  /*Dec*/ metricSheet.getRange("H13").setValue("=ARRAYFORMULA(COUNTIFS(month('Form Responses 2'!E3:E), 12, year('Form Responses 2'!E3:E), "+nextYr+", 'Form Responses 2'!D3:D,\"Completed\"))");
  
  
  
  
  //------- Set OSS IR count to Next year-------------------------------
  
  /*January*/ metricSheet.getRange("I2").setValue("=ARRAYFORMULA(sumifs('Form Responses 2'!AK3:AK,month('Form Responses 2'!E3:E), 1, year('Form Responses 2'!E3:E), "+nextYr+"))");
  /*Feb*/ metricSheet.getRange("I3").setValue("=ARRAYFORMULA(sumifs('Form Responses 2'!AK3:AK,month('Form Responses 2'!E3:E), 2, year('Form Responses 2'!E3:E), "+nextYr+"))");
  /*Mar*/ metricSheet.getRange("I4").setValue("=ARRAYFORMULA(sumifs('Form Responses 2'!AK3:AK,month('Form Responses 2'!E3:E), 3, year('Form Responses 2'!E3:E), "+nextYr+"))");
  /*Apr*/ metricSheet.getRange("I5").setValue("=ARRAYFORMULA(sumifs('Form Responses 2'!AK3:AK,month('Form Responses 2'!E3:E), 4, year('Form Responses 2'!E3:E), "+nextYr+"))");
  /*May*/ metricSheet.getRange("I6").setValue("=ARRAYFORMULA(sumifs('Form Responses 2'!AK3:AK,month('Form Responses 2'!E3:E), 5, year('Form Responses 2'!E3:E), "+nextYr+"))");
  /*June*/ metricSheet.getRange("I7").setValue("=ARRAYFORMULA(sumifs('Form Responses 2'!AK3:AK,month('Form Responses 2'!E3:E), 6, year('Form Responses 2'!E3:E), "+nextYr+"))");
  /*July*/ metricSheet.getRange("I8").setValue("=ARRAYFORMULA(sumifs('Form Responses 2'!AK3:AK,month('Form Responses 2'!E3:E), 7, year('Form Responses 2'!E3:E), "+nextYr+"))");
  /*Aug*/ metricSheet.getRange("I9").setValue("=ARRAYFORMULA(sumifs('Form Responses 2'!AK3:AK,month('Form Responses 2'!E3:E), 8, year('Form Responses 2'!E3:E), "+nextYr+"))");
  /*Sept*/ metricSheet.getRange("I10").setValue("=ARRAYFORMULA(sumifs('Form Responses 2'!AK3:AK,month('Form Responses 2'!E3:E), 9, year('Form Responses 2'!E3:E), "+nextYr+"))");
  /*Oct*/ metricSheet.getRange("I11").setValue("=ARRAYFORMULA(sumifs('Form Responses 2'!AK3:AK,month('Form Responses 2'!E3:E), 10, year('Form Responses 2'!E3:E), "+nextYr+"))");
  /*Nov*/ metricSheet.getRange("I12").setValue("=ARRAYFORMULA(sumifs('Form Responses 2'!AK3:AK,month('Form Responses 2'!E3:E), 11, year('Form Responses 2'!E3:E), "+nextYr+"))");
  /*Dec*/ metricSheet.getRange("I13").setValue("=ARRAYFORMULA(sumifs('Form Responses 2'!AK3:AK,month('Form Responses 2'!E3:E), 12, year('Form Responses 2'!E3:E), "+nextYr+"))");
  
  
  
  //-----rename metArch Sheet---------------------------
  sheet.setActiveSheet(metArch);
  sheet.renameActiveSheet('OSS Metrics Archive ' + (today.getMonth()+1) + "-" + today.getDate()+"-"+ today.getFullYear());
  
  
  
  //-----recreate Triggers---------------------------
  
  ScriptApp.newTrigger("Sort_Timestamp").forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger("Sort_Timestamp").forSpreadsheet(ss).onOpen().create();
  ScriptApp.newTrigger("sendEmail").forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger("notifyEPC").forSpreadsheet(ss).onFormSubmit().create();
  
  
}// Close archiveMetrics








//-----------------------------------------------------------------------------------------------------------------------------------------










function archiveResponses() 

/*This function starts off by deleting all triggers to prevent any accidental emails
Then it creates a new sheet and moves all completed rows from the current year to a new sheet
Lastly, It renames metrics archive sheet with today's date and adds all triggers back in


***Note: DO NOT RUN this function without first running archiveMetrics. Metrics references this data and as such values will lose values when responses are archived.
*/

{
  
  
  //------Delete Triggers----------------------------------------
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  
  
  
  //------------------Sheet Info---------------------------------------- 
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var responseSheet = sheet.getSheetByName("Form Responses 2");
  
  //-----Date--------------------------------------
  var today = new Date(); // today's date
  var year = today.getFullYear();
  
  //------------------Row Info-----------------
  var row = 3;
  var lastRow = sheet.getSheetByName("Form Responses 2").getLastRow();
  
  
  
  
  
  
  //-----Archive Responses----------------------------------
  sheet.insertSheet("Archive"); // creates new sheet called Archive
  var archSheet = sheet.getSheetByName("Archive");
  
  responseSheet.getRange("A1:AJ2").copyTo(archSheet.getRange("A1:AJ2")); //Copy Headers to first row of new sheet
  
  
  //-----------------Loop Through All Rows--------------------------  
  
  while(row<=lastRow)
  {
    var status = responseSheet.getRange("D"+row).getValue(); // status in last row
    
    if (status == "Completed")
    {            
      archSheet.insertRowsBefore(3,1);//create new row above row 3 of archive sheet
      var rangeToCopy = responseSheet.getRange("A" + row + ":AJ" + row);
      var rangeToPaste = archSheet.getRange("A3:AJ3");
      rangeToCopy.copyTo(rangeToPaste);//past values into top row of Archive sheet
      responseSheet.deleteRow(row);
    }//closes if (status = completed)
    //------------------------------------------------------------    
    else if (status != "Completed")
    {
      row++;
      status = responseSheet.getRange("D"+row).getValue(); // status in last row
    }
    
    
  }//closes archive Responses
  
  
  sheet.setActiveSheet(archSheet);
  sheet.renameActiveSheet('OSS Response Archive ' + (today.getMonth()+1) + "-" + today.getDate()+"-"+ today.getFullYear());
  
  //-----recreate Triggers---------------------------
  
  ScriptApp.newTrigger("Sort_Timestamp").forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger("Sort_Timestamp").forSpreadsheet(ss).onOpen().create();
  ScriptApp.newTrigger("sendEmail").forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger("notifyEPC").forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger("validateIR").forSpreadsheet(ss).onChange().create();
  
  
}//Closes whole function





//-----------------------------------------------------------------------------------------------------------------------






function archiveAll()
{
  archiveMetrics();
  archiveResponses();
  
}


