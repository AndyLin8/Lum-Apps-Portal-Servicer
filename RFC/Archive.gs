//outdated and not functional as of 8/15. will update

function archive() 
{  
 //------------------Sheet Info---------------------------------------- 
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var responseSheet = sheet.getSheetByName("Form Responses 2");
  
  sheet.insertSheet("Archive"); // creates new sheet called Archive
  var archSheet = sheet.getSheetByName("Archive");
  
  var today = new Date(); // today's date
  
 //-----------Copy Headers to first row of new sheet--------------------------
 
  responseSheet.getRange("A1:AB2").copyTo(archSheet.getRange("A1:AB2"));
  
  
  
 //------------------Row Info-----------------
  var row = 3;
  
  var lastRow = sheet.getSheetByName("Form Responses 2").getLastRow();

  
  //-----------------Loop Through All Rows--------------------------  
  
  while(row<=lastRow)
  {
    var status = responseSheet.getRange("D"+row).getValue(); // status in last row
    
    if (status == "Completed")
    {            
      archSheet.insertRowsBefore(3,1);//create new row at top of archive sheet
      var rangeToCopy = responseSheet.getRange("A" + row + ":AB" + row);
      var rangeToPaste = archSheet.getRange("A3:AB3");
      rangeToCopy.copyTo(rangeToPaste);//past values into top row of Archive sheet
      responseSheet.deleteRow(row);
    }//closes if (status = completed)
    
    
   
    else if (status != "Completed")
    {
      row++;
      status = responseSheet.getRange("D"+row).getValue(); // status in last row
    }
    

  }//closes while loop
  

 sheet.setActiveSheet(archSheet);
 sheet.renameActiveSheet('Archive ' + (today.getMonth()+1) + "-" + today.getDate()+"-"+ today.getFullYear());

}//Closes whole function


