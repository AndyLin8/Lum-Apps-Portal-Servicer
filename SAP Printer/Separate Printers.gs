function separatePrinters() 
{
  
  var idRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses").getRange("A3:F3");
  var idPasteRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("A3:F3");
  
  var p1Range = "G3:N3";
  var p2Range= "P3:W3";
  var p3Range = "Y3:AF3";
  var p4Range = "AH3:AO3";
  var p5Range = "AQ3:AX3";
  var p6Range = "AZ3:BG3";
  var p7Range = "BI3:BP3";
  var p8Range = "BR3:BY3";
  var p9Range = "CA3:CH3";
  var p10Range = "CJ3:CQ3"; 
  
  var printers = [p1Range, p2Range, p3Range, p4Range, p5Range, p6Range, p7Range, p8Range, p9Range, p10Range];
  

  
  for (i = 0; i < printers.length; i++)
  {
    var test =SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses").getRange(printers[i]).getValues()[0][1]; //gets value of printer Name field

    if (test != "")  
    {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").insertRowsBefore(3,1);
      var rangeToCopy = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses").getRange(printers[i]);
      var rangeToPaste = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("G3:N3");
      idRange.copyTo(idPasteRange);
      rangeToCopy.copyTo(rangeToPaste);
      
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Printers").getRange("A3:R3").setBackground("Red");
    }
  }
  
  //SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses").deleteRow(3); //deletes request from Form sheet after copying values
}


