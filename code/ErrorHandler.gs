///****************************** ERROR HANDLER ***************************************************
// This takes care of Error Handling . Specifically the below tasks 
// 1. Show the Error in a Modal Dialog 
// 2. Add the Error in a new Sheet 


var GSheet = SpreadsheetApp.getActiveSpreadsheet();
var ErrorSheetName = 'ErrorLog';
var headers = ['SNo','Date','Error Detail '];

function ErrorLog()
{

  this.LogError = function(resp)
  {
    var newSheet;
    if(GSheet.getSheetByName(ErrorSheetName)){
      newSheet = GSheet.getSheetByName(ErrorSheetName);
    } else {
      newSheet = GSheet.insertSheet(ErrorSheetName);
      if(headers != null){
        newSheet.getRange(2,1,newSheet.getMaxRows()-1,newSheet.getMaxColumns()).setFontSize("8");
        headersRange = newSheet.getRange(1, 1, 1, headers.length);
        headersRange.setValues([headers]);
        headersRange.setBackground("#0C8EFF").setFontColor("#ffffff").setFontSize("9");
      }
    }

    var lastRow = newSheet.getLastRow();
    newSheet.appendRow([lastRow-1,Date(),resp]);


  };



}
