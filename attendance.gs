function recordAttendance() {
  var date = Browser.inputBox("Enter Date of Event (mm/dd/yyyy):", Browser.Buttons.OK);
  var type = Browser.inputBox("Enter Branch Type (i.e. Values - Core):", Browser.Buttons.OK);
  var note = Browser.inputBox("Enter note such as location of HTT, if applicable:", Browser.Buttons.OK);
  
  var startRow = getLastRow("Attendance2")+1;
  var added=0;
  
  var attSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance2");
  
  var sheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Woodlands 2.0");
  var pDataRange = sheet.getRange(8,1, getLastRow("Woodlands 2.0"), 2);
  var pData = pDataRange.getValues();
  var email = '';
  for (i in pData) {
    var row = pData[i];
    if (row[0].toLowerCase() == "x") { 
      attSheet.getRange(startRow, 2).setValue(row[1]);
      attSheet.getRange(startRow, 3).setValue(date);
      attSheet.getRange(startRow, 4).setValue(type);
      attSheet.getRange(startRow, 6).setValue(note);
      //Logger.log(i);
      //sheet.getRange(i+8, 1).setValue("");
      
      startRow++;
      added++;
    }
  }
  
  Logger.log(added.toString() + " attendance Records added");
  if (Browser.msgBox("Clear Attendance Markers, or leave for another entry?", Browser.Buttons.YES_NO) == "yes") {
    sheet.getRange(8,1, getLastRow("Woodlands 2.0"), 1).clear();
  }
}
