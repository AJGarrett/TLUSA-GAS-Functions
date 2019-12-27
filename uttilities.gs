function getParentEmail(parent) {
  var sheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Adults");
  var pDataRange = sheet.getRange(2,1,100,5);
  var pData = pDataRange.getValues();
  var email = '';
  for (i in pData) {
    var row = pData[i];
    if (row[0] == parent) { 
      email = row[4];
      break;
    }
  }
  return email;
}

function getParentGroup(trailman) {
  var sheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Youth");
  var pDataRange = sheet.getRange(8,1,120,4);
  var pData = pDataRange.getValues();
  var email = '';
  for (i in pData) {
    var row = pData[i];
    if (row[0].trim() == trailman) { 
      email = row[3];
      break;
    }
  }
  return email;
}

function getParentEmailbyTrailman(trailman) {
  return getParentEmail(getParentGroup(trailman.trim())); 
}


function getLastRow(sheetName) {
  var sheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheetName == "Attendance2") {
    var range = sheet.getRange("B1:B").getValues();
    return range.filter(String).length; 
  } else if (sheetName == "Youth") {
    var range = sheet.getRange("A1:A").getValues();
    return range.filter(String).length; 
  } else if (sheetName == "Woodlands 2.0") {
    var range = sheet.getRange("B1:B").getValues();
    return range.filter(String).length; 
  } else {
    return sheet.getLastRow();
  }
}


function checkQuota() {
  throw 'You have ' + MailApp.getRemainingDailyQuota() + ' emails remaining'; 
}