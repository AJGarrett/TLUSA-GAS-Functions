function duesReminder() {
     var html = HtmlService.createHtmlOutputFromFile("DuesReminderForm");
     html.setTitle('Dues Reminder Email'); 
     SpreadsheetApp.getUi().showSidebar(html);
}

function sendDueEmails(argForm) {
  //variable
  var parentsEmailed = [];
  
  //loop though youth sheet
  var sheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Youth");
  var pDataRange = sheet.getRange(8,1, getLastRow("Youth"), 8);
  var pData = pDataRange.getValues();
  var email = '';
  for (i in pData) {
    var row = pData[i];
    if (row[7].toLowerCase() == "" && row[2] == "Active") { 
      var email = getParentEmail(row[3]);
      if (email.indexOf("@") > -1 && parentsEmailed.indexOf(email) == -1) {
        var subject = argForm.txtSub;
        var message = argForm.txtBody;
        parentsEmailed.push(email);
        MailApp.sendEmail(email, subject, message, {htmlBody: message});
      }
      
      
    }
  }
  
  //do the youth have siblings?
  
  //add tto parents emails
  
  //send emailed
  Logger.log(parentsEmailed);
}
  