//changeLog:
//-v0.1 - AGarrett: initial code of all basic functionality
//-v0.2 - MJacobsen: minor tweak of Mt. Lion (K-Z) - was previously set up to send Fox material to Mt Lion A-Z leader and helper. 
//-v0.3 - AGarrett: added sidebar to edit message.

var body = "<p>Hey Guys</p>";
body += "<p>This week's lesson is on {lesson}.   Here is the link to the {plan} and {help} documents.</p>";
body += "<p>Let me know if you need anything.   Also, please ensure the rooms are returned to order when finished.</p>";
body += "<p>Thanks!</p>";
body += "<p>Matt</p>";

var admin = "no";

function woodlandReminder(argForm) {
  //admin = Browser.msgBox("Test Email", Browser.Buttons.YES_NO);
  var date = argForm.txtDate; //Browser.inputBox("Enter Date of Meeting:", Browser.Buttons.OK);
  //Logger.log(date);
  body = argForm.txtBody;
  var subject = argForm.txtSub;
  
  if (argForm.txtAdmin) {
    //Logger.log("Admin");
    admin = "yes";
  }
  
  var sheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Woodlands Schedule");
  var pDataRange = sheet.getRange(4,1, 28, 15);
  
  var pData = pDataRange.getValues();
  for (i in pData) {
    var row = pData[i];
    //Logger.log(Utilities.formatDate(row[0], "GMT+0", "yyyy-MM-dd"));
    if (Utilities.formatDate(row[0], "GMT+0", "yyyy-MM-dd") == date) {
      var sheetRow = parseInt(i)+4;
      //Foxes
      BuildEmail(row, 5, 6, "Fox", date, sheetRow, subject);
      //Hawks 1
      BuildEmail(row, 7, 8, "Hawks (A-J)", date, sheetRow, subject);
      //Hawks 2
      BuildEmail(row, 9, 10, "Hawks (K-Z)", date, sheetRow, subject);
      //Mt. Lions 1
      BuildEmail(row, 11, 12, "Mt. Lions (A-J)", date, sheetRow, subject);
      //Mt. Lions 2
      BuildEmail(row, 13, 14, "Mt. Lions (K-Z)", date, sheetRow, subject);
    }
  }
  
}

function BuildEmail(row, teachIndex, assistantIndex, patrol, date, rowNum, subject) {
 
  var teacherEmail = "";
  var assistantEmail = "";
  if (row[teachIndex] != "") {
    teacherEmail = getParentEmail(row[teachIndex]);
  } else {
    Browser.msgBox("No Teacher for " + patrol); 
  }
  
  if (row[assistantIndex] != "") {
    assistantEmail = getParentEmail(row[assistantIndex]);
  } else {
    Browser.msgBox("No Assistant for " + patrol); 
  }
  
  
  var subject = BuildSubject(subject, patrol, date); 
  var message = BuildMessage(row[2], GETURL(rowNum, 4), GETURL(rowNum, 5));
  var email = Session.getActiveUser().getEmail();
  var testmessage = "<p><b>TEST Email - </b> Email will be sent to: " + teacherEmail + ", " + assistantEmail + "</p>";
    
  if (admin == "yes") {
    message = testmessage + message; 
    
  } else {
    if (teacherEmail != "") { 
      email += "," + teacherEmail; 
    }
    if (assistantEmail != "") { 
      email += "," + assistantEmail; 
    }
  }
  MailApp.sendEmail(email, subject, message, {htmlBody: message});
  
}

function BuildSubject(subject, patrol, date) {
  subject = subject.replace("{patrol}", patrol);
  subject = subject.replace("{date}", date);
  return subject;
}

function BuildMessage(lesson, plan, help) {
 
  body = body.replace(/\n/g, "<br />");
  body = body.replace("{lesson}", lesson);
  body = body.replace("{plan}", "<a href=" + plan + ">Lesson Plan</a>");
  body = body.replace("{help}", "<a href=" + help + ">Help</a>");
  
  return body
  
}

function GETURL(row, column) {
  Logger.log(row + " " + column);
  var range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Woodlands Schedule").getRange(row, column);
  var url = /"(.*?)"/.exec(range.getFormulaR1C1())[0];
  return url;
}

function TeachingReminder() {
     var html = HtmlService.createHtmlOutputFromFile("ReminderText");
     html.setTitle('Teaching Reminder Email'); 
     SpreadsheetApp.getUi().showSidebar(html);
}
