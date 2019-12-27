
var admin = false;

function getRowNumber() {
  
  var sh = SpreadsheetApp.getActiveSheet();
  var cell = sh.getActiveCell();
  var SelectRow = cell.getRow();
  
  var dataRange = sh.getRange(SelectRow, 1, 1, 52)
  
  var data = dataRange.getValues();
  var test = Browser.msgBox("Testing?", "Should I send to Andy?", Browser.Buttons.YES_NO)
  if (test == "yes") { admin = true; }

  sendEmail(data);
}

function getEmail(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheet2 = sheets[0];
  var email = "";
  var pDataRange = sheet2.getRange(1,1,140, 45);
  var pData = pDataRange.getValues();
  for (i in pData) {
    var row = pData[i];
    if (row[0] == name) { 
      email = row[11];
      break;
    }
  }
  
  return email;
}

function sendEmail(data) {
  
  var email = getParentEmailbyTrailman(data[0][1]);

  if (email.indexOf("@") > -1) {
    var row = data[0];
    Logger.log(admin);
    if (admin == "yes") {
      email = 'andyjgarrett@gmail.com';
    }
    var name = data[0][1].slice(data[0][1].indexOf(",")+1)
    var subject = 'Advancement Update for ' + name;
    var message = BuildAdvEmail(data);
    Logger.log(email);
    MailApp.sendEmail(email, subject, message, {htmlBody: message});
  } else {
    Browser.msgBox("No email for " + data[0][1]); 
  }
  
}

function BuildAdvEmail(data)
{
  var awdsNight = "Dec. 19th";
  var dueDate = "Dec. 5th";
  var email = "<p>Parents,</p><p>Enclosed is the status of your son's advancement in the Woodlands Program.   Our next Awards Night will be <b>" + awdsNight + "</b>.  In order to receive advancement, the following Activities will need to be turned in by the meeting on <b>" + dueDate + "</b>";
  if (data[0][48] < 7) {
    email += "<h4>Progress Towards the Forest Award</h4>";
    email += BuildBranchProgress("Hobbies", data[0][9], 1,2,1, data[0][6] , data[0][7] , data[0][8], data[0][11], 2);
    email += BuildBranchProgress("Heritage", data[0][15], 2,1,1, data[0][12] , data[0][13] , data[0][14], data[0][17], 2);
    email += BuildBranchProgress("Life Skills", data[0][21], 3,1,1, data[0][18] , data[0][19] , data[0][20], data[0][23], 3);
    email += BuildBranchProgress("Outdoor Skills", data[0][33], 3,1,1, data[0][30] , data[0][31] , data[0][32], data[0][35], 0);
    email += BuildBranchProgress("Science & Technology", data[0][39], 2,1,1, data[0][36] , data[0][37] , data[0][38], data[0][41], 3);
    email += BuildBranchProgress("Sports & Fitness", data[0][27], 2,1,1, data[0][24] , data[0][25] , data[0][26], data[0][29], 3);
    email += BuildBranchProgress("Values", data[0][45], 3,1,1, data[0][42] , data[0][43] , data[0][44], data[0][47], 2);
  } else {
    email += "<h4>Progress Towards Slyvan Stars</h4>";
    email += BuildSlyvanProgress("Hobbies", data[0][10], 1,2,1, data[0][6] , data[0][7] , data[0][8], data[0][11], 2);
    email += BuildSlyvanProgress("Heritage", data[0][16], 2,1,1, data[0][12] , data[0][13] , data[0][14], data[0][17], 2);
    email += BuildSlyvanProgress("Life Skills", data[0][22], 3,1,1, data[0][18] , data[0][19] , data[0][20], data[0][23], 3);
    email += BuildSlyvanProgress("Outdoor Skills", data[0][34], 3,1,1, data[0][30] , data[0][31] , data[0][32], data[0][35], 0);
    email += BuildSlyvanProgress("Science & Technology", data[0][40], 2,1,1, data[0][36] , data[0][37] , data[0][38], data[0][41], 3);
    email += BuildSlyvanProgress("Sports & Fitness", data[0][28], 2,1,1, data[0][24] , data[0][25] , data[0][26], data[0][29], 3);
    email += BuildSlyvanProgress("Values", data[0][46], 3,1,1, data[0][42] , data[0][43] , data[0][44], data[0][47], 2);

  }
  
  email += "<p><br /><br />This report is based on meetings through November.   You can view the <a href='https://docs.google.com/spreadsheets/d/e/2PACX-1vTcKnisoUOdtxNQRjRJ6SLYzBOzgDFZx95s-nmQgFzyfMGcAJ-YHwRgtTopXet5Pde01K1o2Px79QF8/pubhtml?gid=0&single=true'>schedule</a> for Woodlands.</p>";//  Trailmen who are will countinue in their current patrol, will have the option to complete branch work next year as well.</p>";
  email += "<p>You can report At-Home activities completed either via email or by bringing a printed copy of the sheet to a meeting with your son(s) name and a signature of what was completed.<p>"; 
  email += "<p>If you have any questions, please let me know and I'll be happy to clarify.   You can see all At-Home Activities, using <a href='https://drive.google.com/open?id=11dw0VxUtwu52jUA_nqJq54QfebarhzFN'>this PDF</a>.</p><p>Walk Worthty, <br /> Andy<p>";
  
  
  return email;
}

function BuildBranchProgress(name, completed, reqCore, reqElective, reqHTT, totalCore, totalElective, totalHTT, AHA, meetingsLeft) {
  return BranchProgress(name, completed, reqCore, reqElective, reqHTT, totalCore, totalElective, totalHTT, AHA, meetingsLeft, 2)
}

function BuildSlyvanProgress(name, completed, reqCore, reqElective, reqHTT, totalCore, totalElective, totalHTT, AHA, meetingsLeft) {
  return BranchProgress(name, completed, reqCore, reqElective, reqHTT, totalCore-reqCore, totalElective-reqElective, totalHTT-reqHTT, AHA, meetingsLeft, 4)
}

function BranchProgress(name, completed, reqCore, reqElective, reqHTT, totalCore, totalElective, totalHTT, AHA, meetingsLeft, AHAOptions) {
  //Browser.msgBox(name + ", " + completed + ", " + total + ", " + totalCompleted + ", " + AHA + ", " + meetingsLeft);
  var total = reqCore + reqElective + reqHTT;
  var totalCompleted = 0;
  if (reqCore <= totalCore) { totalCompleted += reqCore; } else { totalCompleted += totalCore; }
  if (reqElective <= totalElective) { totalCompleted += reqElective; } else { totalCompleted += totalElective; }
  if (reqHTT <= totalHTT) { totalCompleted += reqHTT; } else { totalCompleted += totalHTT; }
  var progress = "<p>" + name + ": ";
  if (completed == "x" || totalCompleted >= total) {
    progress += "<i>Completed</i>";
  } else {
    if (totalCompleted <= 1) {
      progress += "<i>Meetings for this branch will take place later in the program year</i>";
    } else {
      var totalDue = total - totalCompleted;
      var totalAvailable = totalDue - meetingsLeft;
      //  Browser.msgBox(name + " td: " + totalDue);
      Logger.log(name);
      Logger.log("Due: " + totalDue);
      if (meetingsLeft == 0) {
        //can they do AHA?
        AHAOptions = AHAOptions - AHA;
        Logger.log("AHAOptions: " + AHAOptions);
        if (totalDue < AHAOptions+1 && (totalDue - AHA) == 0) {
          var activities = (totalDue) * 2
          //Browser.msgBox(name + " HTT: " + activities);
          progress += "Needs to complete " + activities + " At Home Activities";
        } else {
          if (totalDue <= AHAOptions) {
            var activities = (totalDue)*2;
            Logger.log(activities);
            progress += "You will need to complete " + activities + " At Home Activities.";
          } else {
            progress += "There aren't enough available activities scheduled at this point.  Please discuss options with me.";
          }
        }
      } else {
        if (totalDue < AHAOptions+1 && (totalDue - AHA) == 0) {
          //AHA or Meetings
          var activities = ((totalDue-AHA) + (AHAOptions-AHA)) * 2
          progress += "Needs to attend meetings and complete " + activities + " At Home Activities";
        } else {
          var meetingsRequired = totalDue - (2-AHA);
          var activities = (totalDue-meetingsRequired) * 2;
          Logger.log(meetingsRequired);
          if (meetingsRequired <= meetingsLeft) {
            progress += "Must attend meeting/s and complete " + activities + " At Home Activities"; 
          } else {
            progress += "There aren't enough available activities scheduled at this point.  Please discuss options with me."; 
          }
        }
        
      }
    }
    
    
  }
  
  progress += "</p>";
  return progress;
}

function SendAll() {
  var sh = SpreadsheetApp.getActiveSheet();
  var test = Browser.msgBox("Testing?", "Should I send to Andy?", Browser.Buttons.YES_NO)
  var selected = Browser.msgBox("Only send to selected?", Browser.Buttons.YES_NO)
  if (test == "yes") { admin = true; }
  var row = Browser.inputBox("Where Should I Start?", Browser.Buttons.OK);
  var endRow = Browser.inputBox("Where Should I End?", Browser.Buttons.OK);
  do {
    var dataRange = sh.getRange(row, 1, 1, 52)
    var data = dataRange.getValues();
    
    if (data[0][1].length > 0 && data[0][4] == "Active") {
      if ((selected == "yes" && data[0][0]=="x") || selected == "no") {
        sendEmail(data); 
      }
    }
    row++
  } while (row <= endRow)  
}

function PassData(data, email) {
  Browser.msgBox(data[0][1] + " - " + email)
}