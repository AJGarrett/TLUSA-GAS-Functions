var completedBranches = 0;
var sheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Woodlands 2.0");
var write = "n";
var advSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Advancement");

function createAdvancementReport() {
    var row = 8;
    var endRow = Browser.inputBox("Where Should I End?", Browser.Buttons.OK);
    write = Browser.inputBox("Should I write changes? (y/n)", Browser.Buttons.OK);
    ClearLog();
    var logRow = 1;
    
    do {
      var dataRange = sheet.getRange(row, 1, 1, 52)
      var data = dataRange.getValues();
      
      if (data[0][1].length > 0 && data[0][4] == "Active") {
        var awards = "";
        completedBranches = data[0][48];
        startCompletedBranches = data[0][48]
        //Joining Patch
        if (data[0][5].toString().toLowerCase() != "x") { 
          awards += data[0][2] + " Joining Patch"; 
          writeUpdate(row, 5);
        }
        //Hobbies
        awards += calculateBranch(awards, "Hobbies", 1,2,1, data[0][6], data[0][7], data[0][8], data[0][9].toString().toLowerCase(), data[0][10].toString().toLowerCase(), data[0][50].toString().toLowerCase(), 9, row);
        //Heritage
        awards += calculateBranch(awards, "Heritage", 2,1,1, data[0][12], data[0][13], data[0][14], data[0][15].toString().toLowerCase(), data[0][16].toString().toLowerCase(), data[0][50].toString().toLowerCase(), 15, row);
        //Life Skills
        awards += calculateBranch(awards, "Life Skills", 3,1,1, data[0][18], data[0][19], data[0][20], data[0][21].toString().toLowerCase(), data[0][22].toString().toLowerCase(), data[0][50].toString().toLowerCase(), 21, row);
        //Outtdoor Skills
        awards += calculateBranch(awards, "Outdoor Skills", 3,1,1, data[0][30], data[0][31], data[0][32], data[0][33].toString().toLowerCase(), data[0][34].toString().toLowerCase(), data[0][50].toString().toLowerCase(), 33, row);
        //Sports & Fitness
        awards += calculateBranch(awards, "Sports & Fitness", 2,1,1, data[0][24], data[0][25], data[0][26], data[0][27].toString().toLowerCase(), data[0][28].toString().toLowerCase(), data[0][50].toString().toLowerCase(), 27, row);
        //Sci/Tech
        awards += calculateBranch(awards, "Science & Technology", 2,1,1, data[0][36], data[0][37], data[0][38], data[0][39].toString().toLowerCase(), data[0][40].toString().toLowerCase(), data[0][50].toString().toLowerCase(), 39, row);
        //Values
        awards += calculateBranch(awards, "Values", 3,1,1, data[0][42], data[0][43], data[0][44], data[0][45].toString().toLowerCase(), data[0][46].toString().toLowerCase(), data[0][50].toString().toLowerCase(), 45, row);
        //Forest
        if (completedBranches == 7 && completedBranches != startCompletedBranches) {
            awards += ", Forest Award";
            writeUpdate(row, 50);
        }
        
        //Worthy Life
        if (data[0][51].toString().toLowerCase() == "c") {
          var comma = "";
          if (awards.length > 0) { comma = ", "; }
          awards += comma + "Worthy Life Award";
          if (write == "y") {
            sheet.getRange(row, 52).setValue("R");
          }
        }
        
        
        
        //Write to Advancement Log
        if (awards.length > 0) {
          Logger.log(data[0][1] + " - " + data[0][2] + " - " + awards); 
          advSheet.getRange(logRow, 1).setValue(data[0][1]);
          advSheet.getRange(logRow, 2).setValue(data[0][2]);
          advSheet.getRange(logRow, 3).setValue(awards);
          logRow++;
        }
        
        //Update Sheet
      }
      
      
      row++;
    } while (row <= endRow)
}


function calculateBranch (awards, branch, coreReq, electiveReq, httReq, coreComp, electiveComp, httComp, prevAwardedBranch, prevAwardedStar, Forest, column, row)  {
  if (prevAwardedStar != "x") {
    var awardType = "Branch";
    if (Forest == "x") {
       coreReq = coreReq * 2;
       electiveReq = electiveReq * 2;
       httReq = httReq * 2;
       awardType = "Star"
    }
    
    if (prevAwardedBranch != "x" || Forest == "x") {
      if (coreComp >= coreReq && electiveComp >= electiveReq && httComp >= httReq) {
         var comma = "";
         if (awards.length > 0) { comma = ", "; }
         if (Forest != "x") {
           completedBranches++;
           writeUpdate(row, column)
         } else {
            writeUpdate(row, column+1)
         }
         Logger.log(awards + comma + branch + " " + awardType);
         return comma + branch + " " + awardType;
      } else {
        return "";
      }
    } else {
      return ""; 
    }
  } else {
    return ""; 
  }
    
}
  
  function writeUpdate(row, column) {
    if (write == "y") {
      sheet.getRange(row, column+1).setValue("x");
    }
  }

  function ClearLog() {
    if (advSheet.getLastRow() > 0) {
      advSheet.insertRowAfter(advSheet.getLastRow());
    }
    //Delete all Existing Rows
    if (advSheet.getLastRow() > 0)
      advSheet.deleteRows(1, advSheet.getLastRow());
  }