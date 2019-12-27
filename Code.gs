function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Trail Life Menu')
      .addSubMenu(ui.createMenu('Woodlands')
          .addItem('Send Leader Reminders', 'woodlandReminder')
          .addItem('Record Meeting Attendance', 'recordAttendance'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Advancement')
          .addItem('Send Individual Advancement', 'getRowNumber')
          .addItem('Send All Woodlands Advancement', 'SendAll')
          .addItem('Create Advancement Summary', 'createAdvancementReport'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Coming Soon')
          .addItem('Get Patrol List', 'menuItem2'))
      .addToUi(); 
}


function Testt() {
  Logger.log(Browser.msgBox("Clear Attendance Markers, or leave for another entry?", Browser.Buttons.YES_NO));
}