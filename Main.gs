function onOpen() {
  //Created By Kennen Lawrence
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ui.createMenu('Utilities').addSubMenu(ui.createMenu('Contact Kennen').addItem('By Phone','menuItem1').addItem('By Email','menuItem2')).addItem('Refresh Menu Counts','menuRefresh').addToUi();
  //.addItem('Reset Statistics','reset').addItem('Refresh CA Ranking','rank').addToUi();
  var message = 'The spreadsheet has loaded successfully! Have a great day!';
  var title = 'Complete!';
  ss.toast(message, title);
  ss.getSheetByName("calc").hideSheet();
}
function menuItem1() {
  //Created By Kennen Lawrence
  SpreadsheetApp.getUi().alert('Call or text (720) 317-5427');
}
function menuItem2() {
  //Created By Kennen Lawrence
  var ui = SpreadsheetApp.getUi();
  var input = ui.prompt('Send Email:','Describe the issue you\'re having in the box below, then press "Ok" to submit your issue via email:',ui.ButtonSet.OK_CANCEL);
  if (input.getSelectedButton() == ui.Button.OK) {
    MailApp.sendEmail('kennen.lawrence@schomp.com','HELP Sales Daily_January',input.getResponseText(),{name:getName()});
    SpreadsheetApp.getActiveSpreadsheet().toast('Email sent successfully! Your issue should be addressed within a few hours! For more immediate assistance, contact Kennen by phone.','Email Sent');
  } else if (input.getSelectedButton() == ui.Button.CANCEL) {
    Logger.log('User cancelled');
  }
}
function menuRefresh(){
  //Created By Kennen Lawrence
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName("calc");
  var range=sheet.getRange("N2");
  var data=range.getValue();
  range.setValue(parseInt(data)+1);
}
function getName(){
  //Created By Kennen Lawrence
  //Version 1.0
  var email = Session.getActiveUser().getEmail();
  var name;var first;var last;
  name = email.split("@schomp.com");
  name=name[0];
  name=name.split(".");
  first=name[0];
  last=name[1];
  first= first[0].toUpperCase() + first.substring(1);
  last=last[0].toUpperCase() + last.substring(1);
  name=first+" "+last;
  Logger.log(name);
  return name;
}