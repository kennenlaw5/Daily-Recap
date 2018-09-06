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

function newMonth() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var current, range, pass;
  var ignore = ["SNAPSHOT","calc"];
  
//  Update the month to the next month
  range = ss.getSheetByName(ignore[0]).getRange(2, 1);
  current = range.getDisplayValue();
  pass = parseInt(current.split("/")[0])+1;
  if(pass>10){ current = pass + current.substring(2); }
  else{ current = pass + current.substring(1); }
  range.setValue(current);
  for(var i=0;i<sheets.length;i++){
    pass = true;
    current = sheets[i].getSheetName();
    for(var j=0;j<ignore.length;j++){ if(current==ignore[j]){ pass = false; } }
    if(pass){
      current = ss.getSheetByName(current);
      current.showSheet();
      current.getRange(3, 2, current.getLastRow(), 9).setValue("");
      current.getRange(3, 2, current.getLastRow(), 9).clearNote();
      SpreadsheetApp.flush();
      current.getRange(3, 12, current.getLastRow(), current.getLastColumn()-11).setValue("");
      current.getRange(3, 12, current.getLastRow(), current.getLastColumn()-11).clearNote();
      SpreadsheetApp.flush();
      ss.toast('Wiped sheet "' + current.getSheetName() +'"', 'Completed:');
    }
  }
  current=ss.getSheetByName(ignore[0]);
  range = current.getRange(2, 1, 31, 2).getDisplayValues();
  for(i=0;i<range.length;i++){
    pass=parseInt(range[i][0].split("/")[1]);
    if(pass<i || range[i][1]=="Sunday"){ ss.getSheetByName(sheets[i+1].getSheetName()).hideSheet(); }
  }
  menuRefresh();
  ss.getSheetByName(ignore[1]).hideSheet();
  updateTradePVR();
}

function updateTradePVR(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var link = ui.prompt('Trade PVR Link', 'Please paste the link for the new Trade PVR sheet in the box below:', ui.ButtonSet.OK_CANCEL);
  if(link.getSelectedButton()==ui.Button.CANCEL){ return; }
  ss.getSheetByName('SNAPSHOT').getRange(39, 16).setValue('=IMPORTRANGE("'+link.getResponseText()+'","Info!A3:F6")');
}