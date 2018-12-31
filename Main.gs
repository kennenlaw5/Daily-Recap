function onOpen() {
  //Created By Kennen Lawrence
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ui.createMenu('Utilities').addSubMenu(ui.createMenu('Help').addItem('By Phone','menuItem1').addItem('By Email','menuItem2')).addItem('Refresh Menu Counts','menuRefresh')
  .addItem('SNAPSHOT', 'snapshot').addToUi();
  var month = new Date().getMonth();
  var sheetMonth = ss.getSheetByName('SNAPSHOT').getRange(2, 1).getValue().getMonth();
  if (month == sheetMonth) {
    ss.getSheetByName('SNAPSHOT').activate();
    SpreadsheetApp.flush();
    ss.getSheetByName('Deposit Log').activate();
    SpreadsheetApp.flush();
    ss.getSheetByName('SNAPSHOT').activate();
    SpreadsheetApp.flush();
    var day = new Date().getDate();
    var sheets = ss.getSheets();
    var name;
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetName().indexOf(day) != -1) { name = sheets[i].getSheetName(); }
    }
    if (name != undefined) { ss.getSheetByName(name).activate(); }
  }
  ss.toast('The spreadsheet has loaded successfully! Have a great day!', 'Complete!');
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

function snapshot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName('SNAPSHOT'));
}

function menuRefresh(){
  //Created By Kennen Lawrence
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName("calc");
  var range=sheet.getRange("M2");
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
  var current, range, pass, year;
  var ignore = ["SNAPSHOT","calc"];
  var special = ['Deposit Log'];
  
//  Update the month to the next month
  range = ss.getSheetByName(ignore[0]).getRange(2, 1);
  current = range.getDisplayValue().split('/');
  
  for (var i = 0; i < current.length; i++) { current[i] = parseInt(current[i], 10); }
  
  if (current[0] == 12) { current[0] = 1; current[2] += 1; } 
  else { current[0] +=1; }
  
  current = current.join('/');
  range.setValue(current);
  for (var i = 0; i < sheets.length; i++) {
    pass = true;
    current = sheets[i].getSheetName();
    for (var j = 0; j < ignore.length; j++) {
      if (current == ignore[j]) {
        pass = false; 
      }
      if (current == special[j]) {
        pass = false;
        current = ss.getSheetByName(current);
        current.showSheet();
        current.getRange(3, 2, current.getLastRow(), 6).setValue("");
        current.getRange(3, 2, current.getLastRow(), 6).clearNote();
        SpreadsheetApp.flush();
        ss.toast('Wiped sheet "' + current.getSheetName() +'"', 'Completed:');
      }
    }
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
    if(pass<i || range[i][1]=="Sunday"){ ss.getSheetByName(sheets[i+2].getSheetName()).hideSheet(); }
  }
  menuRefresh();
  ss.getSheetByName(ignore[1]).hideSheet();
  updateTradePVR();
  updateDailyGoals();
  refreshDataValidation();
  protectRanges();
}

function updateTradePVR(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var link = ui.prompt('Trade PVR Link', 'Please paste the link for the new Trade PVR sheet in the box below:', ui.ButtonSet.OK_CANCEL);
  if(link.getSelectedButton()==ui.Button.CANCEL){ return; }
  ss.getSheetByName('SNAPSHOT').getRange(39, 16).setValue('=IMPORTRANGE("'+link.getResponseText()+'","Info!A3:F6")');
}

function protectRanges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var range, protection, protections;
  var ignore = ["SNAPSHOT","calc","Deposit Log","BROKER SHEET"];
  var me = Session.getEffectiveUser();
  for (var i = 0; i < sheets.length; i++) {
    if (ignore.indexOf(sheets[i].getSheetName()) == -1) {
      protections = ss.getSheetByName(sheets[i].getSheetName()).getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (var j = 0; j < protections.length; j++) {
        protection = protections[j];
        if (protection.canEdit()) {
          protection.remove();
        }
      }
      SpreadsheetApp.flush();
      range = ss.getSheetByName(sheets[i].getSheetName()).getRange("K:K");
      protection = range.protect().setDescription('Using formulas! Do NOT delete entire rows. Instead only clear out the data.');
      protection.removeEditors(protection.getEditors());
      protection.addEditor(me);
      SpreadsheetApp.flush();
    }
  }
}

function fillFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var ignore = ["SNAPSHOT","calc","Deposit Log","BROKER SHEET"];
  var sheet, range, values, formula, name;
  var columns = [1,1,11];
  var rows = [2,1,1];
  var referenceColumns = ['A','B','C'];
  
  if (columns.length != rows.length || rows.length != referenceColumns.length) {
    throw 'Arrays are not equal. Data is missing!';
    return;
  }
  
  var preFormula = "=SNAPSHOT!";
  var number = 2;
  for (var i = 0; i < sheets.length; i++) {
    name = sheets[i].getSheetName();
    if (ignore.indexOf(name) == -1) {
      for (var j = 0; j < rows.length; j++) {
        formula = preFormula + referenceColumns[j] + number;
        sheet = ss.getSheetByName(name);
        range = sheet.getRange(rows[j], columns[j]);
        range.setValue(formula);
      }
      number++;
    }
  }
}

function updateDailyGoals () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('SNAPSHOT');
  var dailyGoals = [0, 15, 13, 17, 15, 25, 35];
  var day = sheet.getRange(2, 1).getValue().getDay();
  var values = [];
  for (var i = 0; i < 31; i++) {
    values[i] = [dailyGoals[day]];
    if (day == 6) { day = 0; }
    else { day++; }
  }
  sheet.getRange(2, 3, values.length).setValues(values);
}

function refreshDataValidation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var ignore = ["SNAPSHOT","calc","Deposit Log","BROKER SHEET"];
  var validationCols = [3, 15, 21];
  var referenceCols = [24, 26, 25];
  var sheet, range, end, rule;
  rule = [];
  sheet = ss.getSheetByName('SNAPSHOT');
  for (var i = 0; i < referenceCols.length; i++) {
    range = sheet.getRange(2, referenceCols[i], sheet.getLastRow());
    rule[i] = SpreadsheetApp.newDataValidation().requireValueInRange(range, true);
  }
  for (i = 0; i < sheets.length; i++) {
    sheet = sheets[i].getSheetName();
    if (ignore.indexOf(sheet) == -1) {
      sheet = ss.getSheetByName(sheet);
      for (var j = 0; j < validationCols.length; j++) {
        range = sheet.getRange(3, validationCols[j], sheet.getLastRow() - 2);
        range.setDataValidation(rule[j]);
      }
    }
  }
}