function onOpen() {
  //Created By Kennen Lawrence
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ui.createMenu('Utilities').addSubMenu(ui.createMenu('Help').addItem('By Phone','menuItem1').addItem('By Email','menuItem2')).addItem('Refresh Menu Counts','menuRefresh')
  .addItem('SNAPSHOT', 'snapshot').addToUi();
  var month      = new Date().getMonth();
  var sheetMonth = ss.getSheetByName('SNAPSHOT').getRange(2, 1).getValue().getMonth();
  
  if (month === sheetMonth) {
    ss.getSheetByName('SNAPSHOT').activate();
    SpreadsheetApp.flush();
    var day    = new Date().getDate();
    var sheets = ss.getSheets();
    
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetName().indexOf(day) !== -1) { sheets[i].activate(); break; }
    }
  }
  
  ss.toast('The spreadsheet has loaded successfully! Have a great day!', 'Complete!');
  ss.getSheetByName('calc').hideSheet();
}

function menuItem1() {
  //Created By Kennen Lawrence
  SpreadsheetApp.getUi().alert('Call or text (720) 317-5427');
}

function menuItem2() {
  //Created By Kennen Lawrence
  var ui = SpreadsheetApp.getUi();
  var input = ui.prompt('Send Email:','Describe the issue you\'re having in the box below, then press "Ok" to submit your issue via email:', ui.ButtonSet.OK_CANCEL);
  
  if (input.getSelectedButton() === ui.Button.OK) {
    MailApp.sendEmail('kennen.lawrence@schomp.com','HELP Sales Daily_January', input.getResponseText(), { name: getName() });
    SpreadsheetApp.getActiveSpreadsheet().toast('Email sent successfully! Your issue should be addressed within a few hours! For more immediate assistance, contact Kennen by phone.','Email Sent');
  }
}

function snapshot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName('SNAPSHOT'));
}

function menuRefresh() {
  //Created By Kennen Lawrence
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("calc");
  var range = sheet.getRange("M2");
  var data  = range.getValue();
  range.setValue(parseInt(data, 10) + 1);
}

function getName() {
  //Created By Kennen Lawrence
  var email = Session.getActiveUser().getEmail();
  var name  = email.split('@')[0].split('.');
  var first = name[0][0].toUpperCase() + name[0].substring(1);
  var last  = name[1][0].toUpperCase() + name[1].substring(1);
  return first + ' ' + last;
}

function newMonth() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var sheets      = ss.getSheets();
  var ignore      = ['SNAPSHOT', 'calc', 'Master'];
  var special     = ['Deposit Log'];
  var master      = ss.getSheetByName('Master');
  var formatRange = master.getRange(1, 1, master.getMaxRows(), master.getMaxColumns());
  var sheet, numCol;
  
  // Update the month to the next month
  var range = ss.getSheetByName(ignore[0]).getRange(2, 1);
  var year  = range.getDisplayValue().split('/');
  
  for (var i = 0; i < year.length; i++) { year[i] = parseInt(year[i], 10); }
  
  if (year[0] === 12) {
    year[0] = 1;
    year[2] ++;
  } else {
    year[0] ++;
  }
  
  year = year.join('/');
  range.setValue(year);
  
  for (var i = 0; i < sheets.length; i++) {
    sheet = sheets[i];
    
    if (special.indexOf(sheet.getSheetName()) !== -1) {
      sheet.showSheet();
      numCol = 6;
    } else if (ignore.indexOf(sheet.getSheetName()) === -1) {
      sheet.showSheet();
      numCol = 10;
      sheet.getRange(3, 13, sheet.getLastRow() - 2, sheet.getLastColumn() - 12).setValue('');
      sheet.getRange(3, 13, sheet.getLastRow() - 2, sheet.getLastColumn() - 12).clearNote();
      formatRange.copyFormatToRange(sheet, 1, master.getMaxColumns(), 1, master.getMaxRows());
    } else {
      continue;
    }
    
    sheet.getRange(3, 2, sheet.getLastRow() - 2, numCol).setValue('');
    sheet.getRange(3, 2, sheet.getLastRow() - 2, numCol).clearNote();
    ss.toast('Wiped sheet "' + sheet.getSheetName() +'"', 'Completed:');
  }
  
  sheet = ss.getSheetByName(ignore[0]);
  range = sheet.getRange(2, 1, 31, 2).getDisplayValues();
  
  for(i = 0; i < range.length; i++) {
    var value = parseInt(range[i][0].split('/')[1]);
    
    if (value < i || range[i][1] === 'Sunday') sheets[i + 2].hideSheet();
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
      range = ss.getSheetByName(sheets[i].getSheetName()).getRange("L:L");
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
  var columns = [1, 1, 12];
  var rows = [2, 1, 1];
  var referenceColumns = ['A', 'B', 'C'];
  
  if (columns.length != rows.length || rows.length != referenceColumns.length) {
    throw 'Arrays are not equal. Data is missing!';
    return;
  }
  
  var preFormula = "=SNAPSHOT!";
  var number     = 2;
  
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
  var validationCols = [4, 16, 22];
  var referenceCols = [24, 26, 25];
  var sheet, range, end, rule;
  rule = [];
  sheet = ss.getSheetByName('SNAPSHOT');
  
  for (var i = 0; i < referenceCols.length; i++) {
    range = sheet.getRange(2, referenceCols[i], sheet.getLastRow());
    rule[i] = SpreadsheetApp.newDataValidation().requireValueInRange(range, true);
  }
  
  for (i = 0; i < sheets.length; i++) {
    sheet = sheets[i];
    
    if (ignore.indexOf(sheet.getSheetName()) === -1) {
      for (var j = 0; j < validationCols.length; j++) {
        range = sheet.getRange(3, validationCols[j], sheet.getLastRow() - 2);
        range.setDataValidation(rule[j]);
      }
    }
  }
}

function refreshManagerValidation() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var sheets      = ss.getSheets();
  var skip_sheets = ['SNAPSHOT', 'Deposit Log', 'BROKER SHEET', 'calc'];
  var sheet       = ss.getSheetByName('SNAPSHOT');
  var range       = sheet.getRange(2, 25, sheet.getLastRow() - 1);
  var rule        = SpreadsheetApp.newDataValidation()
                       .requireValueInRange(range, true)
                       .build();
  
  for (var i = 0; i < sheets.length; i++) {
    sheet = sheets[i];
    
    if (skip_sheets.indexOf(sheet.getSheetName()) !== -1) continue;
    range = sheet.getRange(3, 22, sheet.getLastRow() - 2)
    range.setDataValidation(rule);
    ss.toast(sheet.getSheetName(), sheet.getSheetName());
  }
}

function correctFormat() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var sheets      = ss.getSheets();
  var skip_sheets = ['SNAPSHOT', 'Deposit Log', 'BROKER SHEET', 'calc', 'Master'];
  var master      = ss.getSheetByName('Master');
  var formatRange = master.getRange(1, 1, master.getMaxRows(), master.getMaxColumns());
  
  for (var i = 0; i < sheets.length; i++) {
    if (skip_sheets.indexOf(sheets[i].getSheetName()) !== -1) continue;
    
    formatRange.copyFormatToRange(sheets[i], 1, master.getMaxColumns(), 1, master.getMaxRows());
  }
}

function addNewColumn() {
  var ss               = SpreadsheetApp.getActiveSpreadsheet();
  var sheets           = ss.getSheets();
  var skip_sheets      = ['SNAPSHOT', 'Deposit Log', 'BROKER SHEET', 'calc', '26th', '27th', '29th', '30th', '31st'];
  var sheet            = ss.getSheetByName('26th');
  var formatting       = sheet.getRange(1, 24, sheet.getLastRow());
  var columnWidth      = sheet.getColumnWidth(2);
//  var range       = sheet.getRange(2, 24, sheet.getLastRow() - 1, 2);
//  var rule             = SpreadsheetApp.newDataValidation()
//                       .requireValueInList(['Ready', 'Issue'], true)
//                       .setHelpText('Ignore the warning you recieve if entering a name that is not present in the dropdown.')
//                       .build();
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    
    if (skip_sheets.indexOf(sheet.getSheetName()) !== -1) continue;
    
    sheet.insertColumnBefore(24);
    formatting.copyFormatToRange(sheet, 24, 24, 1, sheet.getLastRow());
    sheet.getRange(2, 24).setValue('Time Dropped');
    sheet.setColumnWidth(24, columnWidth);
//    var range = sheet.getRange(3, 2, sheet.getLastRow() - 2);
//    range.setDataValidation(rule);
//    conditionalRule1 = conditionalRule1.setRanges([range]).build();
//    conditionalRule2 = conditionalRule2.setRanges([range]).build();
//    sheet.setConditionalFormatRules([conditionalRule1, conditionalRule2].concat(sheet.getConditionalFormatRules()));
//    return;
  }
}




