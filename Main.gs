const ss = SpreadsheetApp.getActiveSpreadsheet();

function onOpen() {
  //Created By Kennen Lawrence
  renderMenu();
  var date = new Date();
  var sheetMonth = ss.getSheetByName('SNAPSHOT').getRange(2, 1).getValue().getMonth();
  
  if (date.getMonth() === sheetMonth) {
    var sheets = ss.getSheets();
    ss.getSheetByName('SNAPSHOT').activate();
    SpreadsheetApp.flush();

    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetName().indexOf(date.getDate().toString()) !== -1) {
        sheets[i].activate();
        break;
      }
    }
  }
  
  ss.toast('The spreadsheet has loaded successfully! Have a great day!', 'Complete!');
  ss.getSheetByName('calc').hideSheet();
}

function renderMenu() {
  var ui = SpreadsheetApp.getUi();
  var subMenu = ui.createMenu('Help')
    .addItem('By Phone','menuItem1')
    .addItem('By Email','menuItem2');
  ui.createMenu('Utilities')
    .addSubMenu(subMenu)
    .addItem('Refresh Menu Counts','menuRefresh')
    .addItem('SNAPSHOT', 'snapshot')
    .addToUi();
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
    MailApp.sendEmail('kennen.lawrence@schomp.com','HELP Daily Recap', input.getResponseText(), {name: getName()});
    SpreadsheetApp.getActiveSpreadsheet().toast('Email sent successfully! Your issue should be addressed within a few hours! For more immediate assistance, contact Kennen by phone.','Email Sent');
  }
}

function snapshot() {
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
  var ignore      = ['SNAPSHOT', 'calc', 'Master'];
  var special     = ['Deposit Log'];
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var sheets      = ss.getSheets();
  var master      = ss.getSheetByName('Master');
  var formatRange = master.getRange(1, 1, master.getMaxRows(), master.getMaxColumns());
  var sheet, numCol, dates;
  
  // Update the month to the next month
  var range = ss.getSheetByName(ignore[0]).getRange(2, 1);
  var date  = range.getDisplayValue().split('/');

  date = date.map(function (value) {
    return parseInt(value, 10);
  });

  if (date[0] === 12) {
    date[0] = 0;
    date[2] ++;
  }

  date[0] ++;
  date = date.join('/');
  range.setValue(date);

  sheets.forEach(function (sheet) {
    if (special.indexOf(sheet.getSheetName()) !== -1) {
      sheet.showSheet();
      numCol = 6;
    } else if (ignore.indexOf(sheet.getSheetName()) === -1) {
      sheet.showSheet();
      numCol = 11;
      sheet.getRange(3, 14, sheet.getLastRow() - 2, sheet.getLastColumn() - 13).setValue('');
      sheet.getRange(3, 14, sheet.getLastRow() - 2, sheet.getLastColumn() - 13).clearNote();
      formatRange.copyFormatToRange(sheet, 1, master.getMaxColumns(), 1, master.getMaxRows());
    } else {
      return
    }

    sheet.getRange(3, 2, sheet.getLastRow() - 2, numCol).setValue('');
    sheet.getRange(3, 2, sheet.getLastRow() - 2, numCol).clearNote();

    protectRanges(sheet);

    ss.toast('Wiped sheet "' + sheet.getSheetName() +'"', 'Completed:');
  });
  
  sheet = ss.getSheetByName(ignore[0]);
  dates = sheet.getRange(2, 1, 31, 2).getDisplayValues();
  const offDays = ['Sunday', 'Monday']

  dates.forEach(function (date, index) {
    // dates is a 2d array so actual date is at the 0th position
    var weekday = parseInt(date[0].split('/')[1], 10);

    // If value < index then next month is bleeding in
    if (weekday < index || offDays.includes(date[1])) sheets[index + 2].hideSheet();
  });

  ss.getSheetByName(ignore[1]).hideSheet();
  menuRefresh();
  updateTradePVR();
  updateDailyGoals();
  refreshDataValidation();
}

function updateTradePVR() {
  var ui = SpreadsheetApp.getUi();
  var link = ui.prompt('Trade PVR Link', 'Please paste the link for the new Trade PVR sheet in the box below:', ui.ButtonSet.OK_CANCEL);
  
  if (link.getSelectedButton() === ui.Button.CANCEL) return;
  
  ss.getSheetByName('SNAPSHOT')
      .getRange(36, 16)
      .setValue('=IMPORTRANGE("' + link.getResponseText() + '","Info!A3:F6")');
}

function protectRanges(sheet) {
  var editWithWarningRanges = ['A:A', '1:2'];
  var ignore = ['SNAPSHOT', 'calc', 'Deposit Log', 'BROKER SHEET'];
  var sheets = sheet ? [sheet] : SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var me = Session.getEffectiveUser();
  
  sheets = sheets.filter(function (sheet) {
    return ignore.indexOf(sheet.getSheetName()) === -1;
  });
  
  sheets.forEach(function (sheet) {
    var range = sheet.getRange('M:M');
    var protection;
    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (protection) {
      if (protection.canEdit()) protection.remove();
    });
    
    
    protection = range.protect().setDescription('Using formulas! Do NOT delete entire rows. Instead only clear out the data.');
    protection.removeEditors(protection.getEditors());
    protection.addEditor(me);
    
    editWithWarningRanges.forEach(function (range) {
      sheet.getRange(range).protect().setWarningOnly(true);
    });
  });
}

function fillFormulas() {
  var sheets = ss.getSheets();
  var ignore = ['SNAPSHOT', 'calc', 'Deposit Log', 'BROKER SHEET'];
  var sheet, range, values, formula;
  var columns = [1, 1, 12];
  var rows = [2, 1, 1];
  var referenceColumns = ['A', 'B', 'C'];
  
  if (columns.length !== rows.length || rows.length !== referenceColumns.length) {
    throw 'Arrays are not equal. Data is missing!';
  }
  
  var number     = 2;
  
  sheets.forEach(function (sheet) {
    var name = sheet.getSheetName();
    
    if (ignore.indexOf(name) !== -1) return;
    
    rows.forEach(function (row, index) {
      formula = '=SNAPSHOT!' + referenceColumns[index] + number;
      sheet.getRange(row, columns[index]).setValue(formula);
    });
    number ++;
  });
}

function updateDailyGoals () {
  var dailyGoals = {
    sunday: 0,
    monday: 0,
    tuesday: 13,
    wednesday: 13,
    thursday: 13,
    friday: 15,
    saturday: 20
  };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SNAPSHOT');
  var day = sheet.getRange(2, 1).getValue().getDay();
  var values = [];
  
  for (var i = 0; i < 31; i++) {
    values.push([dailyGoals[Object.keys(dailyGoals)[day]]]);
    day = day === 6 ? 0 : day + 1;
  }
  
  sheet.getRange(2, 3, values.length).setValues(values);
}

function refreshDataValidation() {
  var ignore = ['SNAPSHOT', 'calc', 'Deposit Log', 'BROKER SHEET'];
  var validationCols = [5, 17, 24];
  var referenceCols = [24, 26, 25];
  var rules = [];
  var sheet = ss.getSheetByName('SNAPSHOT');

  referenceCols.forEach(function (referenceCol) {
    var range = sheet.getRange(2, referenceCol, sheet.getLastRow());
    rules.push(SpreadsheetApp.newDataValidation().requireValueInRange(range, true));
  });

  ss.getSheets().forEach(function (sheet) {
    if (ignore.indexOf(sheet.getSheetName()) !== -1) return;

    validationCols.forEach(function (column, index) {
      sheet.getRange(3, column, sheet.getLastRow() - 2).setDataValidation(rules[index]);
    });
    
    ss.toast('Successfully updated validation!', 'Sheet: ' + sheet.getSheetName(), 5);
  });
}

function refreshManagerValidation() {
  var ignore = ['SNAPSHOT', 'Deposit Log', 'BROKER SHEET', 'calc'];
  var sheet = ss.getSheetByName('SNAPSHOT');
  var range = sheet.getRange(2, 25, sheet.getLastRow() - 1);
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(range, true).build();

  ss.getSheets().forEach(function (sheet) {
    if (ignore.indexOf(sheet.getSheetName()) !== -1) return;

    sheet.getRange(3, 22, sheet.getLastRow() - 2).setDataValidation(rule);
    ss.toast(sheet.getSheetName(), sheet.getSheetName());
  });
}

function correctFormat() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var ignore = ['SNAPSHOT', 'Deposit Log', 'BROKER SHEET', 'calc', 'Master'];
  var master      = ss.getSheetByName('Master');
  var formatRange = master.getRange(1, 1, master.getMaxRows(), master.getMaxColumns());

  ss.getSheets().forEach(function (sheet) {
    if (ignore.indexOf(sheet.getSheetName()) !== -1) return;

    formatRange.copyFormatToRange(sheet, 1, master.getMaxColumns(), 1, master.getMaxRows());
  })
}

function getGeniusValidation(names) {
  names = names.filter(function (name) { return name[0] !== ''});
  names = names.map(function (name) {
    var initials = [];
    
    if (name[0].toUpperCase() === 'KENNETH SCHROEDER') return ['Yes - KJ'];
    
    name = name[0].split(' ');
    name.forEach(function (initial) {
      Logger.log(initial);
      initials.push(initial[0].toString().toUpperCase());
    });
    
    return ['Yes - ' + initials.join('')];
  });
  names.push('No');
  names.push('N/A');
  
  return names
}

function getParsedTeams(teams) {
  teams = teams.filter(function (team) { return team[0] !== ''});
  teams = teams.map(function (team) {
    team = team[0].split(' ');
    team = team.length === 1 ? team[0] : team[1];
    
    if (team === 'TB') team = 'ATB';
    
    return [team];
  });
  
  return teams;
}

function freshStartFromMaster() {
  var ignoreSheets = ['Master', 'SNAPSHOT', 'Deposit Log', 'BROKER SHEET', 'calc']
  var masterSheet = ss.getSheetByName(ignoreSheets[0]);
  var sheets = ss.getSheets();
  var sheetNames = ['1st', '2nd', '3rd', '4th', '5th', '6th', '7th', '8th', '9th', '10th', '11th', '12th', '13th', '14th', '15th', '16th', '17th', '18th', '19th', '20th', '21st', '22nd', '23rd', '24th', '25th', '26th', '27th', '28th', '29th', '30th', '31st'];
  
  sheets.forEach(function (sheet) {
    if (ignoreSheets.indexOf(sheet.getSheetName()) === -1) ss.deleteSheet(sheet);
  });
  
  sheetNames.forEach(function (sheetName, index) {
    var newSheet = masterSheet.copyTo(ss).setName(sheetName);
    newSheet.activate()
    newSheet.getRange(1, 1, 2).setValues([
      ['=SNAPSHOT!B' + (index + 2)],
      ['=SNAPSHOT!A' + (index + 2)]
    ]);
    newSheet.getRange(1, 13).setValue('=SNAPSHOT!C' + (index + 2));
    SpreadsheetApp.flush();
    ss.moveActiveSheet(index + 3);
  });
  
  refreshSnapshotFormulas();
  newMonth();
}

function addNewColumn() {
  var newColumn = {
    number: 26,
    header: 'Time Dropped'
  }
  var ss               = SpreadsheetApp.getActiveSpreadsheet();
  var skipSheets      = ['SNAPSHOT', 'Deposit Log', 'BROKER SHEET', 'calc', 'Master'];
  var sheet            = ss.getSheetByName('Master');
  var formatting       = sheet.getRange(1, newColumn.number, sheet.getLastRow());
  var columnWidth      = sheet.getColumnWidth(newColumn.number);
  var rule             = SpreadsheetApp.newDataValidation().requireCheckbox().build();


  ss.getSheets().forEach(function (sheet) {
    if (skipSheets.indexOf(sheet.getSheetName()) !== -1) return;

    sheet.insertColumnBefore(newColumn.number);
    SpreadsheetApp.flush();
    formatting.copyFormatToRange(sheet, newColumn.number, newColumn.number, 1, sheet.getLastRow());
    sheet.getRange(2, newColumn.number).setValue(newColumn.header);
    sheet.setColumnWidth(newColumn.number, columnWidth);
    sheet.getRange(3, newColumn.number, sheet.getLastRow() - 2).setDataValidation(rule);
//    conditionalRule1 = conditionalRule1.setRanges([range]).build();
//    conditionalRule2 = conditionalRule2.setRanges([range]).build();
//    sheet.setConditionalFormatRules([conditionalRule1, conditionalRule2].concat(sheet.getConditionalFormatRules()));
//    return;
  });
}

function refreshSnapshotFormulas() {
  var sheet = ss.getSheetByName('SNAPSHOT');
  var range = sheet.getRange('D2:W32');
  var formulas = range.getFormulas();
  
  range.setValue('');
  SpreadsheetApp.flush();
  range.setValues(formulas);
}

function endMonth() {
  const ignore = ['SNAPSHOT', 'calc', 'Master']
  
  ss.getSheets().forEach(function(sheet) {
    if (ignore.includes(sheet.getSheetName())) return
    
    sheet.protect().setDescription('A sheet for the new month has been created. This sheet no longer reflects the current date range.').setWarningOnly(true)
  })
}
