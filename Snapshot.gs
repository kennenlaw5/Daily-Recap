function snapshotDriver() {
  return {
    columns: {
      genius: 24,
      manager: 19,
      fandiGross: 12,
      rebate: 18 
    },
    teams: [
        'Ben Wegener',
        'Ben Brahler',
        'Joshua Buchanan',
        'Ace Taylor Brown',
        'MER'
      ],
  };
}

function snapshot31 (x,y) {
  return snapshotCore(['23rd', '24th', '25th', '26th', '27th', '28th', '29th', '30th', '31st']);
}

function snapshot22(x,y) {
  return snapshotCore(['16th', '17th', '18th', '19th', '20th', '21st', '22nd']);
}

function snapshot15(x,y) {
  return snapshotCore(['8th', '9th', '10th', '11th', '12th', '13th', '14th', '15th']);
}

function snapshot7(x,y) {
  return snapshotCore(['1st', '2nd', '3rd', '4th', '5th', '6th', '7th']);
}
//MGR=col 18; F_I=12; type=0
//[0] mike[1] jeff[2] chris[3] dean[4]

function snapshotCore(sheetNames) {
  //Created By Kennen Lawrence
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var driver = snapshotDriver();
  var sheet, sheetName, range, sheet;
  var newCount  = [0, 0, 0, 0, 0, 0];
  var newF_I    = [0, 0, 0, 0, 0, 0];
  var cpoCount  = [0, 0, 0, 0, 0];
  var cpoF_I    = [0, 0, 0, 0, 0];
  var usedCount = [0, 0, 0, 0, 0];
  var usedF_I   = [0, 0, 0, 0, 0];
  var newPvr    = [];
  var usedPvr   = [];
  var cpoPvr    = [];
  var team;
  
  sheetNames.forEach(function (sheetName) {
    sheet = ss.getSheetByName(sheetName);

    if (!sheet) return;
    
    rangeValues = sheet.getRange('D3:AB53').getValues();

    rangeValues.forEach(function (rowValues) {
      if (rowValues[0] === '' && rowValues[driver.columns.manager] === '') return;
      var value;
      
      team = rowValues[driver.columns.manager].toString().replace('-', ' ');

      value = rowValues[driver.columns.genius].toLowerCase();
      if (value.indexOf('yes') !== -1) newCount[newCount.length - 1] ++;
      else if (value === 'no') newF_I[newF_I.length - 1] ++;

      team = driver.teams.indexOf(team);

      value = parseInt(rowValues[driver.columns.fandiGross]) || 0;
      if (rowValues[0].toLowerCase() === 'n' && team !== -1) {
        newCount[team] ++;
        newF_I[team] += value;
      }
      else if (rowValues[0].toLowerCase() === 'u' && team !== -1) {
        usedCount[team] ++;
        usedF_I[team] += value;
      }
      else if (rowValues[0].toLowerCase() === 'c' && team !== -1) {
        cpoCount[team] ++;
        cpoF_I[team] += value;
      }
    });
  });
  
  return [newCount, newF_I, '', cpoCount, cpoF_I, '', usedCount, usedF_I];
}