function snapshotDriver() {
  return {
    columns: {
      genius: 24,
      manager: 19,
      fandiGross: 12,
      rebate: 18 
    },
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
  
  for (var i = 0; i < sheetNames.length - 2; i++) {
    sheetName = sheetNames[i];
    sheet     = ss.getSheetByName(sheetName);
    
    if (sheet) {
      range = sheet.getRange('C3:AA53').getValues();
      
      for (var j = 0; j < sheet.getLastRow() - 2; j++) {
        if (range[j][0] != '' && range[j][driver.columns.manager] != '') {
          team = range[j][driver.columns.manager].toString().replace('-', ' ');
          
          if (range[j][driver.columns.genius].toLowerCase().indexOf('yes') !== -1) newCount[newCount.length - 1] ++;
          else if (range[j][driver.columns.genius].toLowerCase() === 'no') newF_I[newF_I.length - 1] ++;
          
          if (team === 'Ben Wegener')           { team = 0; }
          else if (team === 'Ben Brahler')      { team = 1; }
          else if (team === 'Joshua Buchanan')  { team = 2; }
          else if (team === 'Ace Taylor Brown') { team = 3; }
          else if (team === 'MER')              { team = 4; }
          else                                  { team = 5; }
          
          if (range[j][0].toLowerCase() == 'n' && team != 5) {
            newCount[team] ++; 
            
            if (!isNaN(parseInt(range[j][driver.columns.fandiGross]))) {
              newF_I[team] += parseInt(range[j][driver.columns.fandiGross]); 
            } 
          } else if (range[j][0].toLowerCase() == 'u' && team != 5) {
            usedCount[team] ++; 
            
            if (!isNaN(parseInt(range[j][driver.columns.fandiGross]))) {
              usedF_I[team] += parseInt(range[j][driver.columns.fandiGross]);
            }
          } else if (range[j][0].toLowerCase() == 'c' && team != 5) {
            cpoCount[team] ++; 
            
            if (!isNaN(parseInt(range[j][driver.columns.fandiGross]))) {
              cpoF_I[team] += parseInt(range[j][driver.columns.fandiGross]);
            }
          }
          
        } else if (range[j + 1][0] == '' &&
                   range[j + 1][driver.columns.rebate] == '' &&
                   range[j + 2][0] == '' &&
                   range[j + 2][driver.columns.rebate] == '' &&
                   range[j + 3][0] == '' &&
                   range[j + 3][driver.columns.rebate] == '') 
        { 
          j = sheet.getLastRow(); 
        }
      }
    }
  }
  
  return [newCount, newF_I, '', cpoCount, cpoF_I, '', usedCount, usedF_I];
}