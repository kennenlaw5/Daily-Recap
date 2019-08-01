function snapshot31 (x,y) {
  //Created By Kennen Lawrence
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  var sheetNames = ['23rd', '24th', '25th', '26th', '27th', '28th', '29th', '30th', '31st'];
  
  for (var i = 0; i < sheetNames.length - 2; i++) {
    sheetName = sheetNames[i];
    sheet     = ss.getSheetByName(sheetName);
    
    if (sheet != null) {
      range = sheet.getRange('C3:Z53').getValues();
      for (var j = 0; j < sheet.getLastRow() - 2; j++) {
        if (range[j][0] != '' && range[j][19] != '') {
          team = range[j][19];
          team = team.toString().replace('-', ' ');
          
          if (range[j][23].toLowerCase().indexOf('yes') !== -1) { newCount[newCount.length - 1] ++; }
          else if (range[j][23].toLowerCase() === 'no') { newF_I[newF_I.length - 1] ++; }
          
          if (team === 'Merrie Thomsen')        { team = 0; }
          else if (team === 'Ben Brahler')      { team = 1; }
          else if (team === 'Liz Liggett')      { team = 2; }
          else if (team === 'Joshua Buchanan')  { team = 3; }
          else if (team === 'Ace Taylor Brown') { team = 4; }
          else                                  { team = 5; }
          
          if (range[j][0].toLowerCase() == 'n' && team != 5) { newCount[team] ++; if (!isNaN(parseInt(range[j][12]))) { newF_I[team] += parseInt(range[j][12]); } }
          else if (range[j][0].toLowerCase() == 'u' && team != 5) { usedCount[team] ++; if (!isNaN(parseInt(range[j][12]))){usedF_I[team] += parseInt(range[j][12]); } }
          else if (range[j][0].toLowerCase() == 'c' && team != 5) { cpoCount[team] ++; if (!isNaN(parseInt(range[j][12]))){cpoF_I[team] += parseInt(range[j][12]); } }
          
        }
        else if (range[j + 1][0] == '' && range[j + 1][18] == '' && range[j + 2][0] == '' && range[j + 2][18] == '' && range[j + 3][0] == '' && range[j + 3][18] == '') { j = sheet.getLastRow(); }
      }
      //Logger.log(sheetName+'\n'+newCount+'\n'+newF_I+'\n'+cpoCount+'\n'+cpoF_I+'\n'+usedCount+'\n'+usedF_I);
    }
  }
  
  //Logger.log(newCount);Logger.log(newPvr);Logger.log(cpoCount);Logger.log(cpoPvr);Logger.log(usedCount);Logger.log(usedPvr);
  return [newCount, newF_I, '', cpoCount, cpoF_I, '', usedCount, usedF_I];
}
function snapshot22(x,y) {
  //Created By Kennen Lawrence
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  var sheetNames =  ['16th', '17th', '18th', '19th', '20th', '21st', '22nd'];
  
  for (var i = 0; i < sheetNames.length - 2; i++) {
    sheetName = sheetNames[i];
    sheet     = ss.getSheetByName(sheetName);
    
    if ( sheet != null) {
      range = sheet.getRange('C3:Z53').getValues();
      for (var j = 0; j < sheet.getLastRow() - 2; j++) {
        if (range[j][0] != '' && range[j][19] != '') {
          team = range[j][19];
          team = team.toString().replace('-', ' ');
          
          if (range[j][23].toLowerCase().indexOf('yes') !== -1) { newCount[newCount.length - 1] ++; }
          else if (range[j][23].toLowerCase() === 'no') { newF_I[newF_I.length - 1] ++; }
          
          if (team === 'Merrie Thomsen')        { team = 0; }
          else if (team === 'Ben Brahler')      { team = 1; }
          else if (team === 'Liz Liggett')      { team = 2; }
          else if (team === 'Joshua Buchanan')  { team = 3; }
          else if (team === 'Ace Taylor Brown') { team = 4; }
          else                                  { team = 5; }
          
          if (range[j][0].toLowerCase() == 'n' && team != 5) { newCount[team] ++; if (!isNaN(parseInt(range[j][12]))) { newF_I[team] += parseInt(range[j][12]); } }
          else if (range[j][0].toLowerCase() == 'u' && team != 5) { usedCount[team] ++; if (!isNaN(parseInt(range[j][12]))){usedF_I[team] += parseInt(range[j][12]); } }
          else if (range[j][0].toLowerCase() == 'c' && team != 5) { cpoCount[team] ++; if (!isNaN(parseInt(range[j][12]))){cpoF_I[team] += parseInt(range[j][12]); } }
          
        }
        else if (range[j + 1][0] == '' && range[j + 1][18] == '' && range[j + 2][0] == '' && range[j + 2][18] == '' && range[j + 3][0] == '' && range[j + 3][18] == '') { j = sheet.getLastRow(); }
      }
      //Logger.log(sheetName+'\n'+newCount+'\n'+newF_I+'\n'+cpoCount+'\n'+cpoF_I+'\n'+usedCount+'\n'+usedF_I);
    }
  }
  
  //Logger.log(newCount);Logger.log(newPvr);Logger.log(cpoCount);Logger.log(cpoPvr);Logger.log(usedCount);Logger.log(usedPvr);
  return [newCount, newF_I, '', cpoCount, cpoF_I, '', usedCount, usedF_I];
}
function snapshot15(x,y) {
  //Created By Kennen Lawrence
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  var sheetNames = ['8th', '9th', '10th', '11th', '12th', '13th', '14th', '15th'];
  
  for (var i = 0; i < sheetNames.length - 2; i++) {
    sheetName = sheetNames[i];
    sheet     = ss.getSheetByName(sheetName);
    
    if ( sheet != null) {
      range = sheet.getRange('C3:Z53').getValues();
      for (var j = 0; j < sheet.getLastRow() - 2; j++) {
        if (range[j][0] != '' && range[j][19] != '') {
          team = range[j][19];
          team = team.toString().replace('-', ' ');
          
          if (range[j][23].toLowerCase().indexOf('yes') !== -1) { newCount[newCount.length - 1] ++; }
          else if (range[j][23].toLowerCase() === 'no') { newF_I[newF_I.length - 1] ++; }
          
          if (team === 'Merrie Thomsen')        { team = 0; }
          else if (team === 'Ben Brahler')      { team = 1; }
          else if (team === 'Liz Liggett')      { team = 2; }
          else if (team === 'Joshua Buchanan')  { team = 3; }
          else if (team === 'Ace Taylor Brown') { team = 4; }
          else                                  { team = 5; }
          
          if (range[j][0].toLowerCase() == 'n' && team != 5) { newCount[team] ++; if (!isNaN(parseInt(range[j][12]))) { newF_I[team] += parseInt(range[j][12]); } }
          else if (range[j][0].toLowerCase() == 'u' && team != 5) { usedCount[team] ++; if (!isNaN(parseInt(range[j][12]))){usedF_I[team] += parseInt(range[j][12]); } }
          else if (range[j][0].toLowerCase() == 'c' && team != 5) { cpoCount[team] ++; if (!isNaN(parseInt(range[j][12]))){cpoF_I[team] += parseInt(range[j][12]); } }
          
        }
        else if (range[j + 1][0] == '' && range[j + 1][18] == '' && range[j + 2][0] == '' && range[j + 2][18] == '' && range[j + 3][0] == '' && range[j + 3][18] == '') { j = sheet.getLastRow(); }
      }
      //Logger.log(sheetName+'\n'+newCount+'\n'+newF_I+'\n'+cpoCount+'\n'+cpoF_I+'\n'+usedCount+'\n'+usedF_I);
    }
  }
  return [newCount, newF_I, '', cpoCount, cpoF_I, '', usedCount, usedF_I];
}
function snapshot7(x,y) {
  //Created By Kennen Lawrence
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  var sheetNames = ['1st', '2nd', '3rd', '4th', '5th', '6th', '7th'];
  
  for (var i = 0; i < sheetNames.length - 2; i++) {
    sheetName = sheetNames[i];
    sheet     = ss.getSheetByName(sheetName);
    
    if ( sheet != null) {
      range = sheet.getRange('C3:Z53').getValues();
      for (var j = 0; j < sheet.getLastRow() - 2; j++) {
        if (range[j][0] != '' && range[j][19] != '') {
          team = range[j][19];
          team = team.toString().replace('-', ' ');
          
          if (range[j][23].toLowerCase().indexOf('yes') !== -1) { newCount[newCount.length - 1] ++; }
          else if (range[j][23].toLowerCase() === 'no') { newF_I[newF_I.length - 1] ++; }
          
          if (team === 'Merrie Thomsen')        { team = 0; }
          else if (team === 'Ben Brahler')      { team = 1; }
          else if (team === 'Liz Liggett')      { team = 2; }
          else if (team === 'Joshua Buchanan')  { team = 3; }
          else if (team === 'Ace Taylor Brown') { team = 4; }
          else                                  { team = 5; }
          
          if (range[j][0].toLowerCase() == 'n' && team != 5) { newCount[team] ++; if (!isNaN(parseInt(range[j][12]))) { newF_I[team] += parseInt(range[j][12]); } }
          else if (range[j][0].toLowerCase() == 'u' && team != 5) { usedCount[team] ++; if (!isNaN(parseInt(range[j][12]))){usedF_I[team] += parseInt(range[j][12]); } }
          else if (range[j][0].toLowerCase() == 'c' && team != 5) { cpoCount[team] ++; if (!isNaN(parseInt(range[j][12]))){cpoF_I[team] += parseInt(range[j][12]); } }
          
        }
        else if (range[j + 1][0] == '' && range[j + 1][18] == '' && range[j + 2][0] == '' && range[j + 2][18] == '' && range[j + 3][0] == '' && range[j + 3][18] == '') { j = sheet.getLastRow(); }
      }
      //Logger.log(sheetName+'\n'+newCount+'\n'+newF_I+'\n'+cpoCount+'\n'+cpoF_I+'\n'+usedCount+'\n'+usedF_I);
    }
  }
//  Logger.log([newCount, newF_I, '', cpoCount, cpoF_I, '', usedCount, usedF_I]);
  return [newCount, newF_I, '', cpoCount, cpoF_I, '', usedCount, usedF_I];
}
//MGR=col 18; F_I=12; type=0
//[0] mike[1] jeff[2] chris[3] dean[4]