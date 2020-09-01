const driver = {
  columns: {
    vehicleType: 4,
    genius: 25,
    manager: 20,
    fandiGross: 12,
    rebate: 19
  },
  rows: {
    first: 3,
    last: 52
  },
  teams: [
    'Ben Wegener',
    'Ben Brahler',
    'Jeff Englert',
    'Ace TB',
    'MER'
  ],
};

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

function snapshotCore (sheetNames) {
  //Created By Kennen Lawrence
  const newCount  = [0, 0, 0, 0, 0, 0];
  const newF_I    = [0, 0, 0, 0, 0, 0];
  const cpoCount  = [0, 0, 0, 0, 0];
  const cpoF_I    = [0, 0, 0, 0, 0];
  const usedCount = [0, 0, 0, 0, 0];
  const usedF_I   = [0, 0, 0, 0, 0];
  const newPvr    = [];
  const usedPvr   = [];
  const cpoPvr    = [];
  
  sheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) return;
    
    const rangeValues = sheet.getRange(
      driver.rows.first,
      driver.columns.vehicleType,
      driver.rows.last - driver.rows.first + 1,
      driver.columns.genius + 1
    ).getValues();

    rangeValues.forEach(rowValues => {
      if (rowValues[0] === '' && rowValues[driver.columns.manager] === '') return;
      
      let value = rowValues[driver.columns.genius].toString().toLowerCase();
  
      if (value.indexOf('yes') !== -1) newCount[newCount.length - 1] ++;
      else if (value === 'no') newF_I[newF_I.length - 1] ++;

      const team = driver.teams.indexOf(rowValues[driver.columns.manager].toString().replace('-', ' '));
  
      if (team === -1) return;

      value = parseInt(rowValues[driver.columns.fandiGross]) || 0;
  
      switch(rowValues[0].toString().toLowerCase()) {
        case 'n':
          newCount[team] ++;
          newF_I[team] += value;
          break;
        case 'u':
          usedCount[team] ++;
          usedF_I[team] += value;
          break;
        case 'c':
          cpoCount[team] ++;
          cpoF_I[team] += value;
          break;
        default:
          return;
      }
    });
  });
  
  return [newCount, newF_I, '', cpoCount, cpoF_I, '', usedCount, usedF_I];
}
