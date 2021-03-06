function importData(ThisData) {
  var FirstExpenseRow = 10000;
  var FirstIncomeRow = 10000;
  ThisData.forEach(function(item, index, array) {
    if (item[0] == 'Income') {
      FirstIncomeRow = index + 1;
    }
    if (item[0] == 'Expenses') {
      FirstExpenseRow = index + 1;
    }
  });
  
  var ia, iy;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var cell;
  var LastRow = sheet.getLastRow();
  for(ia = 1; ia <= LastRow; ia++) {
    cell = sheet.getRange(ia, 1);
//    move Income data
    for(iy = FirstIncomeRow; iy < FirstExpenseRow - 2; iy++) {
      if(ThisData[iy][0] === cell.getValue() && ThisData[iy][0].trim().length > 0) {
        cell = sheet.getRange(ia, 2);
        cell.setValue(ThisData[iy][1]);
        ThisData[iy][3] = '1';
      }
    }
//    Move Expense data
    for(iy = FirstExpenseRow; iy < ThisData.length; iy++) {
      if(ThisData[iy][0] === cell.getValue() && ThisData[iy][0].trim().length > 0) {
        cell = sheet.getRange(ia, 2);
        cell.setValue(ThisData[iy][1]);
        ThisData[iy][3] = '1';
      }
    }
  }
  var RetData = new Array();
  ThisData.forEach(function(item, index, array) {
    if (item[3] !== '1' && item.length === 3 && item[1] !== '' && item[0] !== 'Category' && item[0] !== 'Total Income' && item[0] !== 'Total Expenses' && item[0] !== 'Income less Expenses') {
      RetData.push(item);
    }
  });
  
  //update last updated date:
  var curDat = new Date();
  sheet.getRange(1, 4).setValue(curDat);
  if(RetData.length > 0) {
    sheet.getRange(1, 8).setValue('Missing ' + RetData.length + ' categories.');
  } else {
    sheet.getRange(1, 8).setValue('');
  }
  return RetData;
}