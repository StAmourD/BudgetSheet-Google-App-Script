function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  var dateToday = new Date();
  var thisMonth = dateToday.getMonth();
  var nextMonth = thisMonth + 1;
  var nextMonthYear = new Date(dateToday);
  nextMonthYear.setMonth(nextMonth);
  
  ui.createMenu('Budget Tools')
      .addItem('Import MS Money Data', 'showSidebar')
      .addItem('Create Copy: ' + months[nextMonth % 12] + ' ' + nextMonthYear.getFullYear(), 'copyToNextMonth')
      .addItem('Create Copy: ' + months[thisMonth % 12] + ' ' + dateToday.getFullYear(), 'copyToThisMonth')
      .addToUi();
}

function copyToNextMonth(){
  copyMonth(1);
}

function copyToThisMonth() {
  copyMonth(0);
}

function copyMonth(MonthsToAdd) {
  var months = ['January','February','March','April','May','June','July','August','September','October','November','December'];   
  var nextMonth = new Date().getMonth() + MonthsToAdd;
  var ui = SpreadsheetApp.getUi(); 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var newSheet;
  var nextMonthYear = new Date(); nextMonthYear.setMonth(nextMonth);
  var ThisRange;
  var y = nextMonthYear.getFullYear(), m = nextMonthYear.getMonth();
  var StartDate = new Date(y, m, 1);
  var EndDate = new Date(y, m + 1, 0);
  
  var result = ui.alert(
     'Please confirm',
     'Create a new sheet ' + months[nextMonth % 12] + '.' + nextMonthYear.getFullYear(),
     ui.ButtonSet.YES_NO);  
  if (result == ui.Button.YES) {
      try {
        ss.setActiveSheet(ss.getSheetByName(months[nextMonth % 12] + '.' + nextMonthYear.getFullYear()));
        ss.toast('Sheet already exists.');
      }
      catch (e) {
        // sheet doesnt already exist OK to copy
        newSheet = sheet.copyTo(ss);
        newSheet.setName(months[nextMonth % 12] + '.' + nextMonthYear.getFullYear())
        ss.setActiveSheet(newSheet);
        ss.moveActiveSheet(sheet.getIndex());
        sheet.setTabColor('BROWN');
        newSheet.setTabColor('ORANGE');
        // clear old actuals
        ThisRange = newSheet.getRange("IncomeActuals");
        ThisRange.setValue('0');
        ThisRange = newSheet.getRange("ExpenseActuals");
        ThisRange.setValue('0');
        // set date range
        ThisRange = newSheet.getRange("DateRange");
        ThisRange.clearContent();
        ThisRange.setValue(StartDate.toDateString() + ' - ' + EndDate.toDateString());
      }
  }
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Import MS Money Data')
      .setWidth(300);
  SpreadsheetApp.getUi()
      .showSidebar(html);
}

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
  return RetData;
}

function sortExpenses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var expenseRange = sheet.getRange('expenseAllColumns');
  
  expenseRange.sort(1);
}

function InsertNewExpense(catName, amount) {
  catName = catName || 'testName1';
  amount = Number(amount) || 12;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var columnCount = sheet.getRange('expenseAllColumns').getWidth() + 1;
  var rangeToCopy = sheet.getRange(10, 1, 1, columnCount);
  var rangeDestination = sheet.getRange(11, 1, 1, columnCount);
  
  sheet.insertRowAfter(10);
  rangeToCopy.copyTo(rangeDestination);
  sheet.getRange(11, 1).setValue(catName);
  sheet.getRange(11, 2).setValue(amount);
  sheet.getRange(11, 3).setValue(0);
  sortExpenses()
  return;
}
