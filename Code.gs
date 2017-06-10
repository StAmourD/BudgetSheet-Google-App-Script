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
      .addItem('Create Copy for Next Month: ' + months[nextMonth % 12] + ' ' + nextMonthYear.getFullYear(), 'copyToNextMonth')
      .addItem('Create Copy for This Month: ' + months[thisMonth % 12] + ' ' + dateToday.getFullYear(), 'copyToThisMonth')
      .addToUi();
}

function copyToNextMonth() {
  var ui = SpreadsheetApp.getUi(); 
  var nextMonth = new Date().getMonth() + 1;
  var months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  var result = ui.alert(
     'Please confirm',
     'Create a new sheet for ' + months[nextMonth % 12],
     ui.ButtonSet.YES_NO);
     
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var newSheet;
  var nextMonthYear = new Date();
  nextMonthYear.setMonth(nextMonth);
  var myRanges;
  
  if (result == ui.Button.YES) {
      try {
        ss.setActiveSheet(ss.getSheetByName(((nextMonth % 12) + 1) + '.' + nextMonthYear.getFullYear()));
        ss.toast('Sheet already exists.');
      }
      catch (e) {
        // sheet doesnt already exist OK to copy
        newSheet = sheet.copyTo(ss);
        newSheet.setName(((nextMonth % 12) + 1) + '.' + nextMonthYear.getFullYear())
        ss.setActiveSheet(newSheet);
        ss.moveActiveSheet(sheet.getIndex());
        sheet.setTabColor('BROWN');
        newSheet.setTabColor('ORANGE');
        myRanges = ss.getNamedRanges();
//        myRanges['7.2017\'!IncomeActuals'].getRange().clearContent();
        Logger.log(myRanges[1]);
      }
  }
}


function copyToThisMonth() {
  var ui = SpreadsheetApp.getUi(); 
  var result = ui.alert('Not implemented!');
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
