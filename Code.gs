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

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Import MS Money Data')
      .setWidth(300);
  SpreadsheetApp.getUi()
      .showSidebar(html);
}

function addExpenseRow() {
  var html = HtmlService.createHtmlOutputFromFile('AddExpense')
    .setWidth(300)
    .setHeight(120);
  SpreadsheetApp.getUi()
    .showModalDialog(html, 'Add New Expense');
}

function InsertNewExpense(catName, amount, toReturn) {
  catName = catName || 'testName1';
  amount = Number(amount) || 12;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var columnCount = sheet.getRange('expenseAllColumns').getWidth() + 1;
  var firstRow = sheet.getRange('expenseAllColumns').getRow();
  var rangeToCopy = sheet.getRange(firstRow, 1, 1, columnCount);
  var rangeDestination = sheet.getRange(firstRow + 1, 1, 1, columnCount);
  
  sheet.insertRowAfter(firstRow);
  rangeToCopy.copyTo(rangeDestination);
  sheet.getRange(firstRow + 1, 1).setValue(catName);
  sheet.getRange(firstRow + 1, 2).setValue(amount);
  sheet.getRange(firstRow + 1, 3).setValue(0);
  sortExpenses();
  return toReturn;
}

function sortExpenses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var expenseRange = sheet.getRange('expenseAllColumns');
  
  expenseRange.sort(1);
}