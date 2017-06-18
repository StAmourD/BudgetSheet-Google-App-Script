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