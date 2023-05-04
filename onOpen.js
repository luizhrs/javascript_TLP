function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Forecast')
      .addItem('Generate Forecast', 'generateForcast')
      .addItem('Clear Forecast', 'clearForecast')
      .addItem('Adjust Forecast Table', 'adjustMonth')
      .addToUi();
}

function generateForecast() {
  var mainSpreadsheet = SpreadsheetApp.openById("google_spreadsheet_ID");
  var summarySheet = mainSpreadsheet.getSheetByName("Summary");
  var expensesSheet = mainSpreadsheet.getSheetByName("Expenses import");
  var incomesSheet = mainSpreadsheet.getSheetByName("Incomes import");
  var forecastSheet = mainSpreadsheet.getSheetByName("Forecasting");
  var oneOffExpensesSheet = mainSpreadsheet.getSheetByName("One-off-expenses");
  var oneOffIncomesSheet = mainSpreadsheet.getSheetByName("One-off-incomes");

  // Import one-off expenses and incomes
  importRange(oneOffExpensesSheet.getRange("A:A"), expensesSheet.getRange(checkLastRow(expensesSheet, "A:A")+1, 2), 1);
  importRange(oneOffIncomesSheet.getRange("A:A"), incomesSheet.getRange(checkLastRow(incomesSheet, "A:A")+1, 2), 1);

  // Create forecast table
  var lastRow = checkLastRow(expensesSheet, "B:B");
  var lastUpdate = expensesSheet.getRange(lastRow, 4).getValue();
  var monthsLeft = monthDiff(lastUpdate, summarySheet.getRange("BUDGETYEAR").getValue());
  createForecastTable(expensesSheet, lastRow, forecastSheet, lastUpdate, monthsLeft);

  SpreadsheetApp.flush();
}

function clearForecast() {
  var mainSpreadsheet = SpreadsheetApp.openById("google_spreadsheet_ID");
  var expensesSheet = mainSpreadsheet.getSheetByName("Expenses import");
  var incomesSheet = mainSpreadsheet.getSheetByName("Incomes import");

  // Clear one-off expenses and incomes
  clearContentAfterRow(expensesSheet, checkLastRow(expensesSheet, "A:A")+1, 2, expensesSheet.getLastColumn()-2);
  clearContentAfterRow(incomesSheet, checkLastRow(incomesSheet, "A:A")+1, 2, incomesSheet.getLastColumn()-2);
}

function adjustMonth() {
  var mainSpreadsheet = SpreadsheetApp.openById("google_spreadsheet_ID");
  var forecastSheet = mainSpreadsheet.getSheetByName("Forecasting");

  var currentMonth = new Date().getMonth();
  var forecastMonth = forecastSheet.getRange(1, 3).getValue().getMonth();
  if (forecastMonth > currentMonth && currentMonth != 0) return;

  var monthlyTable = getRangeToPaste(forecastSheet,"C:C").getValues();
  monthlyTable.forEach((row, index) => {
    row[2].setMonth(row[2].getMonth() + 1);
    forecastSheet.getRange(index+1,3).setValue(row[2]);
  });
}
