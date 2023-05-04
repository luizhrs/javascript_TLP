function createForcastTable(expensesSheet, lastRow, forecastSheet, lastUpdate, monthsLeft) {
  const monthlyTable = getRangeToPaste(forecastSheet, "A:A").getValues();

  for (let i = 1; i <= monthsLeft; i++) {
    const forecastTable = monthlyTable.filter(row => row[2] > lastUpdate);

    if (forecastTable.length > 0) {
      importRange(forecastTable, expensesSheet.getRange(lastRow + 1, 2), 2);
    }

    lastRow = checkLastRow(expensesSheet, "B:B");
    monthlyTable.forEach(row => row[2].setMonth(row[2].getMonth() + 1));
  }
}

function getLastRowSpecial(range) {
  for (var row = range.length - 1; row >= 0; row--) {
    if (range[row][0] !== "") {
      return row + 1;
    }
  }
  return 0;
}

function importRange(sourceRng, destStartRange, type) {
  var sourceVals = type === 1 ? sourceRng.getValues() : sourceRng;
  var destSheet = destStartRange.getSheet();
  var numRows = sourceVals.length;
  var numCols = sourceVals[0].length;
  var destRange = destSheet.getRange(destStartRange.getRow(), destStartRange.getColumn(), numRows, numCols);
  destRange.setValues(sourceVals);
  SpreadsheetApp.flush();
}

function checkLastRow(oSheet, sColumn) {
  var columnValues = oSheet.getRange(sColumn).getValues();
  return getLastRowSpecial(columnValues);
}

function getRangeToPaste(sSheet, sColumn) {
  var numRows = checkLastRow(sSheet, sColumn);
  return sSheet.getRange(1, 1, numRows, sSheet.getLastColumn());
}
