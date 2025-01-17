function getCalendarWeek(currentDate) {
  // Create a copy of the date to avoid modifying the original
  let date = new Date(currentDate);

  // set to DEBUG date
  // date = new Date('2025-01-19');

  // Set the date to the nearest Thursday
  date.setUTCDate(date.getUTCDate() + 4 - (date.getUTCDay() || 7));

  // Get the first day of the year
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));

  // Calculate the difference in milliseconds and convert to weeks
  const weekNumber = Math.ceil(((date - yearStart) / 86400000) / 7);

  return String(weekNumber);
}

function isItStart(rowNumber, currentMonthSheet) {
  // check if it's start or end by checking icon in previous row
  if (currentMonthSheet.getRange(rowNumber, 3, 1, 1).getValue() == iconStart) {
    return false;
  }
  else {
    return true;
  }
}

function onSheetUpdate(e) {
  // spreadsheet - file itself
  // sheet - one of many sheets in spreadsheet

  let updatedSheet = SpreadsheetApp.getActiveSheet();
  if (updatedSheet.getIndex() == "1.0" && e.changeType == "INSERT_ROW") {

    // it's new log entry!
    let currentDate = new Date();
    let currentMonthSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(currentDate.toLocaleDateString('en-US', { year: 'numeric', month: 'long' }));

    if (currentMonthSheet != null) {
      // sheet exists, append
      // calculate row
      let rowNumber = currentMonthSheet.getDataRange().getValues().length;

      if (isItStart(rowNumber, currentMonthSheet)) {
        workStart(currentMonthSheet, rowNumber);
      } else {
        workEnd(currentMonthSheet);
      }
    } else {
      // sheet doesn't exist, create and then append
      workStart(createNewMonthSheet()); // add first log to 4th row 
    }
  } else {
    // ignore
    Logger.log("Change ignored!");
  }
}

function notify(notification) {
  SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange("F1").setValue(notification);
}
