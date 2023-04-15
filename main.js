// PARAMS:

const rowWidth = 3; // how many columns to colorize when new row is added
// const colorStart = '#ffff00';

const colorStart = 'yellow';
const iconStart = "🚩";

const colorEnd = 'orange';
const iconEnd = "🏁";

const summaryColor = '#4285f4';

const debugDate = new Date(2023, 1, 1, 1, 1, 1);

const colIcon = 3;
const colDate = 2;

// ⏳🏁🚩🆕🆓⏩⏮️🔼◀️▶️⬅️⬆️⬇️↖️↔️↕️🔝🔛☑️🔚🔙

function isItStart(rowNumber, processedSheet) {
  // check if it's start or end by checking icon in previous row

  if (processedSheet.getRange(rowNumber, colIcon, 1, 1).getValue() == iconStart) {
    return false;
  }
  else {
    return true;
  }
}

function onSheetUpdate(e) {
  let logSheet = SpreadsheetApp.getActiveSheet();

  if (logSheet.getIndex() == "1.0" && e.changeType == "INSERT_ROW") {
    // it's new log entry!

    // let rowNumber = logSheet.getDataRange().getValues().length;

    let currentDate = new Date();
    let processedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(currentDate.toLocaleDateString('en-US', { year: 'numeric', month: 'long' }));



    if (processedSheet != null) {
      // sheet exists, append

      // calculate row
      let rowNumber = processedSheet.getDataRange().getValues().length;

      if (isItStart(rowNumber, processedSheet)) {
        workStart(processedSheet, rowNumber);
      } else {
        workEnd(processedSheet);
      }
    } else {
      // sheet doesn't exist, create and then append
      let currentDate = new Date();
      SpreadsheetApp.getActiveSpreadsheet().insertSheet(currentDate.toLocaleDateString('en-US', { year: 'numeric', month: 'long' }));

      // wait for sheet to be created
      SpreadsheetApp.flush();

      let processedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(currentDate.toLocaleDateString('en-US', { year: 'numeric', month: 'long' }));

      // set column width
      processedSheet.setColumnWidth(colDate, 132);
      processedSheet.setColumnWidth(colIcon, 25);

      workStart(processedSheet, 0);
    }
  } else {
    // ignore
    Logger.log("Change ignored!");
  }
}


function workStart(processedSheet, rowNumber) {
  Logger.log("Work start!");
  let currentDate = new Date();

  // append and stylize "header"
  processedSheet.appendRow([currentDate.toLocaleDateString('en-US', { weekday: 'long', day: 'numeric', month: 'long' })]);

  processedSheet.getRange(rowNumber + 1, 1, 1, rowWidth).mergeAcross();

  processedSheet.getRange(rowNumber + 1, 1, 1, rowWidth).setFontWeight('bold').setHorizontalAlignment('center');

  // append actual log
  processedSheet.appendRow(["You started working", currentDate, iconStart]);

  processedSheet.getRange(processedSheet.getLastRow(), 1, 1, rowWidth).setBackground(colorStart);


}

function workEnd(processedSheet) {

  Logger.log("Work end!");
  let currentDate = new Date();



  processedSheet.appendRow(["You stop working", currentDate, iconEnd]);

  processedSheet.getRange(processedSheet.getLastRow(), 1, 1, rowWidth).setBackground(colorEnd);

  // it's end, summarize
  // date column right now: 2
  // get start/end date from previous rows
  let startDate = new Date(processedSheet.getRange(processedSheet.getLastRow() - 1, colDate).getValue());
  let endDate = new Date(processedSheet.getRange(processedSheet.getLastRow(), colDate).getValue());

  let elapsed = ((endDate.getTime() - startDate.getTime()) / 1000);

  processedSheet.appendRow(["Work over!: ", elapsed + "s"]);
  processedSheet.getRange(processedSheet.getLastRow(), 1, 1, rowWidth).setBackground(summaryColor);

  processedSheet.appendRow([" "]);

}


function getMonthName(monthNumber) {
  switch (monthNumber) {
    case 0:
      return "January"
  }
}

