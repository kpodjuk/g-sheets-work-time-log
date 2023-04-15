// PARAMS:
const rowWidth = 5; // how many columns to colorize when new row is added

// columns width
const statusColWidth = 110;
const dateColWidth = 132
const iconColWidth = 25

// column order
const colStatus = 1;
const colIcon = 3;
const colDate = 2;
const colBreak = 4;

// ⏳🏁🚩🆕🆓⏩⏮️🔼◀️▶️⬅️⬆️⬇️↖️↔️↕️🔝🔛☑️🔚🔙
// start working
const colorStart = 'ACCENT1';
const iconStart = "🚩";

// end working
const colorEnd = 'ACCENT2';
const iconEnd = "🏁";

// summary
const summaryColor = 'orange';
// const summaryColor = '#4285f4';
const summaryIcon = '🕐';

const debugDate = new Date(2023, 4, 16, 8, 0, 0);

// how long for usual day of work
const usualDayAtWorkHours = 8;
const usualDayAtWorkMinutes = 30;
const usualDayAtWorkMs = (usualDayAtWorkHours*60*60*1000)+usualDayAtWorkMinutes*60*1000;



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
      processedSheet.setColumnWidth(colStatus, 70);

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

  Logger.log("usualDayAtWorkMs: "+usualDayAtWorkMs);



  // append and stylize "header"
  processedSheet.appendRow([currentDate.toLocaleDateString('en-US', { weekday: 'long', day: 'numeric', month: 'long' }), "Break: "]);

  processedSheet.getRange(rowNumber + 1, 1, 1, rowWidth).mergeAcross();

  processedSheet.getRange(rowNumber + 1, 1, 1, rowWidth).setFontWeight('bold').setHorizontalAlignment('center');

  // append actual log
  processedSheet.appendRow(["Started: ", currentDate.toLocaleTimeString(), iconStart, "Leave at: ", new Date(currentDate.getTime() + usualDayAtWorkMs).toLocaleTimeString()]);

  processedSheet.getRange(processedSheet.getLastRow(), 1, 1, rowWidth).setBackground(colorStart);


}

function workEnd(processedSheet) {

  Logger.log("Work end!");
  let currentDate = new Date();



  processedSheet.appendRow(["Stopped", currentDate.toLocaleTimeString(), iconEnd]);

  processedSheet.getRange(processedSheet.getLastRow(), 1, 1, rowWidth).setBackground(colorEnd);

  // it's end, summarize
  // get start/end date from previous rows
  let startDate = new Date(processedSheet.getRange(processedSheet.getLastRow() - 1, colDate).getValue());
  let endDate = new Date(processedSheet.getRange(processedSheet.getLastRow(), colDate).getValue());

  let elapsed = ((endDate.getTime() - startDate.getTime()) / 1000);
  // processedSheet.appendRow(["Worked ", elapsed + "s", summaryIcon]);

  const timeWorkedFormula = "=INDIRECT(ADDRESS(ROW()-1;COLUMN()))-INDIRECT(ADDRESS(ROW()-2;COLUMN()))"; 
  processedSheet.appendRow(["Worked ", timeWorkedFormula, summaryIcon]);
  

  processedSheet.getRange(processedSheet.getLastRow(), 1, 1, rowWidth).setBackground(summaryColor);

  processedSheet.appendRow([" "]);

}


function getMonthName(monthNumber) {
  switch (monthNumber) {
    case 0:
      return "January"
  }
}

