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

function createNewMonthSheet() {
  let currentDate = new Date();
  SpreadsheetApp.getActiveSpreadsheet().insertSheet(currentDate.toLocaleDateString('en-US', { year: 'numeric', month: 'long' }));

  // wait for sheet to be created
  SpreadsheetApp.flush();

  let currentMonthSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(currentDate.toLocaleDateString('en-US', { year: 'numeric', month: 'long' }));

  let timeColumnWidth = 60;
  let textColumnWidth = 70;
  let iconColumnWidth = 23;

  // set column width
  currentMonthSheet.setColumnWidth(1, textColumnWidth);
  currentMonthSheet.setColumnWidth(2, timeColumnWidth);
  currentMonthSheet.setColumnWidth(3, iconColumnWidth);
  currentMonthSheet.setColumnWidth(4, textColumnWidth);
  currentMonthSheet.setColumnWidth(5, timeColumnWidth);

  currentMonthSheet.setColumnWidth(1 + 6, textColumnWidth);
  currentMonthSheet.setColumnWidth(2 + 6, timeColumnWidth);
  currentMonthSheet.setColumnWidth(3 + 6, iconColumnWidth);
  currentMonthSheet.setColumnWidth(4 + 6, textColumnWidth);
  currentMonthSheet.setColumnWidth(5 + 6, timeColumnWidth);

  currentMonthSheet.getRange("H:H").setNumberFormat('HH:mm:ss');
  currentMonthSheet.getRange("K:K").setNumberFormat('HH:mm:ss');
  // currentMonthSheet.setColumnWidth(7, 132); // break time config column
  // currentMonthSheet.setColumnWidth(10, 132); // work time config column

  // add fields with configurable break time and work time
  currentMonthSheet.getRange("A1").setValue("Config").setFontWeight('bold');
  currentMonthSheet.getRange("A1:D1").mergeAcross().setHorizontalAlignment('center');

  currentMonthSheet.getRange("A2").setValue("Desired break time:").setFontWeight('bold').setBackground("#ffd966");
  currentMonthSheet.getRange("C2").setValue("0:30").setBackground("#ffe599").setFontStyle("italic");
  currentMonthSheet.getRange("A2:B2").mergeAcross();
  currentMonthSheet.getRange("C2:D2").mergeAcross();

  currentMonthSheet.getRange("A3").setValue("Desired work time:").setFontWeight('bold').setBackground("#ffd966");
  currentMonthSheet.getRange("C3").setValue("8:00").setBackground("#ffe599").setFontStyle("italic");
  currentMonthSheet.getRange("A3:B3").mergeAcross();
  currentMonthSheet.getRange("C3:D3").mergeAcross();

  currentMonthSheet.appendRow([" "]);

  return currentMonthSheet;
}

function workStart(currentMonthSheet) {

  Logger.log("Work start!");
  let currentDate = new Date();

  // append and stylize header for specific day
  currentMonthSheet.appendRow([currentDate.toLocaleDateString('en-US', { weekday: 'long', day: 'numeric', month: 'long' })]);
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 1, 1, 11).mergeAcross().setFontWeight('bold').setHorizontalAlignment('center').setBorder(true, true, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  // append actual log
  let desiredBreakTime = '"' + currentMonthSheet.getRange("C2").getDisplayValue() + '"';
  let desiredWorkTime = '"' + currentMonthSheet.getRange("C3").getDisplayValue() + '"';
  let leaveAtFormula = "=INDIRECT(ADDRESS(ROW();COLUMN()-3))+" + desiredBreakTime + "+" + desiredWorkTime;
  currentMonthSheet.appendRow(["Started", currentDate.toLocaleTimeString(), iconStart, "Leave at", leaveAtFormula]);
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 1, 1, rowWidth).setBackground(colorStart).setHorizontalAlignment('center');

  // append log with rounded data
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 7, 1, rowWidth).setValues([["Started", '=MROUND(INDIRECT(ADDRESS(ROW();COLUMN()-6));"00:10:00")', iconStart, "Leave at", '=MROUND(INDIRECT(ADDRESS(ROW();COLUMN()-6));"00:10:00")']]);
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 7, 1, rowWidth).setBackground(colorStart).setHorizontalAlignment('center');




  notify("Started at " + currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 2).getDisplayValue() 
  + " (" +  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 2+6).getDisplayValue() + ")\n" + 
  "Leave at " + currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 5).getDisplayValue() 
  + " (" +  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 5+6).getDisplayValue() + ")"

  );
}


function notify(notification) {
  SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange("F1").setValue(notification);
  // Logger.log(SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange("F1").getValue());




}

function workEnd(currentMonthSheet) {
  Logger.log("Work end!");
  let currentDate = new Date();
  
  
  // append ending date
  currentMonthSheet.appendRow(["Stopped", currentDate.toLocaleTimeString(), iconEnd]);
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 1, 1, rowWidth).setBackground(colorEnd).setHorizontalAlignment('center');

  let notifyString = "Stopped at " + currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 2).getDisplayValue();

  // append rounded ending date
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 7, 1, rowWidth).setValues([["Stopped", '=MROUND(INDIRECT(ADDRESS(ROW();COLUMN()-6));"00:10:00")', iconEnd, "", ""]]);
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 7, 1, rowWidth).setBackground(colorEnd).setHorizontalAlignment('center');
  notifyString += " (" +  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 2+6).getDisplayValue() + ")\n";

  // append time worked 
  const timeWorkedFormula = "=INDIRECT(ADDRESS(ROW()-1;COLUMN()))-INDIRECT(ADDRESS(ROW()-2;COLUMN()))";
  currentMonthSheet.appendRow(["Worked", timeWorkedFormula, summaryIcon]);
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 1, 1, rowWidth).setBackground(summaryColor).setHorizontalAlignment('center');
  notifyString += "Worked for " + currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 2).getDisplayValue();

  // append rounded time worked
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 7, 1, rowWidth).setValues([["Worked", '=MROUND(INDIRECT(ADDRESS(ROW();COLUMN()-6));"00:10:00")', summaryIcon, "", ""]]);
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 7, 1, rowWidth).setBackground(summaryColor).setHorizontalAlignment('center');
  notifyString += " (" +  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 2+6).getDisplayValue() + ")\n";

  // append free space before new log 
  currentMonthSheet.appendRow([" "]);

  notify(notifyString);
}

function addNotificationContents(start) {

  if (start) {
    // work start notification

  } else {
    // work end notification
  }

}

function isItStart(rowNumber, currentMonthSheet) {
  // check if it's start or end by checking icon in previous row

  if (currentMonthSheet.getRange(rowNumber, colIcon, 1, 1).getValue() == iconStart) {
    return false;
  }
  else {
    return true;
  }
}
