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

  addRaportButton(currentMonthSheet);

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
  currentMonthSheet.setColumnWidth(6 + 6, timeColumnWidth);
  currentMonthSheet.setColumnWidth(7 + 6, 1);
  currentMonthSheet.setColumnWidth(8 + 6, timeColumnWidth);
  currentMonthSheet.setColumnWidth(9 + 6, timeColumnWidth);

  currentMonthSheet.getRange("H:H").setNumberFormat('HH:mm:ss');
  currentMonthSheet.getRange("K:K").setNumberFormat('HH:mm:ss');
  // currentMonthSheet.setColumnWidth(7, 132); // break time config column
  // currentMonthSheet.setColumnWidth(10, 132); // work time config column

  // add fields with configurable break time and work time
  currentMonthSheet.getRange("A1").setValue("Config").setFontWeight('bold');
  currentMonthSheet.getRange("A1:D1").mergeAcross().setHorizontalAlignment('center');

  currentMonthSheet.getRange("A2").setValue("Default break time:").setFontWeight('bold').setBackground("#ffd966");
  currentMonthSheet.getRange("C2").setValue("0:30").setBackground("#ffe599").setFontStyle("italic");
  currentMonthSheet.getRange("A2:B2").mergeAcross();
  currentMonthSheet.getRange("C2:D2").mergeAcross();

  currentMonthSheet.getRange("A3").setValue("Default work time:").setFontWeight('bold').setBackground("#ffd966");
  currentMonthSheet.getRange("C3").setValue("8:00").setBackground("#ffe599").setFontStyle("italic");
  currentMonthSheet.getRange("A3:B3").mergeAcross();
  currentMonthSheet.getRange("C3:D3").mergeAcross();

  currentMonthSheet.getRange("N1:N2").mergeVertically().setValue("Break").setFontWeight('bold').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setHorizontalAlignment('center');
  currentMonthSheet.getRange("O1:O2").mergeVertically().setValue("Work").setFontWeight('bold').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setHorizontalAlignment('center');


  currentMonthSheet.getRange("N:N").setBackground("grey");
  currentMonthSheet.getRange("O:O").setBackground("grey");

  currentMonthSheet.appendRow([" "]);
  currentMonthSheet.appendRow([" "]);
  currentMonthSheet.appendRow([" "]);
  currentMonthSheet.appendRow([" "]);

  addBalanceStat(currentMonthSheet);

  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveSheet().setColumnWidth(6, 32);

  return currentMonthSheet;
}



//Function to insert image
function addRaportButton(sheet) {
  var image = sheet.insertImage("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQg73FxMo2IUhWG4n28zAtBEprZuVn51qlhntW_qlFBln0OjnjhrRE1_OADbFV7YtDmxts&usqp=CAU", 1, 4); //change the URL to the image you prefer

  image.assignScript("generateReport"); //assign the function to the image
  image.setAnchorCell(sheet.getRange('P1')).setHeight(95).setWidth(95);

}

function generateReport() {
  Logger.log("GENERATE REPORT!!!");
}

function workStart(currentMonthSheet) {

  Logger.log("Work start!");
  let currentDate = new Date();

  // append and stylize header for specific day
  currentMonthSheet.appendRow([currentDate.toLocaleDateString('en-US', { weekday: 'long', day: 'numeric', month: 'long' })]);
  /// add color and border
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 1, 1, 11).mergeAcross().setFontWeight('bold').setHorizontalAlignment('center').setBorder(true, true, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK).setBackground("gray");
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 1, 1, 15).setBorder(true, true, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.DOUBLE);


  // append actual log
  currentMonthSheet.appendRow(["Started", currentDate.toLocaleTimeString(), iconStart, "Leave at", "=INDIRECT(ADDRESS(ROW();COLUMN()-3))+INDIRECT(ADDRESS(ROW();COLUMN()+9))+INDIRECT(ADDRESS(ROW();COLUMN()+10))"]);

  // add border marking editable fields
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 2, 2, 1).setBorder(true, true, true, true, null, null, '#00ffff', SpreadsheetApp.BorderStyle.DASHED);
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 14, 1, 2).setBorder(true, true, true, true, null, null, '#00ffff', SpreadsheetApp.BorderStyle.DASHED);

  // set background for 5 leftmost cells
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 1, 1, 5).setBackground(colorStart).setHorizontalAlignment('center');



  let thisRowWidth = 9;

  // currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 1, 1, thisRowWidth-3).setBackground(colorStart).setHorizontalAlignment('center');

  // append log with rounded data
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 7, 1, thisRowWidth).setHorizontalAlignment('center').setValues([["Started", '=MROUND(INDIRECT(ADDRESS(ROW();COLUMN()-6));"00:10:00")', iconStart, "Leave at", '=MROUND(INDIRECT(ADDRESS(ROW();COLUMN()-6));"00:10:00")',
    "", "", // gray column with data points
    currentMonthSheet.getRange("C2").getDisplayValue(),
    currentMonthSheet.getRange("C3").getDisplayValue()
  ]]);
  // currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 7, 1, thisRowWidth-4).setBackground(colorStart).setHorizontalAlignment('center');
  // set background for cells after divide 
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 7, 1, 5).setBackground(colorStart).setHorizontalAlignment('center')



  notify("Started at:\t\t\t\t" +
    currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 2 + 6).getDisplayValue() +
    " (" + currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 2).getDisplayValue() + ")\n" +
    "Leave at:\t\t\t\t" +
    currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 5 + 6).getDisplayValue() +
    " (" + currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 5).getDisplayValue() + ")"

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


  // append rounded ending date
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 7, 1, rowWidth).setValues([["Stopped", '=MROUND(INDIRECT(ADDRESS(ROW();COLUMN()-6));"00:10:00")', iconEnd, "", ""]]);
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 7, 1, rowWidth).setBackground(colorEnd).setHorizontalAlignment('center');
  
  // let notifyString = "Stopped at " + currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 2).getDisplayValue();
  let notifyString = "Stopped at:\t\t\t\t" + currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 2 + 6).getDisplayValue();
  notifyString += " (" + currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 2).getDisplayValue() + ")\n";

  // append time spent and worked 
  currentMonthSheet.appendRow(["Time spent", '=INDIRECT(ADDRESS(ROW()-1;COLUMN()))-INDIRECT(ADDRESS(ROW()-2;COLUMN()))', summaryIcon, "Worked", '=IF(INDIRECT(ADDRESS(ROW();COLUMN()-3))-INDIRECT(ADDRESS(ROW()-2;COLUMN()+9)) <= 0;"00:00:00";INDIRECT(ADDRESS(ROW();COLUMN()-3))-INDIRECT(ADDRESS(ROW()-2;COLUMN()+9)))']);
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 1, 1, rowWidth).setBackground(summaryColor).setHorizontalAlignment('center');

  // append rounded time worked
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 7, 1, rowWidth + 9).setValues([["Time spent", '=MROUND(INDIRECT(ADDRESS(ROW();COLUMN()-6));"00:10:00")', summaryIcon, "Worked", '=MROUND(INDIRECT(ADDRESS(ROW();COLUMN()-6));"00:10:00")', "=(INDIRECT(ADDRESS(ROW();COLUMN()-4))-INDIRECT(ADDRESS(ROW()-2;COLUMN()+3)))-INDIRECT(ADDRESS(ROW()-2;COLUMN()+2))"
    // "secret" time worked for easy sum of hours
    , "", "", "", "", "", "", "", "=INDIRECT(ADDRESS(ROW();COLUMN()-9))"]]);
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 7, 1, rowWidth + 1).setBackground(summaryColor).setHorizontalAlignment('center');
  // add line at the end of day
  currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 1, 1, 15).setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.DOUBLE);

  notifyString += "Worked for:\t\t\t\t" + currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 5 + 6).getDisplayValue();
  notifyString += " (" + currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 5).getDisplayValue() + ")\n";

  notifyString += "Balance today:\t\t\t\t" + currentMonthSheet.getRange(currentMonthSheet.getLastRow(), 12).getDisplayValue() + "\n";
  notifyString += "Balance general:\t\t\t\t" + currentMonthSheet.getRange("I2:J2").getDisplayValue();

  // append free space before new log 
  currentMonthSheet.appendRow([" "]);

  notify(notifyString);
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

function addBalanceStat(currentMonthSheet) {

  var conditionalFormatRules = currentMonthSheet.getConditionalFormatRules();

  // scale with white in the middle
    //   .setGradientMinpointWithValue('#E67C73', SpreadsheetApp.InterpolationType.NUMBER, '-0,12') // -03:00:00
    // .setGradientMidpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, '0')
    // .setGradientMaxpointWithValue('#57BB8A', SpreadsheetApp.InterpolationType.NUMBER, '0,12')  // 03:00:00

  // rule for day balance
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([currentMonthSheet.getRange('L2:L1000')])
    .setGradientMinpointWithValue('#E67C73', SpreadsheetApp.InterpolationType.NUMBER, '-0,04') // -01:00:00
    .setGradientMidpointWithValue('#fffee1', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMaxpointWithValue('#57BB8A', SpreadsheetApp.InterpolationType.NUMBER, '0,04')  // 01:00:00
    .build());

  // rule for balance
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([currentMonthSheet.getRange('I1:J2')])
    .setGradientMinpointWithValue('#E67C73', SpreadsheetApp.InterpolationType.NUMBER, '-0,12') // -03:00:00
    .setGradientMidpointWithValue('#fffee1', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMaxpointWithValue('#57BB8A', SpreadsheetApp.InterpolationType.NUMBER, '0,12')  // 03:00:00


    
    .build());
  currentMonthSheet.setConditionalFormatRules(conditionalFormatRules);

  // balance
  currentMonthSheet.getRange("I2:J2").mergeAcross();
  currentMonthSheet.getRange('I1:J1').mergeAcross().setValue('Balance').setFontWeight('bold').setBackground("orange");
  currentMonthSheet.getRange("I2").setValue("=SUM(L:L)").setFontStyle('italic');
  currentMonthSheet.getRange('I1:J2').setHorizontalAlignment('center').setBackground("orange");
  currentMonthSheet.getRange('I2').setNumberFormat('[h]:mm:ss');
  // day balance
  currentMonthSheet.getRange('L:L').setNumberFormat('[h]:mm:ss');
  currentMonthSheet.getRange('L1').setHorizontalAlignment('center').setFontWeight('bold');


  // total
  currentMonthSheet.getRange('G1').setNumberFormat('[h]:mm:ss').setValue("Total").setHorizontalAlignment('center').setFontWeight('bold').setBackground("orange");
  currentMonthSheet.getRange('G2').setNumberFormat('[h]:mm:ss').setValue("=SUM(T:T)").setHorizontalAlignment('center').setFontStyle('italic').setBackground("orange");



};


