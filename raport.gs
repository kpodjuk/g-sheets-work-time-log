function generateReport() {

  let ss = SpreadsheetApp.getActiveSheet();

  // Detect which month/year is this sheet from
  let sheetNameParts = ss.getName().split(" ");
  let sheetDate = new Date(sheetNameParts[1] + sheetNameParts[0]);

  // find rows containing data
  const depth = 300;
  // label max search depth, just in case
  ss.getRange(depth, 12, 1, 10).setBackground("pink");

  let allRows = [];

  for (let i = 1; i < depth; i++) {
    // iterate over rows in column P
    searchRange = ss.getRange(i, 16, 1, 1);
    if (searchRange.getValue() != "") {
      // row not empty
      rowData = ss.getRange(i, 16, 1, 5).getDisplayValues()[0]; // get data from row
      allRows.push(rowData);
    }
  }

  Logger.log("Data in raport:");
  Logger.log(allRows);

  // check if raport for that month doesn't already exist
  let thisMonthRaportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetDate.toLocaleDateString('en-US', { year: 'numeric', month: 'long' }) + " RAPORT");
  if (thisMonthRaportSheet == null) {
    Logger.log("Creating new raport!");
    // add new one
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetDate.toLocaleDateString('en-US', { year: 'numeric', month: 'long' }) + " RAPORT",
      // reports go at the end
      SpreadsheetApp.getActiveSpreadsheet().getNumSheets() + 1);
  } else {
    Logger.log("Regenerating already existing raport!");

    // does, so delete old and insert new
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(thisMonthRaportSheet);
    SpreadsheetApp.flush();
    // add new one
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetDate.toLocaleDateString('en-US', { year: 'numeric', month: 'long' }) + " RAPORT",
      // reports go at the end
      SpreadsheetApp.getActiveSpreadsheet().getNumSheets() + 1);
  }


  // wait for sheet to be created
  SpreadsheetApp.flush();

  let raportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetDate.toLocaleDateString('en-US', { year: 'numeric', month: 'long' }) + " RAPORT");

  // fill with data
  raportSheet.getRange(1, 1, allRows.length, 5).setValues(allRows);

  // stylize
  stylizeReport(allRows.length);


  // Browser.msgBox("Month raport was properly generated!");
}

// super ugly macro recording but does the job
function stylizeReport(numberOfRows) {

  var spreadsheet = SpreadsheetApp.getActive();

  spreadsheet.getActiveSheet().setColumnWidth(1, 95);
  spreadsheet.getActiveSheet().setColumnWidth(2, 79);
  spreadsheet.getActiveSheet().setColumnWidth(3, 80);
  spreadsheet.getActiveSheet().setColumnWidth(4, 58);
  spreadsheet.getActiveSheet().setColumnWidth(4, 58);


  spreadsheet.getRange('1:1').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('2:2').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('A1:E1').mergeAcross();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('A2').setValue('Date');
  spreadsheet.getRange('B2').setValue('Start time');
  spreadsheet.getRange('C2').setValue('End time');
  spreadsheet.getRange('D2').setValue('Break time');
  spreadsheet.getRange('E2').setValue('Work time');
  spreadsheet.getRange('A1:E2').setFontWeight('bold').setHorizontalAlignment('center');

  // footer color
  spreadsheet.getActiveSheet().getRange(2, 1, numberOfRows + 2, 5).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  spreadsheet.getActiveSheet().getRange(2, 1, numberOfRows + 2, 5).setHorizontalAlignment('center')
  var banding = spreadsheet.getRange('A2:E12').getBandings()[0];
  banding.setHeaderRowColor('#bdbdbd')
    .setFirstRowColor('#ffffff')
    .setSecondRowColor('#f3f3f3')
    .setFooterRowColor('#dedede');
  spreadsheet.getRange('A1:E1').activate();
  spreadsheet.getActiveRangeList().setFontWeight('bold')
    .setBackground('#666666')
    .setBackground('#b7b7b7')
    .setBackground('#999999');

  // get month from sheet name
  let sheetNameParts = spreadsheet.getActiveSheet().getName().split(" ");
  // part 0 and 1 = month+year
  spreadsheet.getCurrentCell().setValue('Kamil Podjuk working time sheet for ' + sheetNameParts[0] + " " + sheetNameParts[1]);

  spreadsheet.getActiveSheet().getRange(numberOfRows + 3, 4, 1, 1).setFontWeight('bold').setValue('Total:');
  spreadsheet.getActiveSheet().getRange(numberOfRows + 3, 5, 1, 1).setFontStyle('italic').setValue('=SUM(E3:INDIRECT(ADDRESS(ROW()-1;COLUMN())))').setNumberFormat('[h]:mm:ss');;


  spreadsheet.getRange('2:2').activate();
  spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveRange().getLastRow(), 1);
  spreadsheet.getActiveRange().offset(spreadsheet.getActiveRange().getNumRows(), 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('D2:D3').activate()
    .mergeVertically();
  spreadsheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setVerticalAlignment('middle');
  spreadsheet.getRange('E2').activate();
  spreadsheet.getRange('D2:D3').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  spreadsheet.getRange('C2').activate();
  spreadsheet.getRange('D2:D3').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  spreadsheet.getRange('B2').activate();
  spreadsheet.getRange('C2:C3').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  spreadsheet.getRange('A2').activate();
  spreadsheet.getRange('B2:B3').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  // tab color
  spreadsheet.getActiveSheet().setTabColor('orange');

};