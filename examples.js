function selectA() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A:A').activate();

  

};

function colorchange() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D17').activate();
  spreadsheet.getCurrentCell().setValue('test');
  spreadsheet.getActiveRangeList().setFontWeight('bold')
  .setBackground('#3d85c6')
  .setBackground('ACCENT2')
  .setBackground('#cfe2f3');
};

function bgColorArea() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G:G').activate();
  spreadsheet.getActiveRangeList().setBackground('#93c47d');
  spreadsheet.getRange('D13:F27').activate();
  spreadsheet.getActiveRangeList().setBackground('#cc4125');
};


function hello(){
  Logger.log("Hello!");
}

function mergecells() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A32:E32').activate()
  .mergeAcross();
  spreadsheet.getCurrentCell().setValue('Monfsu');
  spreadsheet.getRange('A33').activate();
};

function resizeColumn() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G16').activate();
  spreadsheet.getActiveSheet().setColumnWidth(1, 212);
  spreadsheet.getActiveSheet().setColumnWidth(2, 190);
};

function resizeColumn1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G16').activate();
  spreadsheet.getActiveSheet().setColumnWidth(1, 108);
  spreadsheet.getActiveSheet().setColumnWidth(2, 216);
  spreadsheet.getActiveSheet().setColumnWidth(3, 30);
  spreadsheet.getActiveSheet().setColumnWidth(3, 26);
};

function boldAndCenter() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A6:E6').activate();
  spreadsheet.getActiveRangeList().setFontWeight('bold')
  .setHorizontalAlignment('center');
};