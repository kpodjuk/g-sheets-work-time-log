function getCalendarWeek(date) {
  currentDate = date;
  startDate = new Date(currentDate.getFullYear(), 0, 1);
  var days = Math.floor((currentDate - startDate) /
    (24 * 60 * 60 * 1000));

  return Math.ceil(days / 7);
  // Display the calculated result  
  // console.log("Week number of " + currentDate +
  // " is : " + weekNumber);

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

function addRaportButton(sheet) {
  var image = sheet.insertImage("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQg73FxMo2IUhWG4n28zAtBEprZuVn51qlhntW_qlFBln0OjnjhrRE1_OADbFV7YtDmxts&usqp=CAU", 1, 4); //change the URL to the image you prefer

  image.assignScript("generateReport"); //assign the function to the image
  image.setAnchorCell(sheet.getRange('k1')).setHeight(95).setWidth(95);

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