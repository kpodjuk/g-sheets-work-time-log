const OUTPUT_FOLDER_NAME = "Work time PDFs";
const APP_TITLE = "Work time log"

function sendPDF() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ssId = ss.getId();
  let templateSheet = ss.getActiveSheet();
  let sheetNameParts = templateSheet.getName().split(" ");

  let sheetDate = new Date(sheetNameParts[1] + sheetNameParts[0]);
  let sheetName = sheetDate.toLocaleDateString('en-US', { year: 'numeric', month: 'long' });

  SpreadsheetApp.flush();
  Utilities.sleep(500); // Using to offset any potential latency in creating .pdf
  const pdf = createPDF(ssId, templateSheet, sheetName);

  pdfUrl = pdf.getUrl();
  Logger.log("Created pdf: " + pdfUrl);

  sendEmails(pdfUrl, sheetName);

}

function createPDF(ssId, sheet, pdfName) {
  const fr = 0, fc = 0, lc = 9, lr = 27;
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
    "?format=pdf&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.5&" +
    "bottom_margin=0.25&" +
    "left_margin=0.5&" +
    "right_margin=0.5&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + sheet.getSheetId() + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  const params = {
    method: "GET",
    headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  }
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(pdfName + '.pdf');

  // Gets the folder in Drive where the PDFs are stored.
  const folder = getFolderByName_(OUTPUT_FOLDER_NAME);

  const pdfFile = folder.createFile(blob);
  return pdfFile;
}

function getFolderByName_(folderName) {

  // Gets the Drive Folder of where the current spreadsheet is located.
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const parentFolder = DriveApp.getFileById(ssId).getParents().next();

  // Iterates the subfolders to check if the PDF folder already exists.
  const subFolders = parentFolder.getFolders();
  while (subFolders.hasNext()) {
    let folder = subFolders.next();

    // Returns the existing folder if found.
    if (folder.getName() === folderName) {
      return folder;
    }
  }
  // Creates a new folder if one does not already exist.
  return parentFolder.createFolder(folderName)
    .setDescription(`Created by ${APP_TITLE} application to store PDF output files`);
}


function addSendPdfButton(sheet) {
  var image = sheet.insertImage("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQg73FxMo2IUhWG4n28zAtBEprZuVn51qlhntW_qlFBln0OjnjhrRE1_OADbFV7YtDmxts&usqp=CAU", 1, 4); //change the URL to the image you prefer

  image.assignScript("sendPDF"); //assign the function to the image
  image.setAnchorCell(sheet.getRange('f1')).setHeight(95).setWidth(95);
}