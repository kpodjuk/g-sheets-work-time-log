// Email constants
const EMAIL_SUBJECT = 'Work time log raport for ';
const EMAIL_BODY = 'Hello!\rPlease see the attached PDF document.';
const EMAIL_ADDRESS_OVERRIDE = "kpfixero901@gmail.com"

function sendEmails(pdfLink, pdfDates) {
  const fileId = pdfLink.match(/[-\w]{25,}(?!.*[-\w]{25,})/)
  const attachment = DriveApp.getFileById(fileId);

  const recipient = EMAIL_ADDRESS_OVERRIDE

  GmailApp.sendEmail(recipient, EMAIL_SUBJECT+pdfDates, EMAIL_BODY, {
    attachments: [attachment.getAs(MimeType.PDF)],
    name: APP_TITLE
  });

}

function onOpen(e) {
  console.log(e);
  const menu = SpreadsheetApp.getUi().createMenu(APP_TITLE)
  menu
    .addItem('Send this raport', 'sendPDF')
    .addSeparator()
    .addItem('Reset template', 'clearTemplateSheet')
    .addToUi();
}