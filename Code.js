//ideas from: https://spreadsheet.dev/send-html-email-from-google-sheets


// This constant is written in column C for rows for which an email
// has been sent successfully.
const EMAIL_SENT = 'EMAIL_SENT';

/**
 * Sends non-duplicate emails with data from the current spreadsheet.
 */
function sendEmails(){
  let sheet = SpreadsheetApp.getActiveSheet();
  let startRow = 2; // First row of data to process
  let startCol = 1;
  let numRows = 290; // Number of rows to process
  let numCols = 4;
  const emailSentColumn = 5;
  let dataRange = sheet.getDataRange(); //use sheet.getRange(startRow, startCol, numRows, numCols) to retrieve specific range of rows/columns;

  // Fetch values for each row in the Range.
  let data = dataRange.getValues();
  data.shift(); //remove headers

  data.forEach(function(row, index) {
    let emailAddress = row[1]; // First column
    let link = row[3];
    let htmlMessage = getEmailHtml(link);
    let txtMessage = getEmailText(link);
    let emailSent = row[4]; // 4th column
    let subject = 'Actualização dos contactos dos clientes';
    if(emailAddress && (emailSent !== EMAIL_SENT)){ // Prevents sending duplicates
      MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        body: txtMessage,
        htmlBody: htmlMessage
      });
      sheet.getRange(startRow + index, emailSentColumn).setValue(EMAIL_SENT); //mark the sent row with email sent
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  });
}


function getEmailHtml(data){
  let htmlTemplate = HtmlService.createTemplateFromFile("Template.html");
  htmlTemplate.data = data;
  let htmlBody = htmlTemplate.evaluate().getContent();
  return htmlBody;
}

function getEmailText(data){
  let htmlTemplate = HtmlService.createTemplateFromFile("Template.txt.html");
  htmlTemplate.data = data;
  let htmlBody = htmlTemplate.evaluate().getContent();
  return htmlBody;
}
