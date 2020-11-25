//ideas from: https://spreadsheet.dev/send-html-email-from-google-sheets


// This constant is written in column C for rows for which an email
// has been sent successfully.
var EMAIL_SENT = 'EMAIL_SENT';

/**
 * Sends non-duplicate emails with data from the current spreadsheet.
 */
function sendEmails(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var startCol = 1;
  var numRows = 290; // Number of rows to process
  var numCols = 4;
  var dataRange = sheet.getDataRange(); //use sheet.getRange(startRow, startCol, numRows, numCols) to retrieve specific range of rows/columns;

  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for(var i = 0; i < data.length; ++i){
    var row = data[i];
    var emailAddress = row[1]; // First column
    var link = row[3];
    var htmlMessage = getEmailHtml(link);
    var txtMessage = getEmailText(link);
    var emailSent = row[4]; // 4th column
    if(emailAddress && (emailSent !== EMAIL_SENT)){ // Prevents sending duplicates
      var subject = 'Actualização dos contactos dos clientes';
      MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        body: txtMessage,
        htmlBody: htmlMessage
      });
      sheet.getRange(startRow + i, 5).setValue(EMAIL_SENT); //mark the sent row with email sent
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}


function getEmailHtml(data){
  var htmlTemplate = HtmlService.createTemplateFromFile("Template.html");
  htmlTemplate.data = data;
  var htmlBody = htmlTemplate.evaluate().getContent();
  return htmlBody;
}

function getEmailText(data){
  var htmlTemplate = HtmlService.createTemplateFromFile("Template.txt");
  htmlTemplate.data = data;
  var htmlBody = htmlTemplate.evaluate().getContent();
  return htmlBody;
}
