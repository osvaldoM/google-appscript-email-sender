// This constant is written in column C for rows for which an email
// has been sent successfully.
var EMAIL_SENT = 'EMAIL_SENT';

/**
 * Sends non-duplicate emails with data from the current spreadsheet.
 */
function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var numRows = 290; // Number of rows to process
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 1, numRows, 1);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0]; // First column
    var htmlMessage = `<p dir="ltr">Caro cliente,</p><p dir="ltr">A equipa emprego.co.mz vem atrav&eacute;s desta nota pedir desculpas pelo envio repetitivo de e-mails, no passado dia 06 de Fevereiro corrente.</p><p dir="ltr">Na expectativa de apresentar aos utilizadores as novas funcionalidades da plataforma, registou-se um erro no sistema de e-mail, que causou o envio de m&uacute;ltiplas mensagens para o mesmo remetente.</p><p dir="ltr">Pelos transtornos causados, a equipa emprego.co.mz pede sinceras desculpas.</p><p dir="ltr">&nbsp;&nbsp;</p><p dir="ltr">Cumprimentos,</p><p dir="ltr">A equipa emprego.co.mz</p>';`
    var emailSent = row[1]; // Second column
    if (emailSent !== EMAIL_SENT) { // Prevents sending duplicates
      var subject = 'emprego.co.mz: erro no envio de e-mail';
      MailApp.sendEmail({ito: emailAddress, subject: subject, htmlBody: htmlMessage});
      sheet.getRange(startRow + i, 2).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}