function sendMailMergeWithAttachment() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const folder = DriveApp.getFolderById('1U1JFAERSCEYSU00jShx0d0zsK3TXaLB1');


  try {
    var recipientEmail = '';
    var recipientName = '';
    var i = 1;
    for (i = 1; i < data.length; i++) {
      const row = data[i];
      
      var status = row[4]
      if(status !== 'ok'){
        recipientEmail = row[3];
        const subject = "Thank you for your participation in Officer Training/Grazie per aver partecipato alla formazione per dirigenti";
        
        recipientName = row[0];
        const body = readBodyContentFromDrive().replace("{{1}}",recipientName).replace("{{2}}",recipientName);
        
        const fileName ="COT-Participation_"+i.toString(10)+".pdf"
        const fileByName = folder.getFilesByName(fileName);
        const file = DriveApp.getFileById(fileByName.next().getId());

        const attachments = [file]

      
          MailApp.sendEmail({
            to: recipientEmail,
            bcc:"<put your email address here>",
            subject: subject,
            htmlBody: body, // Assuming HTML body
            attachments: attachments
          });
          Logger.log(`Email sent to ${recipientName} at ${recipientEmail}`);
          sheet.getCurrentCell()
          setCellValue(sheet,'E'+(i+1),'ok');
      } else {
        Logger.log(`Skipping! Email to ${recipientName} at ${recipientEmail} already sent`)
      }
    }
  } catch (error) {
      Logger.log(`Error sending email to ${recipientName} at ${recipientEmail}: ${error}`);
      setCellValue(sheet,'E'+(i+1),'error');
  }
}

function setCellValue(sheet, range, newValue){
  var cell = sheet.getRange(range);   
  cell.setValue(newValue);
}

function readBodyContentFromDrive() {
  const fileName = 'body.txt'; // Replace with the actual file ID

  const folder = DriveApp.getFolderById('16GaXOVWFYYxb-6zU19F_tVpNoqx4ipLH');
  const fileByName = folder.getFilesByName(fileName);
  const file = DriveApp.getFileById(fileByName.next().getId());

  // Assuming a text file
  const content = file.getBlob().getDataAsString();
  return content;
}

