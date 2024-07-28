function checkOverdueInvoices() {
  var invoiceSheetName = 'E-Invoice'; // Name of the sheet where invoices are tracked
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invoiceSheet = ss.getSheetByName(invoiceSheetName);

  if (!invoiceSheet) {
    Logger.log("Sheet not found: " + invoiceSheetName);
    return;
  }

  var invoiceData = invoiceSheet.getDataRange().getValues();
  var currentDate = new Date();

  for (var i = 1; i < invoiceData.length; i++) { // Skip the header row
    var row = invoiceData[i];
    var dueDate = new Date(row[7]); // Assuming DueDate is the 8th column (H)
    var paymentStatus = row[10]; // Assuming Payment Status is the 11th column (K)
    var followUpEmail = row[12]; // Assuming Follow-up Email is the 13th column (M)

    if (paymentStatus === 'Overdue' && followUpEmail !== 'Sent') {
      sendFollowUpEmail(row);
      // Mark the row as email sent
      invoiceSheet.getRange(i + 1, 13).setValue('Sent');
    }
  }
}

function sendFollowUpEmail(row) {
  var customerName = row[2]; // Assuming CustomerName is the 3rd column (C)
  var customerEmail = row[3]; // Assuming Email is the 4th column (D)
  var projectName = row[4]; // Assuming ProjectName is the 5th column (E)
  var dueDate = row[7]; // Assuming DueDate is the 8th column (H)
  var amount = row[9]; // Assuming Amount is the 9th column (I)

  var subject = "Follow-Up: Overdue Invoice for " + projectName;
  var message = `
    <p>Dear ${customerName},</p>
    <p>We hope this message finds you well. We are writing to remind you that the invoice for the project <strong>${projectName}</strong> with the amount of <strong>$${amount}</strong> was due on <strong>${dueDate}</strong>.</p>
    <p>We kindly request you to settle the overdue amount at your earliest convenience. If you have already made the payment, please disregard this message. Otherwise, please let us know if there are any issues or if you need any further assistance.</p>
    <p>Thank you for your prompt attention to this matter.</p>
    <p>Best regards,<br>Your Company</p>
    <p><img src="https://i.postimg.cc/SNJt1nys/Screenshot-2024-07-25-at-12-17-41-AM.png" alt="101 Found Logo" width="150"></p>
  `;

  MailApp.sendEmail({
    to: customerEmail,
    subject: subject,
    htmlBody: message
  });

  Logger.log("Follow-up email sent to: " + customerEmail);
}

// Optional: Function to create a time-driven trigger
function createTimeDrivenTrigger() {
  ScriptApp.newTrigger('checkOverdueInvoices')
    .timeBased()
    .everyHours(1) // Adjust the frequency as needed
    .create();
}

