function sendOverdueReminder(rowNo) {
  var spSheet = SpreadsheetApp.getActiveSpreadsheet();
  var invoiceSheet = spSheet.getSheetByName("E-Invoice");
  var priceSheet = spSheet.getSheetByName("Price_LookUp");

  var invoiceTemplate = DriveApp.getFileById("1O4MnzzG5pFfEV-cHGbiebqcb9GltWHJS7msek_hgkq0");

  var invoiceNo = invoiceSheet.getRange("A" + rowNo).getValue();
  var invoiceDate = invoiceSheet.getRange("B" + rowNo).getValue();
  var dueDate = invoiceSheet.getRange("C" + rowNo).getValue();
  var custID  = invoiceSheet.getRange("D" + rowNo).getValue();
  var custName = invoiceSheet.getRange("E" + rowNo).getValue();
  var custEmail = invoiceSheet.getRange("F" + rowNo).getValue();
  var projectName = invoiceSheet.getRange("G" + rowNo).getValue();

  var quantityItem1 = invoiceSheet.getRange("H" + rowNo).getValue();
  var quantityItem2 = invoiceSheet.getRange("I" + rowNo).getValue();
  var quantityItem3 = invoiceSheet.getRange("J" + rowNo).getValue();

  var priceItem1 = priceSheet.getRange("B2").getValue();
  var priceItem2 = priceSheet.getRange("B3").getValue();
  var priceItem3 = priceSheet.getRange("B4").getValue();
  var taxPercentage = priceSheet.getRange("D2").getValue();

  // Convert tax percentage from decimal to percentage format
  var taxPercentageFormatted = (taxPercentage * 100).toFixed(2);

  var totalPiceItem1 = quantityItem1 * priceItem1;
  var totalPiceItem2 = quantityItem2 * priceItem2;
  var totalPiceItem3 = quantityItem3 * priceItem3;

  var subtotalPrice = totalPiceItem1 + totalPiceItem2 + totalPiceItem3;
  var taxAmount = parseFloat(taxPercentage * subtotalPrice).toFixed(2);
  var totalPrice = (Number(subtotalPrice) + Number(taxAmount)).toFixed(2);

  var formattedInvoiceDate = Utilities.formatDate(new Date(invoiceDate), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var formattedDueDate = Utilities.formatDate(new Date(dueDate), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  var rawFile = DocumentApp.openById(invoiceTemplate.getId());
  var rawFileContent = rawFile.getBody();

  var logoUrl = "https://i.postimg.cc/SNJt1nys/Screenshot-2024-07-25-at-12-17-41-AM.png";

  // invoiceSheet.getRange("L" + rowNo).setValue(totalPrice);

  rawFileContent.replaceText("{Invoice Number}", invoiceNo);
  rawFileContent.replaceText("{Invoice Date}", formattedInvoiceDate);
  rawFileContent.replaceText("{custName}", custName);
  rawFileContent.replaceText("{Email}", custEmail);
  rawFileContent.replaceText("{projectName}", projectName);

  rawFileContent.replaceText("I1Q", quantityItem1);
  rawFileContent.replaceText("I2Q", quantityItem2);
  rawFileContent.replaceText("I3Q", quantityItem3);

  rawFileContent.replaceText("I1P", priceItem1);
  rawFileContent.replaceText("I2P", priceItem2);
  rawFileContent.replaceText("I3P", priceItem3);

  rawFileContent.replaceText("T1", totalPiceItem1);
  rawFileContent.replaceText("T2", totalPiceItem2);
  rawFileContent.replaceText("T3", totalPiceItem3);

  rawFileContent.replaceText("{sub_total}", subtotalPrice);
  rawFileContent.replaceText("{Tax Percentage}", taxPercentageFormatted);
  rawFileContent.replaceText("{Tax}", taxAmount);
  rawFileContent.replaceText("{total}", totalPrice);

  rawFile.saveAndClose();

  var pdfInvoice = rawFile.getAs(MimeType.PDF);
  var subject = "Payment Overdue Notice - Invoice #" + invoiceNo;
  var body = `
    <div style="font-family: Arial, sans-serif; color: #333;">
      <p>Dear ${custName},</p>
      <p>We hope this message finds you well. This is a reminder that your payment for Invoice #${invoiceNo}, which was due on <b>${dueDate.toLocaleDateString()}</b>, is now overdue. <b>The total amount due is RM${totalPrice}</b>.</p>
      <p>Please make the payment as soon as possible to avoid any late fees or service interruptions.</p>
      <p>If you have already made the payment, please disregard this notice. Otherwise, we would appreciate it if you could process the payment at your earliest convenience.</p>
      <p>Thank you for your prompt attention to this matter.</p>
      <p>Best regards,</p>
      <p>101 Found</p>
      <img src="${logoUrl}" width="350" alt="101 Found Logo" style="display: block; margin-top: 20px;">
      <hr>
      <p><strong>101 Found IT Consulting</strong></p>
      <p style="color: #666;">Providing clients with innovative IT solutions.</p>
      <p>Contact us at: <a href="mailto:info@101found.com" style="color: #1a73e8;">info@101found.com</a></p>
      <p>Follow us on: 
        <a href="https://www.linkedin.com/company/101found" style="color: #1a73e8;">LinkedIn</a> | 
        <a href="https://twitter.com/101found" style="color: #1a73e8;">Twitter</a> | 
        <a href="https://www.facebook.com/101found" style="color: #1a73e8;">Facebook</a>
      </p>
      <p style="font-size: 0.8em; color: #999;">You are receiving this email because you are a valued customer of 101 Found.</p>
    </div>
  `;

  // Send the email
  GmailApp.sendEmail(custEmail, subject, body, {htmlBody: body, attachments: [pdfInvoice.getAs(MimeType.PDF)]});
}


function updatePaymentStatus() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("E-Invoice");
  var today = new Date();
  var lastRow = sheet.getLastRow();

  for (var i = 2; i <= lastRow; i++) {
    var dueDate = sheet.getRange("C" + i).getValue();
    var paymentStatus = sheet.getRange("M" + i).getValue();
    var overdueFlag = sheet.getRange("N" + i).getValue(); // Column to track if overdue email was sent

    if (dueDate instanceof Date && paymentStatus) {
      if (paymentStatus.toLowerCase() !== "paid" && dueDate < today) {
        sheet.getRange("M" + i).setValue("Overdue"); // Set to "Overdue" if due date has passed

        if (overdueFlag !== "Sent") {
          sendOverdueReminder(i);
          sheet.getRange("N" + i).setValue("Sent"); // Mark that the overdue email was sent
        }
      }
    }
  }
}


function eInvoice(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  
  // Define the column for "Payment Status" (adjust the column number if necessary)
  var paymentStatusCol = 13; 

  // Check if the edited cell is in the "Payment Status" column
  if (range.getColumn() == paymentStatusCol) {
    var row = range.getRow();
    var paymentStatus = range.getValue();

    // If the payment status is "Paid", create and send the invoice
    if (paymentStatus.toLowerCase() == "paid") {
      sendInvoice(row);
    }
  }
}


function sendInvoice(rowNo) {
  var spSheet = SpreadsheetApp.getActiveSpreadsheet();
  var invoiceSheet = spSheet.getSheetByName("E-Invoice");
  var priceSheet = spSheet.getSheetByName("Price_LookUp");

  var invoiceFolder = DriveApp.getFolderById("1Qh_aqP2lGtr3mA6UmsvNixHLELB7pNSe");
  var invoiceTemplate = DriveApp.getFileById("1O4MnzzG5pFfEV-cHGbiebqcb9GltWHJS7msek_hgkq0");

  var invoiceNo = invoiceSheet.getRange("A" + rowNo).getValue();
  var invoiceDate = invoiceSheet.getRange("B" + rowNo).getValue();
  var dueDate = invoiceSheet.getRange("C" + rowNo).getValue();
  var custID  = invoiceSheet.getRange("D" + rowNo).getValue();
  var custName = invoiceSheet.getRange("E" + rowNo).getValue();
  var custEmail = invoiceSheet.getRange("F" + rowNo).getValue();
  var projectName = invoiceSheet.getRange("G" + rowNo).getValue();

  var quantityItem1 = invoiceSheet.getRange("H" + rowNo).getValue();
  var quantityItem2 = invoiceSheet.getRange("I" + rowNo).getValue();
  var quantityItem3 = invoiceSheet.getRange("J" + rowNo).getValue();

  var priceItem1 = priceSheet.getRange("B2").getValue();
  var priceItem2 = priceSheet.getRange("B3").getValue();
  var priceItem3 = priceSheet.getRange("B4").getValue();
  var taxPercentage = priceSheet.getRange("D2").getValue();

  // Convert tax percentage from decimal to percentage format
  var taxPercentageFormatted = (taxPercentage * 100).toFixed(2);

  var totalPiceItem1 = quantityItem1 * priceItem1;
  var totalPiceItem2 = quantityItem2 * priceItem2;
  var totalPiceItem3 = quantityItem3 * priceItem3;

  var subtotalPrice = totalPiceItem1 + totalPiceItem2 + totalPiceItem3;
  var taxAmount = parseFloat(taxPercentage * subtotalPrice).toFixed(2);
  var totalPrice = (Number(subtotalPrice) + Number(taxAmount)).toFixed(2);

  var formattedInvoiceDate = Utilities.formatDate(new Date(invoiceDate), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var formattedDueDate = Utilities.formatDate(new Date(dueDate), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  var rawInvoiceFile = invoiceTemplate.makeCopy(invoiceFolder);
  var rawFile = DocumentApp.openById(rawInvoiceFile.getId());
  var rawFileContent = rawFile.getBody();

  var logoUrl = "https://i.postimg.cc/SNJt1nys/Screenshot-2024-07-25-at-12-17-41-AM.png";

  // invoiceSheet.getRange("L" + rowNo).setValue(totalPrice);

  rawFileContent.replaceText("{Invoice Number}", invoiceNo);
  rawFileContent.replaceText("{Invoice Date}", formattedInvoiceDate);
  rawFileContent.replaceText("{custName}", custName);
  rawFileContent.replaceText("{Email}", custEmail);
  rawFileContent.replaceText("{projectName}", projectName);

  rawFileContent.replaceText("I1Q", quantityItem1);
  rawFileContent.replaceText("I2Q", quantityItem2);
  rawFileContent.replaceText("I3Q", quantityItem3);

  rawFileContent.replaceText("I1P", priceItem1);
  rawFileContent.replaceText("I2P", priceItem2);
  rawFileContent.replaceText("I3P", priceItem3);

  rawFileContent.replaceText("T1", totalPiceItem1);
  rawFileContent.replaceText("T2", totalPiceItem2);
  rawFileContent.replaceText("T3", totalPiceItem3);

  rawFileContent.replaceText("{sub_total}", subtotalPrice);
  rawFileContent.replaceText("{Tax Percentage}", taxPercentageFormatted);
  rawFileContent.replaceText("{Tax}", taxAmount);
  rawFileContent.replaceText("{total}", totalPrice);

  rawFile.saveAndClose();

  var pdfInvoice = rawFile.getAs(MimeType.PDF);
  pdfInvoice = invoiceFolder.createFile(pdfInvoice).setName("Invoice_" + custID);

  // Get the URL of the PDF file
  var pdfUrl = pdfInvoice.getUrl();

  // Update the spreadsheet with the PDF URL
  invoiceSheet.getRange("K" + rowNo).setValue(pdfUrl);
  
  invoiceFolder.removeFile(rawInvoiceFile);

  var mailSubject = "Invoice #" + invoiceNo + " from 101FOUND";
  var body = `
  <div style="font-family: Arial, sans-serif; color: #333;">
    <p>Dear ${custName},</p>
    <p>We hope this message finds you well.</p>
    <p>Thank you for your recent payment. We are pleased to confirm that we have received the payment for the <b>e-invoice ${invoiceNo}</b></p>      
    <p>For your records, please find the attached e-invoice. If you have any questions or require additional information, please feel free to contact us.</p>
    <p>We appreciate your business and look forward to serving you again.</p>
    <p>Best regards,</p>
    <p>101 Found</p>
    <img src="${logoUrl}" width="350" alt="101 Found Logo" style="display: block; margin-top: 20px;">
    <hr>
    <p><strong>101 Found IT Consulting</strong></p>
    <p style="color: #666;">Providing clients with innovative IT solutions.</p>
    <p>Contact us at: <a href="mailto:info@101found.com" style="color: #1a73e8;">info@101found.com</a></p>
    <p>Follow us on: 
      <a href="https://www.linkedin.com/company/101found" style="color: #1a73e8;">LinkedIn</a> | 
      <a href="https://twitter.com/101found" style="color: #1a73e8;">Twitter</a> | 
      <a href="https://www.facebook.com/101found" style="color: #1a73e8;">Facebook</a>
    </p>
    <p style="font-size: 0.8em; color: #999;">You are receiving this email because you are a valued customer of 101 Found.</p>
  </div>
  `;

  GmailApp.sendEmail(custEmail, mailSubject, body, {htmlBody: body, attachments: [pdfInvoice.getAs(MimeType.PDF)]});


  // Return the PDF invoice to be used in the updatePaymentStatus function
  return pdfInvoice;
}
