function onEditm(e) {
  if (!e || !e.range) {
    return;
  }

  var sheet = e.range.getSheet();
  var range = e.range;

  if (sheet.getName() === 'Leads_Data' && range.getColumn() === 3 && range.getRow() > 1) {
    var leadRow = sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn()).getValues()[0];
    var leadID = range.getRow() - 1; // Start leadID from 1, remove 'ID-' prefix
    var leadName = leadRow[1];
    var email = leadRow[2];
    var subject = "Follow-Up on Our Recent Campaign";
    var logoUrl = "https://i.postimg.cc/SNJt1nys/Screenshot-2024-07-25-at-12-17-41-AM.png"; // Replace with your actual logo URL
    var scriptUrl = "https://script.google.com/macros/s/AKfycbyKFdQz2UtcwBpWACd_qk5__HfP0dD3tIbDyb97NGgEnFbVcx4oO55cukECBR7uOK0j/exec"; // Replace with your actual deployed web app URL

    var message = `
      <div style="font-family: Arial, sans-serif; color: #333;">
        <p>Dear ${leadName},</p>
        <p>Thank you for joining our campaign! We're excited about the opportunity to work with you and help your business achieve great success.</p>
        <p>Please let us know if you are interested in learning more about our product and how it can benefit your business by selecting one of the options below:</p>
        <p>
          <a href="${scriptUrl}?leadID=${leadID}&status=Interested" style="color: #1a73e8;">Interested</a><br>
          <a href="${scriptUrl}?leadID=${leadID}&status=Not%20Interested" style="color: #1a73e8;">Not Interested</a>
        </p>
        <p>We're here to assist you every step of the way.</p>
        <p>Best regards,</p>
        <p>101 Found</p>
        <img src="${logoUrl}" width="200" alt="101 Found Logo" style="display: block; margin-top: 20px;">
        <hr>
        <p><strong>101 Found IT Consulting</strong></p>
        <p style="color: #666;">Providing clients with innovative IT solutions.</p>
        <p>Contact us at: <a href="mailto:info@company.com" style="color: #1a73e8;">info@company.com</a></p>
        <p>Follow us on:
          <a href="https://www.linkedin.com/company/company" style="color: #1a73e8;">LinkedIn</a> |
          <a href="https://twitter.com/company" style="color: #1a73e8;">Twitter</a> |
          <a href="https://www.facebook.com/company" style="color: #1a73e8;">Facebook</a>
        </p>
        <p style="font-size: 0.8em; color: #999;">You are receiving this email because you joined our recent campaign.</p>
      </div>
    `;
    
    GmailApp.sendEmail(email, subject, message, {htmlBody: message});
    sheet.getRange(range.getRow(), 6).setValue(new Date());
    sheet.getRange(range.getRow(), 7).setValue('Qualified');
    sheet.getRange(range.getRow(), 8).setValue('No');
    sheet.getRange(range.getRow(), 1).setValue(leadID); // Set the LeadID in column A
  }
}

function doGet(e) {
  var leadID = e.parameter.leadID;
  var status = e.parameter.status;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leads_Data');
  var leadData = sheet.getDataRange().getValues();

  // Log received parameters
  Logger.log("Received parameters: leadID=" + leadID + ", status=" + status);
  
  var found = false;
  
  // Iterate through the sheet to find the matching leadID
  for (var i = 1; i < leadData.length; i++) {
    if (leadData[i][0] == leadID) {
      // Log the row where the match was found
      Logger.log("Match found at row: " + (i + 1));
      
      // Update the status
      sheet.getRange(i + 1, 7).setValue(status);
      
      // Log the updated status
      Logger.log("Updated status to: " + status);
      
      found = true;
      break;
    }
  }
  
  if (!found) {
    Logger.log("No matching leadID found.");
  }
  
  // Return a simple HTML page confirming the status update
  return ContentService.createTextOutput("Thank you for your response. Your status has been updated to: " + status);
}

function setLeadID(sheet, row) {
  // Assuming LeadID is in the first column (A)
  var leadIDCell = sheet.getRange(row, 1);
  if (!leadIDCell.getValue()) {
    leadIDCell.setValue(row - 1); // Assuming LeadID is just the row number minus the header
    Logger.log("LeadID set to: " + (row - 1));
  }
}








