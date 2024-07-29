function sendFollowUpEmails() {
  // Get the Leads_Data sheet

  var spreadsheetId = '1UlsuYn0Nvdl_RKVpkpmKs_PY48tmqI6MwPmGJWWqdIw';
  var sheetLeads = SpreadsheetApp.openById(spreadsheetId).getSheetByName('Leads_Data');
  if (!sheetLeads) {
    Logger.log("Error: Unable to find 'Leads_Data' sheet.");
    return;
  }
  
  // Get all lead data
  var leadData = sheetLeads.getDataRange().getValues();
  
  // Get the current date
  var currentDate = new Date();
  
  // Iterate through all rows in the sheet, starting from row 2 to skip headers
  for (var rowIndex = 1; rowIndex < leadData.length; rowIndex++) {
    var leadRow = leadData[rowIndex];
    
    if (leadRow.length < 8) {
      Logger.log("Skipping row " + (rowIndex + 1) + ": insufficient data.");
      continue;
    }
    
    var leadID = leadRow[0]; // LeadID (Column A)
    var leadName = leadRow[1]; // LeadName (Column B)
    var email = leadRow[2]; // Email (Column C)
    var followUpDate = new Date(leadRow[4]); // FollowUpDate (Column E)
    var lastContactDate = new Date(leadRow[5]); // LastContactDate (Column F)
    var status = leadRow[6]; // Status (Column G)
    var emailSent = leadRow[7]; // EmailSent (Column H)
    
    Logger.log("Processing Lead ID: " + leadID + ", Status: " + status + ", Last Contact Date: " + lastContactDate + ", Email Sent: " + emailSent);
    
    // Check if the conditions for sending the follow-up email are met
    if (emailSent === 'No' && status === 'Interested') {
      var subject = "Follow-Up on Our Recent Campaign";
      var message = `
        <div style="font-family: Arial, sans-serif; color: #333;">
          <p>Dear ${leadName},</p>
          <p>Thanks for showing your interest in our product. We're excited about the opportunity to work with you and help your business achieve!</p>
          <p>As a follow-up, we wanted to provide you with some valuable resources to help you further understand how we can benefit your organization:</p>
          <p>
            1) <a href='https://www.101found.com/it-infrastructure' style='color: #1a73e8;'>IT Infrastructure</a><br>
            2) <a href='https://www.101found.com/software-development' style='color: #1a73e8;'>Software Development and Integration</a><br>
            3) <a href='https://www.101found.com/cybersecurity-solutions' style='color: #1a73e8;'>Cybersecurity Solutions</a><br>
          </p>
          <p>We’d love to schedule a follow-up call to discuss any questions you may have and explore how we can tailor our solution to meet your specific needs. Please let us know a convenient time for you, or you can book a slot directly on our calendar <a href='https://calendar.app.google/NVeCtW1TL5vTu5FZA' style='color: #1a73e8;'>here</a>.</p>
          <p>Thank you again for your interest. We look forward to the opportunity to partner with you.</p>
          <p>Best regards,</p>
          <p>101 Found</p>
          <img src='https://i.postimg.cc/SNJt1nys/Screenshot-2024-07-25-at-12-17-41-AM.png' alt='101 Found Logo' style='width: 350px;'><br>
        </div>
      `;
      
      Logger.log("Sending follow-up email to: " + email);
      try {
        // Send the follow-up email
        GmailApp.sendEmail(email, subject, message, {htmlBody: message});
        Logger.log("Follow-up email sent to: " + email);
        
        // Update the last contact date to today
        sheetLeads.getRange(rowIndex + 1, 6).setValue(currentDate); // Column F is the 6th column for LastContactDate
        // Mark email as sent
        sheetLeads.getRange(rowIndex + 1, 8).setValue('Yes'); // Column H is the 8th column for EmailSent
      } catch (e) {
        // Log any errors that occur during email sending
        Logger.log("Failed to send email to: " + email + ", Error: " + e.toString());
      }
    }
  }
}