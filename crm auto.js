function sendCombinedFollowUpReminders() {
  var sheetCRM = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CRM_AUTO');
  var crmData = sheetCRM.getDataRange().getValues();
  var currentDate = new Date();
  var formLinkTemplate = "https://docs.google.com/forms/d/e/1FAIpQLSekrLHuWv9B7PGoGa3oNeIFLN2cncho0joxEG80RMs41xSmYg/viewform?usp=pp_url&entry.1513145006={{CUSTOMER_ID}}&entry.489628463={{CUSTOMER_NAME}}&entry.1228162848={{CUSTOMER_EMAIL}}"; // Replace with your actual form link and entry IDs
  var logoUrl = "https://i.postimg.cc/SNJt1nys/Screenshot-2024-07-25-at-12-17-41-AM.png"; // Replace with your actual logo URL

  for (var i = 1; i < crmData.length; i++) {
    var row = crmData[i];
    var customerID = row[0]; // Customer ID
    var customerName = row[1]; // Customer Name
    var email = row[2]; // Email
    var lastInteractionDate = new Date(row[8]); // Last Interaction Date (column I)
    var daysOfNoInteraction = parseInt(row[9]); // Days Of No Interaction (column J)
    var status = row[10]; // Status (column K)
    var nextFollowUpDate = new Date(row[11]); // Next Follow-Up Date (column L)
    var followUpSent = row[12]; // Follow Up Sent (column M)

    // Check if follow-up date is today or in the past and follow-up email has not been sent
    if (nextFollowUpDate <= currentDate && followUpSent === 'No' && status === 'Inactive') {
      var preFilledFormLink = formLinkTemplate.replace("{{CUSTOMER_ID}}", customerID)
                                              .replace("{{CUSTOMER_NAME}}", encodeURIComponent(customerName))
                                              .replace("{{CUSTOMER_EMAIL}}", encodeURIComponent(email)); // Replace placeholders with actual values
      var subject = "Follow-Up Reminder";
      var message = `
        <div style="font-family: Arial, sans-serif; color: #333;">
          <p>Dear ${customerName},</p>
          <p>We hope this message finds you well. This is a reminder to follow up on our previous interactions. Please let us know how we can assist you by clicking the link below:</p>
          <p><a href="${preFilledFormLink}" style="color: #1a73e8;">Follow-Up Form</a></p>
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
      
      // Send the follow-up email
      Logger.log("Sending follow-up email to: " + email);
      GmailApp.sendEmail(email, subject, message, {htmlBody: message});
      Logger.log("Follow-up email sent to: " + email);
      
      // Update the FollowUpSent column to "Yes"
      sheetCRM.getRange(i + 1, 13).setValue('Yes'); // Update the FollowUpSent column (column M)
      Logger.log("Updated FollowUpSent to Yes for Customer ID: " + customerID);
    }
  }
}

function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() === 'CRM_AUTO') {
    var range = e.range;
    var row = range.getRow();
    var column = range.getColumn();
    if (column === 11) { // Assuming the status column is the 11th column (K)
      var status = range.getValue();
      if (status === 'Inactive') {
        sendCombinedFollowUpReminders();
      }
    }
  }
}
