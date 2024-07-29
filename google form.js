function onFormSubmit(e) {
  // Assuming e.values contains the submitted form data
  var formResponse = e.values;

  // Assuming the formResponse array matches the columns in the Interaction sheet
  var customerID = formResponse[1]; // Customer ID is in the second column
  var interactionDate = new Date(formResponse[0]); // Timestamp is in the first column

  // Get the CRM sheet

  var spreadsheetId = '1UlsuYn0Nvdl_RKVpkpmKs_PY48tmqI6MwPmGJWWqdIw';
  var sheetCRM = SpreadsheetApp.openById(spreadsheetId).getSheetByName('CRM_AUTO');
  var crmData = sheetCRM.getDataRange().getValues();

  // Update the CRM sheet
  for (var i = 1; i < crmData.length; i++) {
    var row = crmData[i];
    if (row[0] == customerID) {
      // Update LastInteractionDate
      sheetCRM.getRange(i + 1, 9).setValue(interactionDate);

      // The DaysOfNoInteraction will be recalculated by your formula in Google Sheets

      // Optionally reset FollowUpSent if needed
      sheetCRM.getRange(i + 1, 13).setValue('No');

      Logger.log("Updated interaction for Customer ID: " + customerID);
      break;
    }
  }
}