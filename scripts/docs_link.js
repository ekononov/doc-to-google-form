function linkFormToSheet(args) {
  // Passing Form and Spreadsheet ID as call arguments
  var formId = args.formId;
  var spreadsheetId = args.spreadsheetId;
  
  var form = FormApp.openById(formId);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);
  
  return { status: "success", message: "Form is successfully linked to Spreadsheet" };
}