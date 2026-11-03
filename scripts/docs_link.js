function linkFormToSheet(formId, spreadsheetId) {
  var form = FormApp.openById(formId);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);
  
  return { status: "success", message: "Form is successfully linked to Spreadsheet" };
}