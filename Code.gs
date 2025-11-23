function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle('Inventory Tracker')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * Global utility to get a spreadsheet sheet by name.
 * It also handles the creation of the sheet if it doesn't exist and sets headers.
 * @param {string} sheetName The name of the sheet to retrieve.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet object.
 */
function getSheetByName(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log(`Created sheet: ${sheetName}`);
    
    // Set default headers for newly created sheets
    if (sheetName === OTP_SHEET_NAME) {
      sheet.getRange(1, 1, 1, 3).setValues([['Email', 'Timestamp', 'OTP']]);
    } else if (sheetName === MATERIAL_SHEET_NAME) {
      sheet.getRange(1, 1, 1, 2).setValues([['Product ID', 'Product Name']]);
    } else if (sheetName === SALES_SHEET_NAME) {
      sheet.getRange(1, 1, 1, 4).setValues([['Sale ID', 'Sale PO Number', 'Date of PO', 'Appointment Date']]);
    } else if (sheetName === PO_SHEET_NAME) {
      // Purchase headers are set to a minimum here (A to E)
      sheet.getRange(1, 1, 1, 5).setValues([['Internal PO ID', 'Vendor PO ID', 'PO Date', 'Despatch Date', 'Invoice Number']]);
    }
  }
  return sheet;
}

// NOTE: Additional functions like processEmailRequest and validateOTP from Auth.gs 
// would typically be included here or in a separate Auth.gs file, referencing this utility.