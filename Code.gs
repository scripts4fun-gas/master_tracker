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
      // Sales headers: include Vendor ID and Delivery ID and three link columns
      sheet.getRange(1, 1, 1, 11).setValues([[
        'Sale ID', 'Sale PO Number', 'Date of PO', 'Appointment Date', 'Invoice Number',
        'Vendor ID', 'Delivery ID', 'PO Link', 'Invoice Link', 'EWay Link', '...materials start here'
      ]]);
    } else if (sheetName === PO_SHEET_NAME) {
      // Purchase headers: Internal ID, Vendor PO ID, PO Date, Despatch Date, Invoice, Supplier ID, PO Link, Inv Link, Eway Link
      sheet.getRange(1, 1, 1, 10).setValues([[
        'Internal PO ID', 'Vendor PO ID', 'PO Date', 'Despatch Date', 'Invoice Number',
        'Supplier ID', 'PO Link', 'Inv Link', 'Eway Link', '...materials start here'
      ]]);
    } else if (sheetName === MANUAL_SHEET_NAME) {
      // Manual sheet: Internal ID + Timestamp; material columns appended dynamically
      sheet.getRange(1, 1, 1, 2).setValues([['Internal Manual ID', 'Timestamp']]);
    } else if (sheetName === DATA_SHEET_NAME) {
      // Data sheet: first column is label ('Metric'), material headers start from column B
      sheet.getRange(1, 1, 1, 1).setValues([['Metric']]);
    } else if (sheetName === VENDOR_SHEET_NAME) {
      sheet.getRange(1, 1, 1, 2).setValues([['Vendor ID', 'Vendor Name']]);
    } else if (sheetName === DELIVERY_SHEET_NAME) {
      sheet.getRange(1, 1, 1, 2).setValues([['Delivery ID', 'Delivery Name']]);
    } else if (sheetName === SUPPLIER_SHEET_NAME) {
      sheet.getRange(1, 1, 1, 2).setValues([['Supplier ID', 'Supplier Name']]);
    }
  }
  return sheet;
}

// NOTE: Additional functions like processEmailRequest and validateOTP from Auth.gs 
// would typically be included here or in a separate Auth.gs file, referencing this utility.