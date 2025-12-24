// Note: All constants are now defined in Constants.gs and are globally available here.
// getSheetByName is defined in Code.gs and is globally available.

/**
 * Ensures that the header row of a given sheet contains columns for all materials
 * listed in the Material sheet, starting at the specified column index.
 * Used for Purchase and Manual sheets.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to update (Purchase or Manual).
 * @param {number} materialStartIndex The 0-based index where material columns should start.
 */
function ensureMaterialHeadersExist(sheet, materialStartIndex) {
  const materialMap = getMaterialMap();
  const materialIds = Array.from(materialMap.keys());
  
  if (materialIds.length === 0) return;

  const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const currentMaterialHeaders = currentHeaders.slice(materialStartIndex);

  let headersToAppend = [];
  
  // Identify missing material IDs from the current headers
  const currentHeaderSet = new Set(currentMaterialHeaders.map(h => h.toString().trim()));

  for (const matId of materialIds) {
    if (!currentHeaderSet.has(matId)) {
      headersToAppend.push(matId);
    }
  }

  // If there are new materials, append them to the header row
  if (headersToAppend.length > 0) {
    // Append at the end of the current headers, which is the last column + 1
    const startCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, startCol, 1, headersToAppend.length).setValues([headersToAppend]);
    Logger.log(`Appended ${headersToAppend.length} material IDs to ${sheet.getName()} headers.`);
  }
}

/**
 * Ensures that the header row of the Sales sheet contains columns for all materials
 * for both PO and Dispatch quantities, starting at the specified column index.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Sales sheet to update.
 * @param {number} materialStartIndex The 0-based index where PO material columns should start.
 */
function ensureSalesMaterialHeadersExist(sheet, materialStartIndex) {
  const materialMap = getMaterialMap();
  const materialIds = Array.from(materialMap.keys());
  
  if (materialIds.length === 0) return;

  const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // We need: [Fixed Cols...] [Mat1, Mat2, Mat3...] [Mat1, Mat2, Mat3...]
  // First set = PO quantities, Second set = Dispatch quantities
  // Expected structure: materialStartIndex marks where first set begins
  // Current material headers start from materialStartIndex
  const currentMaterialHeaders = currentHeaders.slice(materialStartIndex);
  
  // Determine how many material columns we currently have
  const expectedTotalMaterialCols = materialIds.length * 2; // PO set + Dispatch set
  const missingCols = expectedTotalMaterialCols - currentMaterialHeaders.length;
  
  if (missingCols <= 0) return; // Already have all columns
  
  let headersToAppend = [];
  
  // If we have fewer than materialIds.length columns, we need to add PO columns first
  if (currentMaterialHeaders.length < materialIds.length) {
    // Add missing PO columns
    const existingPoHeaders = new Set(currentMaterialHeaders.slice(0, Math.min(currentMaterialHeaders.length, materialIds.length)).map(h => h.toString().trim()));
    for (const matId of materialIds) {
      if (!existingPoHeaders.has(matId)) {
        headersToAppend.push(matId);
      }
    }
  }
  
  // Then add Dispatch columns (second set of same material IDs)
  if (currentMaterialHeaders.length < expectedTotalMaterialCols) {
    const existingDispatchStart = Math.max(materialIds.length, currentMaterialHeaders.length);
    const existingDispatchHeaders = new Set(currentMaterialHeaders.slice(materialIds.length).map(h => h.toString().trim()));
    for (const matId of materialIds) {
      if (!existingDispatchHeaders.has(matId)) {
        headersToAppend.push(matId);
      }
    }
  }

  // If there are new headers, append them to the header row
  if (headersToAppend.length > 0) {
    const startCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, startCol, 1, headersToAppend.length).setValues([headersToAppend]);
    Logger.log(`Appended ${headersToAppend.length} material headers to Sales sheet.`);
  }
}

/**
 * Validates and inserts a new Purchase Order into the Purchase sheet.
 * @param {object} poData The purchase order data from the form.
 * @returns {object} An object containing success status and a message.
 */
function addPurchaseOrder(poData) {
  try {
    const poSheet = getSheetByName(PO_SHEET_NAME);
    
    // 0. Ensure Headers
    ensureMaterialHeadersExist(poSheet, PO_COL_FIRST_MATERIAL);
    
    // --- 1. Internal PO ID Generation (Using Counters Sheet) ---
    const countersSheet = getSheetByName(COUNTERS_SHEET_NAME);
    const countersData = countersSheet.getDataRange().getValues();
    
    // Find Purchase row in Counters sheet
    let purchaseRowIndex = -1;
    let currentCounter = 0;
    for (let i = 0; i < countersData.length; i++) { // Start from row 2 (skip header)
      if (countersData[i][COUNTERS_COL_TYPE] && countersData[i][COUNTERS_COL_TYPE].toString().trim() === 'Purchase') {
        purchaseRowIndex = i;
        currentCounter = parseInt(countersData[i][COUNTERS_COL_COUNTER]) || 0;
        break;
      }
    }
    
    if (purchaseRowIndex === -1) {
      return { error: true, message: "Purchase counter not found in Counters sheet." };
    }
    
    // Generate new ID
    const nextPoNumber = currentCounter + 1;
    const newInternalPoId = `P${('0000' + nextPoNumber).slice(-4)}`;
    
    // Update counter in Counters sheet
    countersSheet.getRange(purchaseRowIndex + 1, COUNTERS_COL_COUNTER + 1).setValue(nextPoNumber);
    // --- End PO ID Generation ---

    // 2. Prepare Data Row
    
    // Get the final list of material headers from the sheet to match column order
    const finalHeaders = poSheet.getRange(1, PO_COL_FIRST_MATERIAL + 1, 1, poSheet.getLastColumn() - PO_COL_FIRST_MATERIAL).getValues()[0].map(h => h.toString().trim());

    // Create a row template large enough for all columns
    const numColumns = poSheet.getLastColumn();
    const newRow = new Array(numColumns).fill(''); 

    // Insert fixed fields using the NEW constants
    newRow[P_COL_INTERNAL_ID] = newInternalPoId; // Column A: Generated Internal ID
    newRow[PO_COL_ID] = poData.poId;            // Column B: User-provided Vendor PO ID
    newRow[PO_COL_DATE] = poData.poDate ? new Date(poData.poDate) : '';
    newRow[PO_COL_DESPATCH_DATE] = poData.despatchDate ? new Date(poData.despatchDate) : '';
    newRow[PO_COL_INVOICE] = poData.invoiceNumber;
    newRow[PO_COL_SUPPLIER_ID] = poData.supplierId || '';
    
    // --- 3. Handle file uploads ---
    let poLink = '', invLink = '', ewayLink = '';
    if (poData.filesMeta && poData.filesMeta.length > 0) {
      // safe PURCHASE_FOLDER fallback
      const purchaseFolderName = (typeof PURCHASE_FOLDER !== 'undefined' && PURCHASE_FOLDER) ? PURCHASE_FOLDER : 'PurchaseInternal';

      // Folder path: PURCHASE folder -> YYYY/MM/DD/<INTERNAL_ID>
      const dateObj = poData.poDate ? new Date(poData.poDate) : new Date();
      const yyyy = dateObj.getFullYear();
      const mm = ('0' + (dateObj.getMonth() + 1)).slice(-2);
      const dd = ('0' + dateObj.getDate()).slice(-2);

      // Use Internal ID as folder name (no sanitization needed for P001 format)
      const subPath = `${purchaseFolderName}/${yyyy}/${mm}/${dd}/${newInternalPoId}`;
      const urls = uploadFilesToDrive(PARENT_FOLDER_ID, subPath, poData.filesMeta, poData.poId);
      // Order: [PO, Invoice, EWay]
      poLink = urls[0] || '';
      invLink = urls[1] || '';
      ewayLink = urls[2] || '';
    }
    newRow[PO_COL_PO_LINK] = poLink;
    newRow[PO_COL_INV_LINK] = invLink;
    newRow[PO_COL_EWAY_LINK] = ewayLink;
    
    // Insert material quantities starting from PO_COL_FIRST_MATERIAL
    finalHeaders.forEach((matId, index) => {
        const quantity = poData.materials[matId] || 0;
        // Check if matId exists in submitted data and is > 0
        if (quantity > 0) {
            newRow[PO_COL_FIRST_MATERIAL + index] = quantity;
        }
    });
    
    // 4. Append Data
    poSheet.appendRow(newRow);

    return { success: true, message: `Purchase Order ${poData.poId} recorded. Internal ID: ${newInternalPoId}.` };

  } catch (e) {
    Logger.log("Error in addPurchaseOrder: " + e.toString());
    return { error: true, message: "Server error during PO submission: " + e.toString() };
  }
}

/**
 * Checks if a customer PO number already exists in the Sales sheet.
 * @param {string} customerPoNumber The customer PO number to check.
 * @returns {boolean} True if the PO number already exists, false otherwise.
 */
function isDuplicateSalesPO(customerPoNumber) {
  try {
    const salesSheet = getSheetByName(SALES_SHEET_NAME);
    const salesData = salesSheet.getDataRange().getValues();
    
    // Skip header row and check all PO numbers
    for (let i = 1; i < salesData.length; i++) {
      const existingPO = salesData[i][SALES_COL_PO_NUMBER];
      if (existingPO && existingPO.toString().trim() === customerPoNumber.toString().trim()) {
        return true;
      }
    }
    
    return false;
  } catch (e) {
    Logger.log("Error in isDuplicateSalesPO: " + e.toString());
    return false; // If error, allow submission (fail open)
  }
}

/**
 * Validates and inserts a new Sales Order into the Sales sheet.
 * @param {object} salesData The sales order data from the form.
 * @returns {object} An object containing success status and a message.
 */
function addSalesOrder(salesData) {
  // Check for duplicate PO number
  if (isDuplicateSalesPO(salesData.customerPoId)) {
    return { error: true, message: `Customer PO Number "${salesData.customerPoId}" already exists. Please use a unique PO number.` };
  }
  
  try {
    const salesSheet = getSheetByName(SALES_SHEET_NAME);
    ensureSalesMaterialHeadersExist(salesSheet, SALES_COL_FIRST_MATERIAL);

    // --- 1. Internal Sale ID Generation (Using Counters Sheet) ---
    const countersSheet = getSheetByName(COUNTERS_SHEET_NAME);
    const countersData = countersSheet.getDataRange().getValues();
    
    // Find Sales row in Counters sheet
    let salesRowIndex = -1;
    let currentCounter = 0;
    for (let i = 0; i < countersData.length; i++) { // Start from row 2 (skip header)
      if (countersData[i][COUNTERS_COL_TYPE] && countersData[i][COUNTERS_COL_TYPE].toString().trim() === 'Sales') {
        salesRowIndex = i;
        currentCounter = parseInt(countersData[i][COUNTERS_COL_COUNTER]) || 0;
        break;
      }
    }
    
    if (salesRowIndex === -1) {
      return { error: true, message: "Sales counter not found in Counters sheet." };
    }
    
    // Generate new ID
    const nextSaleNumber = currentCounter + 1;
    const newInternalSaleId = `S${('0000' + nextSaleNumber).slice(-4)}`;
    
    // Update counter in Counters sheet
    countersSheet.getRange(salesRowIndex + 1, COUNTERS_COL_COUNTER + 1).setValue(nextSaleNumber);
    // --- End Sale ID Generation ---
    // --- 2. Prepare Data Row ---
    const finalHeaders = salesSheet.getRange(1, SALES_COL_FIRST_MATERIAL + 1, 1, salesSheet.getLastColumn() - SALES_COL_FIRST_MATERIAL).getValues()[0].map(h => h.toString().trim());
    const numColumns = salesSheet.getLastColumn();
    const newRow = new Array(numColumns).fill('');

    // Insert fixed fields
    newRow[SALES_COL_INTERNAL_ID] = newInternalSaleId;
    newRow[SALES_COL_PO_NUMBER] = salesData.customerPoId;
    newRow[SALES_COL_DATE_PO] = salesData.poDate ? new Date(salesData.poDate) : '';
    newRow[SALES_COL_APPOINTMENT_DATE] = salesData.appointmentDate ? new Date(salesData.appointmentDate) : '';
    newRow[SALES_COL_INVOICE] = salesData.invoiceNumber;
    newRow[SALES_COL_VENDOR_ID] = salesData.vendorId || '';
    newRow[SALES_COL_DELIVERY_ID] = salesData.deliveryId || '';
    newRow[SALES_COL_AMOUNT] = salesData.amount || 0;
    newRow[SALES_COL_GST] = salesData.gst || 0;
    newRow[SALES_COL_TOTAL] = salesData.total || 0;

    // --- 3. Handle file uploads ---
    let poLink = '', invLink = '', ewayLink = '';
    if (salesData.filesMeta && salesData.filesMeta.length > 0) {
      // safe SALE_FOLDER fallback
      const saleFolderName = (typeof SALE_FOLDER !== 'undefined' && SALE_FOLDER) ? SALE_FOLDER : 'SalesInternal';

      // Folder path: SALES folder -> YYYY/MM/DD/<INTERNAL_ID>
      const dateObj = salesData.poDate ? new Date(salesData.poDate) : new Date();
      const yyyy = dateObj.getFullYear();
      const mm = ('0' + (dateObj.getMonth() + 1)).slice(-2);
      const dd = ('0' + dateObj.getDate()).slice(-2);

      // Use Internal ID as folder name (no sanitization needed for S001 format)
      const subPath = `${saleFolderName}/${yyyy}/${mm}/${dd}/${newInternalSaleId}`;
      const urls = uploadFilesToDrive(PARENT_FOLDER_ID, subPath, salesData.filesMeta, salesData.customerPoId);
      // Order: [PO, Invoice, EWay]
      poLink = urls[0] || '';
      invLink = urls[1] || '';
      ewayLink = urls[2] || '';
    }
    newRow[SALES_COL_PO_LINK] = poLink;
    newRow[SALES_COL_INV_LINK] = invLink;
    newRow[SALES_COL_EWAY_LINK] = ewayLink;
    newRow[SALES_COL_COMMENTS] = salesData.comments || '';

    // --- 4. PO Material Quantities ---
    // Headers format: [Fixed cols...] [Mat1, Mat2, Mat3...] [Mat1, Mat2, Mat3...]
    // First set = PO quantities, Second set = Dispatch quantities
    const materialCount = Math.floor(finalHeaders.length / 2);
    const poMatHeaders = finalHeaders.slice(0, materialCount);
    const dispatchMatHeaders = finalHeaders.slice(materialCount);
    
    // Insert PO quantities (first set of material columns)
    poMatHeaders.forEach((matId, index) => {
      const headerStr = matId.toString().trim();
      if (!headerStr) return;
      const quantity = salesData.materials[headerStr] || 0;
      if (quantity > 0) {
        newRow[SALES_COL_FIRST_MATERIAL + index] = quantity;
      }
    });
    
    // Insert Dispatch quantities (second set of material columns) if provided
    if (salesData.dispatchMaterials && Object.keys(salesData.dispatchMaterials).length > 0) {
      dispatchMatHeaders.forEach((matId, index) => {
        const headerStr = matId.toString().trim();
        if (!headerStr) return;
        const quantity = salesData.dispatchMaterials[headerStr] || 0;
        if (quantity > 0) {
          newRow[SALES_COL_FIRST_MATERIAL + materialCount + index] = quantity;
        }
      });
    }

    salesSheet.appendRow(newRow);

    return { success: true, message: `Sales Order ${salesData.customerPoId} recorded. Internal ID: ${newInternalSaleId}.` };
  } catch (e) {
    Logger.log("Error in addSalesOrder: " + e.toString());
    return { error: true, message: "Server error during Sales Order submission: " + e.toString() };
  }
}
/**
 * Updates an existing Sales Order in the Sales sheet.
 * Only Appointment Date and Invoice Number can be updated.
 * @param {object} updateData Object containing internalId, appointmentDate, and invoiceNumber.
 * @returns {object} An object containing success status and a message.
 */
function updateSalesOrder(updateData) {
  try {
    try {
      Logger.log('updateSalesOrder - updateData: ' + JSON.stringify(updateData));
    } catch (e) {
      Logger.log('updateSalesOrder - failed to stringify updateData: ' + e.toString());
      Logger.log('updateSalesOrder - raw updateData: ' + updateData);
    }
    const salesSheet = getSheetByName(SALES_SHEET_NAME);
    const internalId = updateData.internalId.toString().trim();

    if (!internalId) {
        return { error: true, message: "Missing Internal Sales Order ID for update." };
    }

    // 1. Find the row index using the Internal ID (Column A, index 0)
    const dataRange = salesSheet.getDataRange();
    const allValues = dataRange.getValues(); // Including headers

    // Search for the Internal ID in Column A (index 0)
    let rowIndexToUpdate = -1;
    for (let i = 1; i < allValues.length; i++) { // Start from row 2 (index 1)
        if (allValues[i][SALES_COL_INTERNAL_ID].toString().trim() === internalId) {
            rowIndexToUpdate = i; // 0-based index of the row array (which corresponds to row number i+1 in sheet)
            break;
        }
    }

    if (rowIndexToUpdate === -1) {
        return { error: true, message: `Sales Order with Internal ID ${internalId} not found.` };
    }

    // 2. Prepare the new values for the columns to be updated
    const rowNumberInSheet = rowIndexToUpdate + 1; // 1-based index

    // Appointment Date (Column D, index 3)
    const newAppointmentDate = updateData.appointmentDate ? new Date(updateData.appointmentDate) : '';
    salesSheet.getRange(rowNumberInSheet, SALES_COL_APPOINTMENT_DATE + 1).setValue(newAppointmentDate); 

    // Invoice Number (Column E, index 4)
    const newInvoiceNumber = updateData.invoiceNumber || '';
    salesSheet.getRange(rowNumberInSheet, SALES_COL_INVOICE + 1).setValue(newInvoiceNumber);

    // Vendor ID (now mandatory)
    if (typeof updateData.vendorId !== 'undefined' && updateData.vendorId !== null) {
      const newVendorId = updateData.vendorId || '';
      salesSheet.getRange(rowNumberInSheet, SALES_COL_VENDOR_ID + 1).setValue(newVendorId);
    }

    // Amount (now mandatory)
    if (typeof updateData.amount !== 'undefined' && updateData.amount !== null) {
      salesSheet.getRange(rowNumberInSheet, SALES_COL_AMOUNT + 1).setValue(updateData.amount);
    }

    // GST (now mandatory)
    if (typeof updateData.gst !== 'undefined' && updateData.gst !== null) {
      salesSheet.getRange(rowNumberInSheet, SALES_COL_GST + 1).setValue(updateData.gst);
    }

    // Total (calculated field)
    if (typeof updateData.total !== 'undefined' && updateData.total !== null) {
      salesSheet.getRange(rowNumberInSheet, SALES_COL_TOTAL + 1).setValue(updateData.total);
    }

    // Delivery ID (optional)
    if (typeof updateData.deliveryId !== 'undefined' && updateData.deliveryId !== null) {
      const newDeliveryId = updateData.deliveryId || '';
      salesSheet.getRange(rowNumberInSheet, SALES_COL_DELIVERY_ID + 1).setValue(newDeliveryId);
    }

    // Comments (optional)
    if (typeof updateData.comments !== 'undefined' && updateData.comments !== null) {
      salesSheet.getRange(rowNumberInSheet, SALES_COL_COMMENTS + 1).setValue(updateData.comments);
    }

    // Handle file uploads (PO, Invoice, EWay) if provided in updateData.filesMeta
    if (updateData.filesMeta && updateData.filesMeta.length > 0) {
      // safe SALE_FOLDER fallback
      const saleFolderName = (typeof SALE_FOLDER !== 'undefined' && SALE_FOLDER) ? SALE_FOLDER : 'SalesInternal';

      // Use PO date from updateData if available, otherwise get it from existing row, fallback to today
      let dateObj;
      if (updateData.poDate) {
        dateObj = new Date(updateData.poDate);
      } else if (allValues[rowIndexToUpdate][SALES_COL_DATE_PO]) {
        dateObj = new Date(allValues[rowIndexToUpdate][SALES_COL_DATE_PO]);
      } else {
        dateObj = new Date();
      }
      const yyyy = dateObj.getFullYear();
      const mm = ('0' + (dateObj.getMonth() + 1)).slice(-2);
      const dd = ('0' + dateObj.getDate()).slice(-2);

      // Use Internal ID as folder name
      const subPath = `${saleFolderName}/${yyyy}/${mm}/${dd}/${internalId}`;

      // Get existing PO Number for file naming
      const existingPoNumber = allValues[rowIndexToUpdate][SALES_COL_PO_NUMBER].toString().trim();

      // uploadFilesToDrive returns an array of urls in order; caller should pass slots (null allowed)
      const urls = uploadFilesToDrive(PARENT_FOLDER_ID, subPath, updateData.filesMeta, existingPoNumber);

      // Only overwrite cells when a new non-empty URL is returned; otherwise preserve existing links
      const newPoLink = (urls[0] && urls[0].toString().trim()) ? urls[0] : null;
      const newInvLink = (urls[1] && urls[1].toString().trim()) ? urls[1] : null;
      const newEwayLink = (urls[2] && urls[2].toString().trim()) ? urls[2] : null;

      if (newPoLink !== null) {
        salesSheet.getRange(rowNumberInSheet, SALES_COL_PO_LINK + 1).setValue(newPoLink);
      }
      if (newInvLink !== null) {
        salesSheet.getRange(rowNumberInSheet, SALES_COL_INV_LINK + 1).setValue(newInvLink);
      }
      if (newEwayLink !== null) {
        salesSheet.getRange(rowNumberInSheet, SALES_COL_EWAY_LINK + 1).setValue(newEwayLink);
      }
    }

    // --- Update Dispatch Material Quantities ---
    if (updateData.dispatchMaterials && Object.keys(updateData.dispatchMaterials).length > 0) {
      ensureSalesMaterialHeadersExist(salesSheet, SALES_COL_FIRST_MATERIAL);
      
      const finalHeaders = salesSheet.getRange(1, SALES_COL_FIRST_MATERIAL + 1, 1, salesSheet.getLastColumn() - SALES_COL_FIRST_MATERIAL).getValues()[0].map(h => h.toString().trim());
      
      // Headers format: [Mat1, Mat2, Mat3...] [Mat1, Mat2, Mat3...]
      // First set = PO quantities, Second set = Dispatch quantities
      const materialCount = Math.floor(finalHeaders.length / 2);
      const dispatchMatHeaders = finalHeaders.slice(materialCount);
      
      // Update dispatch quantities (second set of material columns)
      dispatchMatHeaders.forEach((matId, index) => {
        const headerStr = matId.toString().trim();
        if (!headerStr) return;
        const quantity = updateData.dispatchMaterials[headerStr] || 0;
        salesSheet.getRange(rowNumberInSheet, SALES_COL_FIRST_MATERIAL + materialCount + index + 1).setValue(quantity);
      });
    }

    return { success: true, message: `Sales Order ${internalId} updated successfully.` };

  } catch (e) {
    Logger.log("Error in updateSalesOrder: " + e.toString());
    return { error: true, message: "Server error during Sales Order update: " + e.toString() };
  }
}

/**
 * Fetches and processes Purchase data.
 * Returns purchase orders with internal ID, vendor PO ID, dates, invoice, and material details.
 * @returns {Array<object>|object} An array of structured purchase PO data or an error object.
 */
function getPurchaseData() {
  try {
    // 1. Get Material Map (ProductID to ProductName)
    const productMap = getMaterialMap();

    // 2. Get Purchase Data
    const poSheet = getSheetByName(PO_SHEET_NAME);
    const poData = poSheet.getDataRange().getValues();
    if (poData.length <= 1) return []; // Only headers or empty

    // Headers: Internal PO ID, Vendor PO ID, PO Date, Despatch Date, Invoice, MatId1, MatId2, ...
    const headers = poData.shift();

    // MatIds start from the column index defined by PO_COL_FIRST_MATERIAL
    const matIdHeaders = headers.slice(PO_COL_FIRST_MATERIAL);

    const purchasePOs = [];

    for (const row of poData) {
      // Check for required fields
      if (!row[PO_COL_ID]) continue;

      // Use constants for fixed columns
      const internalId = row[P_COL_INTERNAL_ID]; // Column A
      const vendorPoId = row[PO_COL_ID]; // Column B
      const invoiceNumber = row[PO_COL_INVOICE]; // Column E
      const supplierId = row[PO_COL_SUPPLIER_ID] || ''; // Column F

      // PO Date (Column C)
      const poDate = row[PO_COL_DATE] instanceof Date ? Utilities.formatDate(row[PO_COL_DATE], Session.getScriptTimeZone(), "MM/dd/yyyy") : row[PO_COL_DATE];
      const rawPoDate = row[PO_COL_DATE] instanceof Date ? Utilities.formatDate(row[PO_COL_DATE], Session.getScriptTimeZone(), "yyyy-MM-dd") : '';

      // Despatch Date (Column D)
      const despatchDate = row[PO_COL_DESPATCH_DATE] instanceof Date ? Utilities.formatDate(row[PO_COL_DESPATCH_DATE], Session.getScriptTimeZone(), "MM/dd/yyyy") : row[PO_COL_DESPATCH_DATE];
      const rawDespatchDate = row[PO_COL_DESPATCH_DATE] instanceof Date ? Utilities.formatDate(row[PO_COL_DESPATCH_DATE], Session.getScriptTimeZone(), "yyyy-MM-dd") : '';

      // File links
      const poLink = row[PO_COL_PO_LINK] || '';
      const invLink = row[PO_COL_INV_LINK] || '';
      const ewayLink = row[PO_COL_EWAY_LINK] || '';

      let displayItemDetails = []; // For modal display (string array)
      let rawItemDetails = {};     // For potential edit form (map of matId: quantity)

      // Iterate through the quantity columns
      for (let i = 0; i < matIdHeaders.length; i++) {
        const matId = matIdHeaders[i].toString().trim();

        // Calculate quantity column index using constant
        const quantity = row[i + PO_COL_FIRST_MATERIAL];
        const numericQuantity = typeof quantity === 'number' ? quantity : (parseInt(quantity) || 0);

        if (numericQuantity > 0) {
          const productName = productMap.get(matId) || `Unknown Product (ID: ${matId})`;
          displayItemDetails.push(`${productName}: ${numericQuantity}`);
          rawItemDetails[matId] = numericQuantity;
        }
      }

      purchasePOs.push({
        internalId: internalId,
        vendorPoId: vendorPoId,
        poDate: poDate, // Display format
        rawPoDate: rawPoDate, // Form input format
        despatchDate: despatchDate, // Display format
        rawDespatchDate: rawDespatchDate, // Form input format
        invoiceNumber: invoiceNumber,
        supplierId: supplierId,
        poLink: poLink,
        invLink: invLink,
        ewayLink: ewayLink,
        displayItemDetails: displayItemDetails.join('\n'),
        rawItemDetails: rawItemDetails
      });
    }

    return purchasePOs;

  } catch (e) {
    Logger.log("Error in getPurchaseData: " + e.toString());
    return { error: true, message: "Failed to load purchase data: " + e.toString() };
  }
}

/**
 * Fetches and processes Sales data, cross-referencing with Material sheet.
 * This function now returns the internalId, raw dates, and rawItemDetails (map)
 * for use in the Edit form.
 * @returns {Array<object>|object} An array of structured sales PO data or an error object.
 */
function getSalesData() {
  try {
    // 1. Get Material Map (ProductID to ProductName)
    const productMap = getMaterialMap(); 

    // 2. Get Vendor Map (VendorID to VendorName)
    const vendorSheet = getSheetByName(VENDOR_SHEET_NAME);
    const vendorMap = new Map();
    if (vendorSheet) {
      const vendorData = vendorSheet.getDataRange().getValues();
      for (let i = 1; i < vendorData.length; i++) {
        if (vendorData[i][VENDOR_COL_ID]) {
          vendorMap.set(vendorData[i][VENDOR_COL_ID].toString().trim(), vendorData[i][VENDOR_COL_NAME]);
        }
      }
    }

    // 3. Get Sales Data
    const salesSheet = getSheetByName(SALES_SHEET_NAME);
    const salesData = salesSheet.getDataRange().getValues();
    if (salesData.length <= 1) return []; // Only headers or empty

    // Headers: SaleID, Sale PO Number, Date of PO, Appointment Date, Invoice, MatId1, MatId2, ...
    const headers = salesData.shift();

    // Material IDs start from SALES_COL_FIRST_MATERIAL
    // Structure: [Fixed cols...] [Mat1, Mat2, Mat3...] [Mat1, Mat2, Mat3...]
    // First set = PO quantities, Second set = Dispatch quantities
    const allMaterialHeaders = headers.slice(SALES_COL_FIRST_MATERIAL);
    
    // Determine how many materials we have (half the material columns)
    const materialCount = Math.floor(allMaterialHeaders.length / 2);
    const poMatHeaders = allMaterialHeaders.slice(0, materialCount);
    const dispatchMatHeaders = allMaterialHeaders.slice(materialCount);

    const salesPOs = [];

    for (const row of salesData) {
        // Check for required fields using constant
        if (!row[SALES_COL_PO_NUMBER]) continue;

        // Use constants for fixed columns (updated to reflect new indices)
        const internalId = row[SALES_COL_INTERNAL_ID]; // Column A (Index 0)
        const poNumber = row[SALES_COL_PO_NUMBER];
        const invoiceNumber = row[SALES_COL_INVOICE]; // Column E (Index 4)
        const vendorId = row[SALES_COL_VENDOR_ID] ? row[SALES_COL_VENDOR_ID].toString().trim() : '';
        const vendorName = vendorId ? (vendorMap.get(vendorId) || vendorId) : '';

        // Date of PO (Column C, index 2)
        const dateOfPO = row[SALES_COL_DATE_PO] instanceof Date ? Utilities.formatDate(row[SALES_COL_DATE_PO], Session.getScriptTimeZone(), "MM/dd/yyyy") : row[SALES_COL_DATE_PO];
        const rawDateOfPO = row[SALES_COL_DATE_PO] instanceof Date ? Utilities.formatDate(row[SALES_COL_DATE_PO], Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
        
        // Appointment Date (Column D, index 3)
        const appointmentDate = row[SALES_COL_APPOINTMENT_DATE] instanceof Date ? Utilities.formatDate(row[SALES_COL_APPOINTMENT_DATE], Session.getScriptTimeZone(), "MM/dd/yyyy") : row[SALES_COL_APPOINTMENT_DATE];
        const rawAppointmentDate = row[SALES_COL_APPOINTMENT_DATE] instanceof Date ? Utilities.formatDate(row[SALES_COL_APPOINTMENT_DATE], Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
        

        let displayItemDetails = []; // For modal display (string array)
        let rawItemDetails = {};     // For edit form population - PO quantities (map of matId: quantity)
        let rawDispatchDetails = {}; // For edit form population - Dispatch quantities (map of matId: quantity)

        // Process PO quantities (first set of material columns)
        for (let i = 0; i < poMatHeaders.length; i++) {
            const matId = poMatHeaders[i].toString().trim();
            if (!matId) continue;
            
            const quantity = row[SALES_COL_FIRST_MATERIAL + i]; 
            const numericQuantity = typeof quantity === 'number' ? quantity : (parseInt(quantity) || 0);

            if (numericQuantity > 0) {
                const productName = productMap.get(matId) || `Unknown Product (ID: ${matId})`;
                displayItemDetails.push(`${productName}: ${numericQuantity}`);
                rawItemDetails[matId] = numericQuantity;
            }
        }

        // Process Dispatch quantities (second set of material columns)
        for (let i = 0; i < dispatchMatHeaders.length; i++) {
            const matId = dispatchMatHeaders[i].toString().trim();
            if (!matId) continue;
            
            const quantity = row[SALES_COL_FIRST_MATERIAL + materialCount + i]; 
            const numericQuantity = typeof quantity === 'number' ? quantity : (parseInt(quantity) || 0);

            if (numericQuantity > 0) {
                rawDispatchDetails[matId] = numericQuantity;
            }
        }

        // Check if appointment date is today and validate required fields
        let validationStatus = 'normal'; // 'normal', 'complete', 'incomplete'
        if (row[SALES_COL_APPOINTMENT_DATE] instanceof Date) {
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            const appointmentDateOnly = new Date(row[SALES_COL_APPOINTMENT_DATE]);
            appointmentDateOnly.setHours(0, 0, 0, 0);
            
            if (appointmentDateOnly.getTime() === today.getTime()) {
                // Appointment date is today - check required fields
                const hasInvoice = invoiceNumber && invoiceNumber.toString().trim() !== '';
                const hasVendor = vendorId && vendorId !== '';
                const hasDelivery = row[SALES_COL_DELIVERY_ID] && row[SALES_COL_DELIVERY_ID].toString().trim() !== '';
                const hasPoLink = row[SALES_COL_PO_LINK] && row[SALES_COL_PO_LINK].toString().trim() !== '';
                const hasInvLink = row[SALES_COL_INV_LINK] && row[SALES_COL_INV_LINK].toString().trim() !== '';
                const hasEwayLink = row[SALES_COL_EWAY_LINK] && row[SALES_COL_EWAY_LINK].toString().trim() !== '';
                const hasDispatchQty = Object.keys(rawDispatchDetails).length > 0;
                
                // All required fields present = complete (green), otherwise incomplete (red)
                if (hasInvoice && hasVendor && hasDelivery && hasPoLink && hasInvLink && hasEwayLink && hasDispatchQty) {
                    validationStatus = 'complete';
                } else {
                    validationStatus = 'incomplete';
                }
            }
        }

        salesPOs.push({
            internalId: internalId, // NEW: Unique ID for updates
            poNumber: poNumber,
            dateOfPO: dateOfPO, // Display format
            rawDateOfPO: rawDateOfPO, // Form input format
            appointmentDate: appointmentDate, // Display format
            rawAppointmentDate: rawAppointmentDate, // Form input format
            invoiceNumber: invoiceNumber, // NEW: For update
            vendorId: vendorId,     // NEW: Vendor ID for form
            vendorName: vendorName, // NEW: Vendor Name for display
            deliveryId: row[SALES_COL_DELIVERY_ID] || '', // NEW: Delivery ID for form
            amount: row[SALES_COL_AMOUNT] || 0, // Amount
            gst: row[SALES_COL_GST] || 0, // GST
            total: row[SALES_COL_TOTAL] || 0, // Total
            poLink: row[SALES_COL_PO_LINK] || '',        // NEW: PO document URL
            invLink: row[SALES_COL_INV_LINK] || '',      // NEW: Invoice document URL
            ewayLink: row[SALES_COL_EWAY_LINK] || '',    // NEW: EWay document URL
            comments: row[SALES_COL_COMMENTS] || '',     // NEW: Comments
            displayItemDetails: displayItemDetails.join('\n'), // For modal button click
            rawItemDetails: rawItemDetails, // For edit form population - PO quantities
            rawDispatchDetails: rawDispatchDetails, // For edit form population - Dispatch quantities
            validationStatus: validationStatus // NEW: Validation status for color coding
        });
    }

    return salesPOs;

  } catch (e) {
    Logger.log("Error in getSalesData: " + e.toString());
    return { error: true, message: "Failed to load inventory data: " + e.toString() };
  }
}

/**
 * Utility function to fetch all materials for the client-side form.
 * @returns {Array<object>} List of material objects {id, name}.
 */
function getMaterialsForForm() {
  try {
    const materialSheet = getSheetByName(MATERIAL_SHEET_NAME);
    const materialData = materialSheet.getDataRange().getValues();
    const materials = [];

    // Skip header (row 0)
    for (let i = 1; i < materialData.length; i++) {
        const row = materialData[i];
        // Use constants for Material ID and Name
        if (row.length >= MATERIAL_COL_NAME + 1 && row[MATERIAL_COL_ID]) {
            materials.push({
                id: row[MATERIAL_COL_ID].toString().trim(),
                name: row[MATERIAL_COL_NAME].toString().trim()
            });
        }
    }
    return materials;
  } catch (e) {
    Logger.log("Error in getMaterialsForForm: " + e.toString());
    throw new Error("Failed to load materials list.");
  }
}

/**
 * Utility function to generate the product map from the Material sheet.
 * @returns {Map<string, string>} Map of Product ID to Product Name.
 */
function getMaterialMap() {
    const materialSheet = getSheetByName(MATERIAL_SHEET_NAME);
    const materialData = materialSheet.getDataRange().getValues();
    const productMap = new Map();

    // Assuming first row of Material is header, start map from second row.
    for (let i = 1; i < materialData.length; i++) {
        const row = materialData[i];
        if (row.length >= MATERIAL_COL_NAME + 1 && row[MATERIAL_COL_ID]) {
            productMap.set(row[MATERIAL_COL_ID].toString().trim(), row[MATERIAL_COL_NAME]);
        }
    }
    return productMap;
}

/**
 * Inserts a manual stock entry into the Manual sheet.
 * Mirrors the addPurchaseOrder structure: generates an internal ID (M001...), writes timestamp and material quantities.
 * Expects manualData = { materials: { "<matId>": <qty>, ... } }
 * Returns { success: true, message: '...' } or { error: true, message: '...' }.
 */
function addManualStock(manualData) {
  try {
    const manualSheet = getSheetByName(MANUAL_SHEET_NAME);

    // 0. Ensure Headers (append any new material IDs)
    ensureMaterialHeadersExist(manualSheet, MANUAL_COL_FIRST_MATERIAL);

    // --- 1. Internal Manual ID Generation (Based on Column A: MANUAL_COL_INTERNAL_ID) ---
    let nextManualNumber = 1;
    if (manualSheet.getLastRow() > 1) {
      const numRows = manualSheet.getLastRow() - 1;
      const ids = manualSheet.getRange(2, MANUAL_COL_INTERNAL_ID + 1, numRows, 1).getValues().map(row => row[0].toString().trim());
      const maxNumber = ids.reduce((max, id) => {
        const match = id.match(/^M(\d+)$/);
        if (match) {
          const currentNum = parseInt(match[1], 10);
          return Math.max(max, currentNum);
        }
        return max;
      }, 0);
      nextManualNumber = maxNumber + 1;
    }
    const newInternalManualId = `M${('000' + nextManualNumber).slice(-3)}`;
    // --- End Manual ID Generation ---

    // 2. Prepare Data Row
    // Get the final list of material headers from the sheet to match column order
    const lastCol = manualSheet.getLastColumn();
    const materialColsCount = Math.max(0, lastCol - MANUAL_COL_FIRST_MATERIAL);
    const finalHeaders = materialColsCount > 0
      ? manualSheet.getRange(1, MANUAL_COL_FIRST_MATERIAL + 1, 1, materialColsCount).getValues()[0].map(h => h.toString().trim())
      : [];

    // Create a row template large enough for all columns
    const numColumns = Math.max(lastCol, MANUAL_COL_FIRST_MATERIAL);
    const newRow = new Array(numColumns).fill('');

    // Insert fixed fields
    newRow[MANUAL_COL_INTERNAL_ID] = newInternalManualId; // Column A: Internal Manual ID
    newRow[MANUAL_COL_TIMESTAMP] = new Date();            // Column B: Timestamp

    // Insert material quantities starting from MANUAL_COL_FIRST_MATERIAL
    finalHeaders.forEach((matId, index) => {
      const quantity = manualData.materials && manualData.materials[matId] ? manualData.materials[matId] : 0;
      if (quantity > 0) {
        newRow[MANUAL_COL_FIRST_MATERIAL + index] = quantity;
      }
    });

    // 4. Append Data
    manualSheet.appendRow(newRow);

    return { success: true, message: `Manual stock recorded. Internal ID: ${newInternalManualId}.` };
  } catch (e) {
    Logger.log("Error in addManualStock: " + e.toString());
    return { error: true, message: "Server error during manual stock submission: " + e.toString() };
  }
}

/**
 * Compute stock by aggregating Purchases and Sales, read last Manual row,
 * update the DATA sheet (headers, sales row, purchases row, net row, manual last row),
 * and return an array of objects: [{ id, name, sales, purchases, net, manual }, ...]
 */
function computeAndUpdateStock() {
  try {
    // 1. Get master material list & map (in order as listed in Material sheet)
    const materialSheet = getSheetByName(MATERIAL_SHEET_NAME);
    const matData = materialSheet.getDataRange().getValues();
    const materialIds = []; // ordered
    const productMap = new Map(); // id => name

    for (let i = 1; i < matData.length; i++) {
      const row = matData[i];
      if (row && row[MATERIAL_COL_ID]) {
        const id = row[MATERIAL_COL_ID].toString().trim();
        materialIds.push(id);
        productMap.set(id, row[MATERIAL_COL_NAME] ? row[MATERIAL_COL_NAME].toString().trim() : id);
      }
    }

    // Initialize sums
    const salesSums = {};
    const purchaseSums = {};
    const manualLast = {};
    const startValues = {};
    const sentSums = {};      // Sales with appointment date in the past
    const outgoingSums = {};  // Sales with no appointment date OR future date
    const receivedSums = {};  // Purchases with despatch date in the past
    const incomingSums = {};  // Purchases with no despatch date
    materialIds.forEach(id => {
      salesSums[id] = 0;
      purchaseSums[id] = 0;
      manualLast[id] = 0;
      startValues[id] = 0;
      sentSums[id] = 0;
      outgoingSums[id] = 0;
      receivedSums[id] = 0;
      incomingSums[id] = 0;
    });

    // Helper to accumulate from a sheet given firstMaterialCol index
    function accumulateFromSheet(sheet, sheetFirstMaterialCol, targetMap) {
      if (!sheet || sheet.getLastRow() < 2) return;
      const headers = sheet.getRange(1, sheetFirstMaterialCol + 1, 1, Math.max(1, sheet.getLastColumn() - sheetFirstMaterialCol)).getValues()[0].map(h => h.toString().trim());
      const numRows = sheet.getLastRow() - 1;
      const data = sheet.getRange(2, sheetFirstMaterialCol + 1, numRows, headers.length).getValues();
      for (let r = 0; r < data.length; r++) {
        const row = data[r];
        for (let c = 0; c < headers.length; c++) {
          const matId = headers[c];
          if (!matId) continue;
          const val = row[c];
          const qty = (typeof val === 'number') ? val : (parseInt(val, 10) || 0);
          if (typeof targetMap[matId] === 'undefined') targetMap[matId] = 0;
          targetMap[matId] += qty;
        }
      }
    }

    // 2. Accumulate Purchases and categorize by date
    const poSheet = getSheetByName(PO_SHEET_NAME);
    accumulateFromSheet(poSheet, PO_COL_FIRST_MATERIAL, purchaseSums);
    
    // Categorize Purchases: Received (past despatch date) vs Incoming (no date)
    if (poSheet && poSheet.getLastRow() >= 2) {
      const today = new Date();
      today.setHours(0, 0, 0, 0); // Reset to start of day for comparison
      const poHeaders = poSheet.getRange(1, PO_COL_FIRST_MATERIAL + 1, 1, Math.max(1, poSheet.getLastColumn() - PO_COL_FIRST_MATERIAL)).getValues()[0].map(h => h.toString().trim());
      const numRows = poSheet.getLastRow() - 1;
      const poData = poSheet.getRange(2, 1, numRows, poSheet.getLastColumn()).getValues();
      
      for (let r = 0; r < poData.length; r++) {
        const row = poData[r];
        const despatchDate = row[PO_COL_DESPATCH_DATE];
        const hasDespatchDate = despatchDate && despatchDate instanceof Date;
        
        // Get material quantities for this row
        for (let c = 0; c < poHeaders.length; c++) {
          const matId = poHeaders[c];
          if (!matId) continue;
          const val = row[PO_COL_FIRST_MATERIAL + c];
          const qty = (typeof val === 'number') ? val : (parseInt(val, 10) || 0);
          
          if (qty > 0) {
            if (!hasDespatchDate) {
              // No despatch date = Incoming
              if (typeof incomingSums[matId] === 'undefined') incomingSums[matId] = 0;
              incomingSums[matId] += qty;
            } else {
              const despatchDateOnly = new Date(despatchDate);
              despatchDateOnly.setHours(0, 0, 0, 0);
              if (despatchDateOnly <= today) {
                // Today or past date = Received
                if (typeof receivedSums[matId] === 'undefined') receivedSums[matId] = 0;
                receivedSums[matId] += qty;
              } else {
                // Future date (tomorrow onwards) = Incoming
                if (typeof incomingSums[matId] === 'undefined') incomingSums[matId] = 0;
                incomingSums[matId] += qty;
              }
            }
          }
        }
      }
    }

    // 3. Accumulate Sales and categorize by date
    const salesSheet = getSheetByName(SALES_SHEET_NAME);
    accumulateFromSheet(salesSheet, SALES_COL_FIRST_MATERIAL, salesSums);
    
    // Categorize Sales: Sent (today or earlier) vs Outgoing (tomorrow onwards OR no date)
    if (salesSheet && salesSheet.getLastRow() >= 2) {
      const today = new Date();
      today.setHours(0, 0, 0, 0); // Reset to start of day for comparison
      const salesHeaders = salesSheet.getRange(1, SALES_COL_FIRST_MATERIAL + 1, 1, Math.max(1, salesSheet.getLastColumn() - SALES_COL_FIRST_MATERIAL)).getValues()[0].map(h => h.toString().trim());
      const numRows = salesSheet.getLastRow() - 1;
      const salesData = salesSheet.getRange(2, 1, numRows, salesSheet.getLastColumn()).getValues();
      
      for (let r = 0; r < salesData.length; r++) {
        const row = salesData[r];
        const appointmentDate = row[SALES_COL_APPOINTMENT_DATE];
        const hasAppointmentDate = appointmentDate && appointmentDate instanceof Date;
        
        // Get material quantities for this row
        for (let c = 0; c < salesHeaders.length; c++) {
          const matId = salesHeaders[c];
          if (!matId) continue;
          const val = row[SALES_COL_FIRST_MATERIAL + c];
          const qty = (typeof val === 'number') ? val : (parseInt(val, 10) || 0);
          
          if (qty > 0) {
            if (!hasAppointmentDate) {
              // No appointment date = Outgoing
              if (typeof outgoingSums[matId] === 'undefined') outgoingSums[matId] = 0;
              outgoingSums[matId] += qty;
            } else {
              const appointmentDateOnly = new Date(appointmentDate);
              appointmentDateOnly.setHours(0, 0, 0, 0);
              if (appointmentDateOnly <= today) {
                // Today or past date = Sent
                if (typeof sentSums[matId] === 'undefined') sentSums[matId] = 0;
                sentSums[matId] += qty;
              } else {
                // Future date (tomorrow onwards) = Outgoing
                if (typeof outgoingSums[matId] === 'undefined') outgoingSums[matId] = 0;
                outgoingSums[matId] += qty;
              }
            }
          }
        }
      }
    }

    // 4. Read last Manual entry (if any) - use header from manual sheet
    const manualSheet = getSheetByName(MANUAL_SHEET_NAME);
    if (manualSheet && manualSheet.getLastRow() >= 2) {
      const lastRowIndex = manualSheet.getLastRow();
      const manualHeaders = manualSheet.getRange(1, MANUAL_COL_FIRST_MATERIAL + 1, 1, Math.max(0, manualSheet.getLastColumn() - MANUAL_COL_FIRST_MATERIAL)).getValues()[0].map(h => h.toString().trim());
      if (manualHeaders.length > 0) {
        const lastValues = manualSheet.getRange(lastRowIndex, MANUAL_COL_FIRST_MATERIAL + 1, 1, manualHeaders.length).getValues()[0];
        for (let i = 0; i < manualHeaders.length; i++) {
          const matId = manualHeaders[i];
          if (!matId) continue;
          const val = lastValues[i];
          const qty = (typeof val === 'number') ? val : (parseInt(val, 10) || 0);
          manualLast[matId] = qty;
        }
      }
    }

    // 4b. Read Start values from Data sheet row 8 (if exists)
    const dataSheet = getSheetByName(DATA_SHEET_NAME);
    if (dataSheet && dataSheet.getLastRow() >= 8) {
      const startHeaders = dataSheet.getRange(1, DATA_COL_FIRST_MATERIAL + 1, 1, Math.max(0, dataSheet.getLastColumn() - DATA_COL_FIRST_MATERIAL)).getValues()[0].map(h => h.toString().trim());
      if (startHeaders.length > 0) {
        const startRowValues = dataSheet.getRange(8, DATA_COL_FIRST_MATERIAL + 1, 1, startHeaders.length).getValues()[0];
        for (let i = 0; i < startHeaders.length; i++) {
          const matId = startHeaders[i];
          if (!matId) continue;
          const val = startRowValues[i];
          const qty = (typeof val === 'number') ? val : (parseInt(val, 10) || 0);
          startValues[matId] = qty;
        }
      }
    }

    // 5. Build per-material result list (use materialIds master order)
    const result = materialIds.map(id => {
      const name = productMap.get(id) || id;
      const sent = sentSums[id] || 0;
      const outgoing = outgoingSums[id] || 0;
      const received = receivedSums[id] || 0;
      const incoming = incomingSums[id] || 0;
      const start = startValues[id] || 0;
      const net = start + received - sent; // Net = start + received - sent
      const manual = manualLast[id] || 0;
      return { id, name, sent, outgoing, received, incoming, net, manual };
    });

    // 6. Update Data sheet
    // Note: dataSheet already retrieved in step 4b

    // Ensure header row: first cell 'Metric', materials across starting from DATA_COL_FIRST_MATERIAL (column B)
    dataSheet.getRange(1, 1).setValue('Metric');
    if (materialIds.length > 0) {
      dataSheet.getRange(1, DATA_COL_FIRST_MATERIAL + 1, 1, materialIds.length).setValues([materialIds]);
    }

    // Prepare rows: Sent (row 2), Outgoing (row 3), Received (row 4), Incoming (row 5), Net (row 6), Manual (row 7)
    // Start is stored in Data sheet row 8 and is manually entered (never overwritten)
    const sentRow = ['Sent'].concat(materialIds.map(id => sentSums[id] || 0));
    const outgoingRow = ['Outgoing'].concat(materialIds.map(id => outgoingSums[id] || 0));
    const receivedRow = ['Received'].concat(materialIds.map(id => receivedSums[id] || 0));
    const incomingRow = ['Incoming'].concat(materialIds.map(id => incomingSums[id] || 0));
    const netRow = ['Net'].concat(materialIds.map(id => (startValues[id] || 0) + (receivedSums[id] || 0) - (sentSums[id] || 0)));
    const manualRow = ['Manual (last)'].concat(materialIds.map(id => manualLast[id] || 0));

    // Write rows (ensure sheet has enough columns)
    const totalColsNeeded = DATA_COL_FIRST_MATERIAL + materialIds.length + 1; // +1 for first label column
    if (dataSheet.getLastColumn() < totalColsNeeded) {
      // extend header row if necessary
      dataSheet.insertColumnsAfter(dataSheet.getLastColumn(), totalColsNeeded - dataSheet.getLastColumn());
    }

    // Set rows (rows are 1-based) - reordered to match new requirement
    dataSheet.getRange(2, 1, 1, 1 + materialIds.length).setValues([sentRow]);
    dataSheet.getRange(3, 1, 1, 1 + materialIds.length).setValues([outgoingRow]);
    dataSheet.getRange(4, 1, 1, 1 + materialIds.length).setValues([receivedRow]);
    dataSheet.getRange(5, 1, 1, 1 + materialIds.length).setValues([incomingRow]);
    dataSheet.getRange(6, 1, 1, 1 + materialIds.length).setValues([netRow]);
    dataSheet.getRange(7, 1, 1, 1 + materialIds.length).setValues([manualRow]);
    // Start row (row 8) is manually entered and should NOT be overwritten by this function

    return result;

  } catch (e) {
    Logger.log("Error in computeAndUpdateStock: " + e.toString());
    return { error: true, message: "Failed to compute/update stock: " + e.toString() };
  }
}

/**
 * Wrapper exposed to client-side for retrieving the current stock data.
 * Called via google.script.run.getCurrentStockData(...)
 */
function getCurrentStockData() {
  return computeAndUpdateStock();
}

/**
 * Returns a list of vendors from the Material sheet or a separate Vendor sheet.
 * For now, assumes a 'Vendor' sheet with columns: [Vendor ID, Vendor Name]
 */
function getVendorsForForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vendorSheet = ss.getSheetByName(VENDOR_SHEET_NAME);
  if (!vendorSheet) return [];
  const data = vendorSheet.getDataRange().getValues();
  const vendors = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][VENDOR_COL_ID]) vendors.push({
      id: data[i][VENDOR_COL_ID],
      name: data[i][VENDOR_COL_NAME]
    });
  }
  return vendors;
}

/**
 * Handles file uploads and returns Drive URLs.
 * Robust: falls back to root if getFolderById fails, preserves input order and length.
 * Files are renamed uniformly as: PO_ID_PO, PO_ID_Inv, PO_ID_Eway
 * @param {string} folderId Parent folder ID
 * @param {string} subPath Subfolder path
 * @param {Array} files Array of file objects with bytes, mimeType, name
 * @param {string} poId PO ID for standardized file naming
 */
function uploadFilesToDrive(folderId, subPath, files, poId) {
  // Resolve base folder with fallback to root on error
  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    Logger.log('uploadFilesToDrive: getFolderById failed for id=' + folderId + ' -> ' + e.toString());
    // Fallback: use root folder to avoid hard failure (log for user to check permissions/id)
    try {
      folder = DriveApp.getRootFolder();
      Logger.log('uploadFilesToDrive: falling back to root folder.');
    } catch (e2) {
      Logger.log('uploadFilesToDrive: failed to access Drive root folder -> ' + e2.toString());
      // Return placeholders matching files length (if any) to avoid throwing in callers
      const emptyUrls = (files && files.length) ? files.map(() => '') : [];
      return emptyUrls;
    }
  }

  const parts = subPath.split('/').filter(Boolean);
  for (const part of parts) {
    let found = false;
    const folders = folder.getFoldersByName(part);
    if (folders.hasNext()) {
      folder = folders.next();
      found = true;
    }
    if (!found) {
      folder = folder.createFolder(part);
    }
  }

  // Prepare result array with same length as input files (preserve order)
  const urls = Array.isArray(files) ? new Array(files.length).fill('') : [];

  if (!files || files.length === 0) return urls;

  // Define standard file names based on order: [PO, Invoice, EWay]
  const standardNames = ['PO', 'Inv', 'Eway'];
  const safePoId = (poId || 'UNKNOWN').toString().trim().replace(/[\\\/:*?"<>|]/g, '_');

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    if (!file) {
      urls[i] = ''; // preserve position for missing files
      continue;
    }
    try {
      // Determine file extension from original filename or mimeType
      let fileExt = '';
      if (file.name) {
        const dotIndex = file.name.lastIndexOf('.');
        if (dotIndex > -1) fileExt = file.name.substring(dotIndex);
      }
      if (!fileExt && file.mimeType) {
        // Fallback: guess extension from mimeType
        if (file.mimeType.includes('pdf')) fileExt = '.pdf';
        else if (file.mimeType.includes('image')) fileExt = '.jpg';
      }

      // Generate standardized filename: PO_ID_PO, PO_ID_Inv, PO_ID_Eway
      const standardName = `${safePoId}_${standardNames[i] || 'File'}${fileExt}`;

      let blob;
      if (typeof file.bytes === 'string') {
        // file.bytes is base64 string
        const decoded = Utilities.base64Decode(file.bytes);
        blob = Utilities.newBlob(decoded, file.mimeType || 'application/octet-stream', standardName);
      } else {
        // assume raw bytes (array) or already blob-compatible
        blob = Utilities.newBlob(file.bytes, file.mimeType || 'application/octet-stream', standardName);
      }
      const driveFile = folder.createFile(blob);
      
      // Set sharing to "anyone with the link can view"
      try {
        driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (shareError) {
        Logger.log(`uploadFilesToDrive: failed to set sharing for file index ${i} name=${file.name} -> ${shareError.toString()}`);
        // Continue even if sharing fails - file is still uploaded
      }
      
      urls[i] = driveFile.getUrl();
    } catch (e) {
      Logger.log(`uploadFilesToDrive: failed to upload file index ${i} name=${file && file.name} -> ${e.toString()}`);
      urls[i] = ''; // on error keep placeholder to preserve order
    }
  }
  return urls;
}

/**
 * Returns a list of deliveries from the Delivery sheet.
 * For now, assumes a 'Delivery' sheet with columns: [Delivery ID, Delivery Name]
 */
function getDeliveriesForForm() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const deliverySheet = ss.getSheetByName(DELIVERY_SHEET_NAME);
    if (!deliverySheet) return [];
    const data = deliverySheet.getDataRange().getValues();
    const deliveries = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][DELIVERY_COL_ID]) deliveries.push({
        id: data[i][DELIVERY_COL_ID],
        name: data[i][DELIVERY_COL_NAME]
      });
    }
    return deliveries;
  } catch (e) {
    Logger.log("Error in getDeliveriesForForm: " + e.toString());
    return [];
  }
}

/**
 * Returns a list of suppliers from the Supplier sheet.
 * Assumes a 'Supplier' sheet with columns: [Supplier ID, Supplier Name]
 */
function getSuppliersForForm() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const supplierSheet = ss.getSheetByName(SUPPLIER_SHEET_NAME);
    if (!supplierSheet) return [];
    const data = supplierSheet.getDataRange().getValues();
    const suppliers = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][SUPPLIER_COL_ID]) suppliers.push({
        id: data[i][SUPPLIER_COL_ID],
        name: data[i][SUPPLIER_COL_NAME]
      });
    }
    return suppliers;
  } catch (e) {
    Logger.log("Error in getSuppliersForForm: " + e.toString());
    return [];
  }
}

/**
 * Sends vehicle/PO details email for a Sales Order identified by internalId.
 * - Validates presence of PO Link, Delivery ID and Vendor ID; returns error if any missing.
 * - Makes the PO file (if a Drive URL) shareable "anyone with the link" before sending.
 * - Looks up Delivery Vehicle from Delivery sheet (third column) and Vendor name from Vendor sheet.
 * @param {string} internalId Sales internal ID (e.g. "S001")
 * @returns {object} { success: true, message: '...' } or { error: true, message: '...' }
 */
function sendVehicleEmail(internalId) {
  try {
    if (!internalId) return { error: true, message: 'Missing internalId' };

    const salesSheet = getSheetByName(SALES_SHEET_NAME);
    const data = salesSheet.getDataRange().getValues(); // includes header

    // find row (0-based index within array)
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      const val = data[i][SALES_COL_INTERNAL_ID];
      if (val && val.toString().trim() === internalId.toString().trim()) {
        rowIndex = i;
        break;
      }
    }
    if (rowIndex === -1) return { error: true, message: `Sales order ${internalId} not found.` };

    const row = data[rowIndex];
    const poNumber = row[SALES_COL_PO_NUMBER] || '';
    const poLink = row[SALES_COL_PO_LINK] ? row[SALES_COL_PO_LINK].toString().trim() : '';
    const deliveryId = row[SALES_COL_DELIVERY_ID] ? row[SALES_COL_DELIVERY_ID].toString().trim() : '';
    const vendorId = row[SALES_COL_VENDOR_ID] ? row[SALES_COL_VENDOR_ID].toString().trim() : '';

    // Validate required fields before proceeding
    if (!poLink) {
      return { error: true, message: 'Mail not sent: PO not uploaded' };
    }
    if (!deliveryId) {
      return { error: true, message: 'Mail not sent: Delivery Not selected' };
    }
    if (!vendorId) {
      return { error: true, message: 'Mail not sent: Vendor not selected' };
    }

    // Ensure PO file (if Drive url) is shareable by anyone with link
    let safePoLink = poLink;
    if (poLink && typeof poLink === 'string') {
      const idMatch = poLink.match(/\/d\/([a-zA-Z0-9_-]+)/) || poLink.match(/id=([a-zA-Z0-9_-]+)/);
      const fileId = idMatch ? idMatch[1] : null;
      if (fileId) {
        try {
          const file = DriveApp.getFileById(fileId);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          safePoLink = file.getUrl();
        } catch (e) {
          Logger.log('sendVehicleEmail: failed to set sharing for fileId=' + fileId + ' -> ' + e.toString());
          // proceed with original link if sharing fails
        }
      }
    }

    // Lookup Delivery Vehicle from Delivery sheet
    let deliveryVehicle = '';
    if (deliveryId) {
      const deliverySheet = getSheetByName(DELIVERY_SHEET_NAME);
      if (deliverySheet) {
        const dData = deliverySheet.getDataRange().getValues();
        for (let i = 1; i < dData.length; i++) {
          if (dData[i][DELIVERY_COL_ID] && dData[i][DELIVERY_COL_ID].toString().trim() === deliveryId) {
            deliveryVehicle = dData[i][DELIVERY_COL_VEHICLE] || '';
            break;
          }
        }
      }
    }

    // Lookup Vendor details from Vendor sheet
    let vendorName = '';
    if (vendorId) {
      const vendorSheet = getSheetByName(VENDOR_SHEET_NAME);
      if (vendorSheet) {
        const vData = vendorSheet.getDataRange().getValues();
        for (let i = 1; i < vData.length; i++) {
          if (vData[i][VENDOR_COL_ID] && vData[i][VENDOR_COL_ID].toString().trim() === vendorId) {
            vendorName = vData[i][VENDOR_COL_NAME] || '';
            break;
          }
        }
      }
    }

    // Compose email - parse semicolon-separated recipients
    const emailList = (typeof NOTIFY_EMAIL !== 'undefined' && NOTIFY_EMAIL) ? NOTIFY_EMAIL : '';
    if (!emailList) return { error: true, message: 'Notification email not configured on server.' };

    const emails = emailList.split(';').map(e => e.trim()).filter(e => e);
    if (emails.length === 0) return { error: true, message: 'No valid email addresses configured.' };

    const toEmail = emails[0]; // First email as recipient
    const ccEmails = emails.slice(1).join(','); // Rest as CC (comma-separated for MailApp)

    const subject = `Details for Invoice - PO ${poNumber || 'N/A'}`;
    let body = `PO Number: ${poNumber || 'N/A'}\n\nLink to PO: ${safePoLink || 'N/A'}\n\nDelivery Vehicle: ${deliveryVehicle || 'N/A'}\n\nVendor Name: ${vendorName || 'N/A'}`;

    // Send email with first address as TO and others as CC
    const emailOptions = { to: toEmail, subject: subject, body: body };
    if (ccEmails) emailOptions.cc = ccEmails;
    MailApp.sendEmail(emailOptions);

    return { success: true, message: `Notification sent to ${toEmail}${ccEmails ? ' (CC: ' + ccEmails + ')' : ''}` };
  } catch (e) {
    Logger.log('Error in sendVehicleEmail: ' + e.toString());
    return { error: true, message: 'Failed to send vehicle email: ' + e.toString() };
  }
}

