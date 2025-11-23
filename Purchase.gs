// Note: All constants are now defined in Constants.gs and are globally available here.
// getSheetByName is defined in Code.gs and is globally available.

/**
 * Ensures that the header row of a given sheet contains columns for all materials
 * listed in the Material sheet, starting at the specified column index.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to update (Purchase or Sales).
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
 * Validates and inserts a new Purchase Order into the Purchase sheet.
 * @param {object} poData The purchase order data from the form.
 * @returns {object} An object containing success status and a message.
 */
function addPurchaseOrder(poData) {
  try {
    const poSheet = getSheetByName(PO_SHEET_NAME);
    
    // 0. Ensure Headers
    ensureMaterialHeadersExist(poSheet, PO_COL_FIRST_MATERIAL);
    
    // --- 1. Internal PO ID Generation (Based on Column A: P_COL_INTERNAL_ID) ---
    let nextPoNumber = 1;
    if (poSheet.getLastRow() > 1) {
        const numRows = poSheet.getLastRow() - 1;
        const poIds = poSheet.getRange(2, P_COL_INTERNAL_ID + 1, numRows, 1).getValues().map(row => row[0].toString().trim());
        
        const maxNumber = poIds.reduce((max, id) => {
            const match = id.match(/^P(\d+)$/);
            if (match) {
                const currentNum = parseInt(match[1], 10);
                return Math.max(max, currentNum);
            }
            return max;
        }, 0);
        
        nextPoNumber = maxNumber + 1;
    }
    const newInternalPoId = `P${('000' + nextPoNumber).slice(-3)}`;
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
 * Validates and inserts a new Sales Order into the Sales sheet.
 * @param {object} salesData The sales order data from the form.
 * @returns {object} An object containing success status and a message.
 */
function addSalesOrder(salesData) {
  try {
    const salesSheet = getSheetByName(SALES_SHEET_NAME);
    ensureMaterialHeadersExist(salesSheet, SALES_COL_FIRST_MATERIAL);

    // --- 1. Internal Sale ID Generation ---
    let nextSaleNumber = 1;
    if (salesSheet.getLastRow() > 1) {
      const numRows = salesSheet.getLastRow() - 1;
      const saleIds = salesSheet.getRange(2, SALES_COL_INTERNAL_ID + 1, numRows, 1).getValues().map(row => row[0].toString().trim());
      const maxNumber = saleIds.reduce((max, id) => {
        const match = id.match(/^S(\d+)$/);
        if (match) {
          const currentNum = parseInt(match[1], 10);
          return Math.max(max, currentNum);
        }
        return max;
      }, 0);
      nextSaleNumber = maxNumber + 1;
    }
    const newInternalSaleId = `S${('000' + nextSaleNumber).slice(-3)}`;

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

    // --- 3. Handle file uploads ---
    let poLink = '', invLink = '', ewayLink = '';
    if (salesData.filesMeta && salesData.filesMeta.length > 0) {
      // safe SALE_FOLDER fallback
      const saleFolderName = (typeof SALE_FOLDER !== 'undefined' && SALE_FOLDER) ? SALE_FOLDER : 'SalesInternal';

      // Folder path: SALES folder -> YYYY/MM/DD/<PO_NUMBER>
      const dateObj = salesData.poDate ? new Date(salesData.poDate) : new Date();
      const yyyy = dateObj.getFullYear();
      const mm = ('0' + (dateObj.getMonth() + 1)).slice(-2);
      const dd = ('0' + dateObj.getDate()).slice(-2);

      // Use customerPoId (assumed valid) and sanitize for folder name
      const rawPo = salesData.customerPoId.toString().trim();
      const safePo = rawPo.replace(/[\\\/:\*\?"<>\|]/g, '_');

      const subPath = `${saleFolderName}/${yyyy}/${mm}/${dd}/${safePo}`; // include PO number folder
      const urls = uploadFilesToDrive(PARENT_FOLDER_ID, subPath, salesData.filesMeta);
      // Order: [PO, Invoice, EWay]
      poLink = urls[0] || '';
      invLink = urls[1] || '';
      ewayLink = urls[2] || '';
    }
    newRow[SALES_COL_PO_LINK] = poLink;
    newRow[SALES_COL_INV_LINK] = invLink;
    newRow[SALES_COL_EWAY_LINK] = ewayLink;

    // --- 4. Material Quantities ---
    finalHeaders.forEach((matId, index) => {
      const quantity = salesData.materials[matId] || 0;
      if (quantity > 0) {
        newRow[SALES_COL_FIRST_MATERIAL + index] = quantity;
      }
    });

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

    // Vendor ID (optional)
    if (typeof updateData.vendorId !== 'undefined' && updateData.vendorId !== null) {
      const newVendorId = updateData.vendorId || '';
      salesSheet.getRange(rowNumberInSheet, SALES_COL_VENDOR_ID + 1).setValue(newVendorId);
    }

    // Delivery ID (optional)
    if (typeof updateData.deliveryId !== 'undefined' && updateData.deliveryId !== null) {
      const newDeliveryId = updateData.deliveryId || '';
      salesSheet.getRange(rowNumberInSheet, SALES_COL_DELIVERY_ID + 1).setValue(newDeliveryId);
    }

    // Handle file uploads (PO, Invoice, EWay) if provided in updateData.filesMeta
    if (updateData.filesMeta && updateData.filesMeta.length > 0) {
      // safe SALE_FOLDER fallback
      const saleFolderName = (typeof SALE_FOLDER !== 'undefined' && SALE_FOLDER) ? SALE_FOLDER : 'SalesInternal';

      // Use PO date if available to create folder path, otherwise today
      const dateObj = updateData.poDate ? new Date(updateData.poDate) : new Date();
      const yyyy = dateObj.getFullYear();
      const mm = ('0' + (dateObj.getMonth() + 1)).slice(-2);
      const dd = ('0' + dateObj.getDate()).slice(-2);

      // existingPoNumber is assumed valid; use it and sanitize for folder name
      const existingPoNumber = allValues[rowIndexToUpdate][SALES_COL_PO_NUMBER].toString().trim();
      const safePo = existingPoNumber.replace(/[\\\/:\*\?"<>\|]/g, '_');

      const subPath = `${saleFolderName}/${yyyy}/${mm}/${dd}/${safePo}`; // include PO number folder

      // uploadFilesToDrive returns an array of urls in order; caller should pass slots (null allowed)
      const urls = uploadFilesToDrive(PARENT_FOLDER_ID, subPath, updateData.filesMeta);

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

    return { success: true, message: `Sales Order ${internalId} updated successfully.` };

  } catch (e) {
    Logger.log("Error in updateSalesOrder: " + e.toString());
    return { error: true, message: "Server error during Sales Order update: " + e.toString() };
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

    // 2. Get Sales Data
    const salesSheet = getSheetByName(SALES_SHEET_NAME);
    const salesData = salesSheet.getDataRange().getValues();
    if (salesData.length <= 1) return []; // Only headers or empty

    // Headers: SaleID, Sale PO Number, Date of PO, Appointment Date, Invoice, MatId1, MatId2, ...
    const headers = salesData.shift();

    // MatIds start from the column index defined by SALES_COL_FIRST_MATERIAL
    const matIdHeaders = headers.slice(SALES_COL_FIRST_MATERIAL);

    const salesPOs = [];

    for (const row of salesData) {
        // Check for required fields using constant
        if (!row[SALES_COL_PO_NUMBER]) continue;

        // Use constants for fixed columns (updated to reflect new indices)
        const internalId = row[SALES_COL_INTERNAL_ID]; // Column A (Index 0)
        const poNumber = row[SALES_COL_PO_NUMBER];
        const invoiceNumber = row[SALES_COL_INVOICE]; // Column E (Index 4)

        // Date of PO (Column C, index 2)
        const dateOfPO = row[SALES_COL_DATE_PO] instanceof Date ? Utilities.formatDate(row[SALES_COL_DATE_PO], Session.getScriptTimeZone(), "MM/dd/yyyy") : row[SALES_COL_DATE_PO];
        const rawDateOfPO = row[SALES_COL_DATE_PO] instanceof Date ? Utilities.formatDate(row[SALES_COL_DATE_PO], Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
        
        // Appointment Date (Column D, index 3)
        const appointmentDate = row[SALES_COL_APPOINTMENT_DATE] instanceof Date ? Utilities.formatDate(row[SALES_COL_APPOINTMENT_DATE], Session.getScriptTimeZone(), "MM/dd/yyyy") : row[SALES_COL_APPOINTMENT_DATE];
        const rawAppointmentDate = row[SALES_COL_APPOINTMENT_DATE] instanceof Date ? Utilities.formatDate(row[SALES_COL_APPOINTMENT_DATE], Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
        

        let displayItemDetails = []; // For modal display (string array)
        let rawItemDetails = {};     // For edit form population (map of matId: quantity)

        // Iterate through the quantity columns
        for (let i = 0; i < matIdHeaders.length; i++) {
            const matId = matIdHeaders[i].toString().trim();
            
            // Calculate quantity column index using constant
            const quantity = row[i + SALES_COL_FIRST_MATERIAL]; 
            const numericQuantity = typeof quantity === 'number' ? quantity : (parseInt(quantity) || 0);

            if (numericQuantity > 0) {
                const productName = productMap.get(matId) || `Unknown Product (ID: ${matId})`;
                // Format: "Material Name: Quantity" separated by newlines (\n)
                displayItemDetails.push(`${productName}: ${numericQuantity}`);
                rawItemDetails[matId] = numericQuantity;
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
            vendorId: row[SALES_COL_VENDOR_ID] || '',     // NEW: Vendor ID for form
            deliveryId: row[SALES_COL_DELIVERY_ID] || '', // NEW: Delivery ID for form
            poLink: row[SALES_COL_PO_LINK] || '',        // NEW: PO document URL
            invLink: row[SALES_COL_INV_LINK] || '',      // NEW: Invoice document URL
            ewayLink: row[SALES_COL_EWAY_LINK] || '',    // NEW: EWay document URL
            displayItemDetails: displayItemDetails.join('\n'), // For modal button click
            rawItemDetails: rawItemDetails // For edit form population
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
    materialIds.forEach(id => {
      salesSums[id] = 0;
      purchaseSums[id] = 0;
      manualLast[id] = 0;
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

    // 2. Accumulate Purchases
    const poSheet = getSheetByName(PO_SHEET_NAME);
    accumulateFromSheet(poSheet, PO_COL_FIRST_MATERIAL, purchaseSums);

    // 3. Accumulate Sales
    const salesSheet = getSheetByName(SALES_SHEET_NAME);
    accumulateFromSheet(salesSheet, SALES_COL_FIRST_MATERIAL, salesSums);

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

    // 5. Build per-material result list (use materialIds master order)
    const result = materialIds.map(id => {
      const name = productMap.get(id) || id;
      const sales = salesSums[id] || 0;
      const purchases = purchaseSums[id] || 0;
      const manual = manualLast[id] || 0;
      const net = purchases - sales;
      return { id, name, sales, purchases, net, manual };
    });

    // 6. Update Data sheet
    const dataSheet = getSheetByName(DATA_SHEET_NAME);

    // Ensure header row: first cell 'Metric', materials across starting from DATA_COL_FIRST_MATERIAL (column B)
    dataSheet.getRange(1, 1).setValue('Metric');
    if (materialIds.length > 0) {
      dataSheet.getRange(1, DATA_COL_FIRST_MATERIAL + 1, 1, materialIds.length).setValues([materialIds]);
    }

    // Prepare rows: Sales (row 2), Purchases (row 3), Net (row 4), Manual Last (row 5)
    const salesRow = ['Sales'].concat(materialIds.map(id => salesSums[id] || 0));
    const purchaseRow = ['Purchases'].concat(materialIds.map(id => purchaseSums[id] || 0));
    const netRow = ['Net'].concat(materialIds.map(id => (purchaseSums[id] || 0) - (salesSums[id] || 0)));
    const manualRow = ['Manual (last)'].concat(materialIds.map(id => manualLast[id] || 0));

    // Write rows (ensure sheet has enough columns)
    const totalColsNeeded = DATA_COL_FIRST_MATERIAL + materialIds.length + 1; // +1 for first label column
    if (dataSheet.getLastColumn() < totalColsNeeded) {
      // extend header row if necessary
      dataSheet.insertColumnsAfter(dataSheet.getLastColumn(), totalColsNeeded - dataSheet.getLastColumn());
    }

    // Set rows (rows are 1-based)
    dataSheet.getRange(2, 1, 1, 1 + materialIds.length).setValues([salesRow]);
    dataSheet.getRange(3, 1, 1, 1 + materialIds.length).setValues([purchaseRow]);
    dataSheet.getRange(4, 1, 1, 1 + materialIds.length).setValues([netRow]);
    dataSheet.getRange(5, 1, 1, 1 + materialIds.length).setValues([manualRow]);

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
 */
function uploadFilesToDrive(folderId, subPath, files) {
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

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    if (!file) {
      urls[i] = ''; // preserve position for missing files
      continue;
    }
    try {
      let blob;
      if (typeof file.bytes === 'string') {
        // file.bytes is base64 string
        const decoded = Utilities.base64Decode(file.bytes);
        blob = Utilities.newBlob(decoded, file.mimeType || 'application/octet-stream', file.name);
      } else {
        // assume raw bytes (array) or already blob-compatible
        blob = Utilities.newBlob(file.bytes, file.mimeType || 'application/octet-stream', file.name);
      }
      const driveFile = folder.createFile(blob);
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