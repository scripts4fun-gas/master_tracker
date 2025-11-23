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
    
    // 0. Ensure Headers
    ensureMaterialHeadersExist(salesSheet, SALES_COL_FIRST_MATERIAL);
    
    // --- 1. Internal Sale ID Generation (Based on Column A: SALES_COL_INTERNAL_ID) ---
    let nextSaleNumber = 1;
    if (salesSheet.getLastRow() > 1) {
        const numRows = salesSheet.getLastRow() - 1;
        // Read all existing internal Sale IDs from the sheet (Column A, index SALES_COL_INTERNAL_ID + 1)
        const saleIds = salesSheet.getRange(2, SALES_COL_INTERNAL_ID + 1, numRows, 1).getValues().map(row => row[0].toString().trim());
        
        const maxNumber = saleIds.reduce((max, id) => {
            const match = id.match(/^S(\d+)$/); // Assuming 'S' prefix for Sales
            if (match) {
                const currentNum = parseInt(match[1], 10);
                return Math.max(max, currentNum);
            }
            return max;
        }, 0);
        
        nextSaleNumber = maxNumber + 1;
    }
    const newInternalSaleId = `S${('000' + nextSaleNumber).slice(-3)}`;
    // --- End Sale ID Generation ---

    // 2. Prepare Data Row
    
    // Get the final list of material headers from the sheet to match column order
    const finalHeaders = salesSheet.getRange(1, SALES_COL_FIRST_MATERIAL + 1, 1, salesSheet.getLastColumn() - SALES_COL_FIRST_MATERIAL).getValues()[0].map(h => h.toString().trim());

    // Create a row template large enough for all columns
    const numColumns = salesSheet.getLastColumn();
    const newRow = new Array(numColumns).fill(''); 

    // Insert fixed fields using NEW constants based on the user's requested order:
    // A: 0, B: 1, C: 2 (Date), D: 3 (Appt Date), E: 4 (Invoice)
    newRow[SALES_COL_INTERNAL_ID] = newInternalSaleId;         // Column A: 0
    newRow[SALES_COL_PO_NUMBER] = salesData.customerPoId;      // Column B: 1
    newRow[SALES_COL_DATE_PO] = salesData.poDate ? new Date(salesData.poDate) : '';         // Column C: 2 (Date of PO)
    newRow[SALES_COL_APPOINTMENT_DATE] = salesData.appointmentDate ? new Date(salesData.appointmentDate) : ''; // Column D: 3 (Appointment Date)
    newRow[SALES_COL_INVOICE] = salesData.invoiceNumber;      // Column E: 4 (Invoice Number)
    
    // Insert material quantities starting from SALES_COL_FIRST_MATERIAL
    finalHeaders.forEach((matId, index) => {
        const quantity = salesData.materials[matId] || 0;
        // Check if matId exists in submitted data and is > 0
        if (quantity > 0) {
            newRow[SALES_COL_FIRST_MATERIAL + index] = quantity;
        }
    });
    
    // 4. Append Data
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
    // Column index is 1-based, so +1
    salesSheet.getRange(rowNumberInSheet, SALES_COL_APPOINTMENT_DATE + 1).setValue(newAppointmentDate); 

    // Invoice Number (Column E, index 4)
    const newInvoiceNumber = updateData.invoiceNumber || '';
    salesSheet.getRange(rowNumberInSheet, SALES_COL_INVOICE + 1).setValue(newInvoiceNumber); // Column index is 1-based, so +1

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