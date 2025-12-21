// Note: All constants are defined in Constants.gs and are globally available here.
// getSheetByName is defined in Code.gs and is globally available.

/**
 * Checks if a value is null or consists only of zeros.
 * Used to identify incomplete/placeholder entries.
 * @param {any} value The value to check
 * @returns {boolean} True if value is null, empty, or only zeros
 */
function checkNull(value) {
  if (!value && value !== 0) return true; // null, undefined, empty string
  
  const strValue = value.toString().trim();
  if (strValue === '' || strValue === '0') return true;
  
  // Check if string contains only zeros (e.g., "000", "0.00")
  const numericValue = parseFloat(strValue);
  if (!isNaN(numericValue) && numericValue === 0) return true;
  
  return false;
}

/**
 * Performs a sanity check on all sales orders older than today.
 * Validates that all required fields are filled and documents are uploaded.
 * Sends an email report with missing information highlighted.
 * 
 * Required fields checked:
 * - PO Number, PO Date, Despatch Date (Appointment Date)
 * - Distributor (Vendor), Delivery, Amount, GST, Total > 0
 * - All 3 documents uploaded (PO, Invoice, EWay)
 * 
 * @returns {object} { success: true, message: '...' } or { error: true, message: '...' }
 */
function salesSanityCheck() {
  try {
    const salesSheet = getSheetByName(SALES_SHEET_NAME);
    const salesData = salesSheet.getDataRange().getValues();
    
    if (salesData.length <= 1) {
      return { error: true, message: 'No sales data available for sanity check.' };
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // Get Vendor map for display
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

    // Get Delivery map for display
    const deliverySheet = getSheetByName(DELIVERY_SHEET_NAME);
    const deliveryMap = new Map();
    if (deliverySheet) {
      const deliveryData = deliverySheet.getDataRange().getValues();
      for (let i = 1; i < deliveryData.length; i++) {
        if (deliveryData[i][DELIVERY_COL_ID]) {
          deliveryMap.set(deliveryData[i][DELIVERY_COL_ID].toString().trim(), deliveryData[i][DELIVERY_COL_NAME]);
        }
      }
    }

    const issues = [];

    // Check each sales order (skip header row)
    for (let i = 1; i < salesData.length; i++) {
      const row = salesData[i];
      
      // Get PO Date to check if order is older than today
      const poDate = row[SALES_COL_DATE_PO];
      if (!poDate || !(poDate instanceof Date)) continue; // Skip if no valid PO date
      
      const poDateOnly = new Date(poDate);
      poDateOnly.setHours(0, 0, 0, 0);
      
      // Only check orders older than today
      if (poDateOnly >= today) continue;

      // Skip if delivery date (Appointment Date) is missing - these are considered incomplete orders
      const appointmentDate = row[SALES_COL_APPOINTMENT_DATE];
      if (!appointmentDate || !(appointmentDate instanceof Date)) continue;

      const internalId = row[SALES_COL_INTERNAL_ID] || '';
      const poNumber = row[SALES_COL_PO_NUMBER] || '';
      const invoiceNumber = row[SALES_COL_INVOICE] || '';
      const vendorId = row[SALES_COL_VENDOR_ID] ? row[SALES_COL_VENDOR_ID].toString().trim() : '';
      const deliveryId = row[SALES_COL_DELIVERY_ID] ? row[SALES_COL_DELIVERY_ID].toString().trim() : '';
      const amount = row[SALES_COL_AMOUNT];
      const gst = row[SALES_COL_GST];
      const total = row[SALES_COL_TOTAL];
      const poLink = row[SALES_COL_PO_LINK] ? row[SALES_COL_PO_LINK].toString().trim() : '';
      const invLink = row[SALES_COL_INV_LINK] ? row[SALES_COL_INV_LINK].toString().trim() : '';
      const ewayLink = row[SALES_COL_EWAY_LINK] ? row[SALES_COL_EWAY_LINK].toString().trim() : '';

      // Track missing fields
      const missingFields = [];

      // Check PO Number (using checkNull)
      if (checkNull(poNumber)) missingFields.push('PO Number');

      // Check PO Date (already validated above, but check for null)
      if (!poDate) missingFields.push('PO Date');

      // Despatch Date is already validated above (we skip entries without it)
      // No need to add to missingFields since we're skipping those entries

      // Check Distributor (Vendor)
      if (checkNull(vendorId)) missingFields.push('Distributor');

      // Check Delivery
      if (checkNull(deliveryId)) missingFields.push('Delivery');

      // Check Amount > 0
      if (checkNull(amount) || amount <= 0) missingFields.push('Amount');

      // Check GST > 0
      if (checkNull(gst) || gst <= 0) missingFields.push('GST');

      // Check Total > 0
      if (checkNull(total) || total <= 0) missingFields.push('Total');

      // Check PO Document
      if (!poLink) missingFields.push('PO Document');

      // Check Invoice Document
      if (!invLink) missingFields.push('Invoice Document');

      // Check EWay Document
      if (!ewayLink) missingFields.push('EWay Document');

      // Check dispatch quantities vs PO quantities
      const salesHeaders = salesData[0];
      const matIdHeaders = salesHeaders.slice(SALES_COL_FIRST_MATERIAL);
      const totalMatCols = matIdHeaders.length;
      const halfPoint = Math.floor(totalMatCols / 2);
      
      let totalDispatchQty = 0;
      let dispatchIssues = [];
      
      // First half: PO quantities, Second half: Dispatch quantities
      for (let j = 0; j < halfPoint; j++) {
        const poQty = row[SALES_COL_FIRST_MATERIAL + j];
        const dispatchQty = row[SALES_COL_FIRST_MATERIAL + halfPoint + j];
        
        const numericPoQty = typeof poQty === 'number' ? poQty : (parseInt(poQty) || 0);
        const numericDispatchQty = typeof dispatchQty === 'number' ? dispatchQty : (parseInt(dispatchQty) || 0);
        
        totalDispatchQty += numericDispatchQty;
        
        // Check if dispatch quantity exceeds PO quantity for this material
        if (numericDispatchQty > numericPoQty && numericPoQty > 0) {
          const matId = matIdHeaders[j] ? matIdHeaders[j].toString().trim() : `Material ${j + 1}`;
          dispatchIssues.push(`${matId}: Dispatch(${numericDispatchQty}) > PO(${numericPoQty})`);
        }
      }
      
      // Check if total dispatch quantity is 0 (user hasn't entered dispatch quantities)
      if (totalDispatchQty === 0) {
        missingFields.push('Dispatch Quantities (all zero)');
      }
      
      // Add dispatch validation issues to missing fields
      if (dispatchIssues.length > 0) {
        missingFields.push('Dispatch > PO: ' + dispatchIssues.join(', '));
      }

      // If any fields are missing, add to issues list
      if (missingFields.length > 0) {
        issues.push({
          internalId: internalId,
          poNumber: poNumber || 'N/A',
          poDate: poDate instanceof Date ? Utilities.formatDate(poDate, Session.getScriptTimeZone(), "dd-MM-yyyy") : 'N/A',
          despatchDate: appointmentDate instanceof Date ? Utilities.formatDate(appointmentDate, Session.getScriptTimeZone(), "dd-MM-yyyy") : 'Missing',
          distributor: vendorId ? (vendorMap.get(vendorId) || vendorId) : 'Missing',
          delivery: deliveryId ? (deliveryMap.get(deliveryId) || deliveryId) : 'Missing',
          amount: amount || 0,
          gst: gst || 0,
          total: total || 0,
          missingFields: missingFields
        });
      }
    }

    // If no issues found, return success
    if (issues.length === 0) {
      return { success: true, message: 'All sales orders are complete. No issues found.' };
    }

    // Generate HTML email with issues
    const emailHtml = generateSanityCheckEmail(issues);

    // Send email
    const emailList = (typeof SANITY_REPORT_EMAIL !== 'undefined' && SANITY_REPORT_EMAIL) ? SANITY_REPORT_EMAIL : '';
    if (!emailList) {
      Logger.log('salesSanityCheck: SANITY_REPORT_EMAIL not configured');
      return { success: true, message: `Found ${issues.length} issues but email not configured.`, issuesCount: issues.length };
    }

    const emails = emailList.split(';').map(e => e.trim()).filter(e => e);
    if (emails.length === 0) {
      return { error: true, message: 'No valid email addresses configured for sanity reports.' };
    }

    const toEmail = emails[0];
    const ccEmails = emails.slice(1).join(',');

    const subject = `Sales Data Sanity Check Report - ${issues.length} Issue(s) Found`;
    const emailOptions = {
      to: toEmail,
      subject: subject,
      htmlBody: emailHtml
    };
    if (ccEmails) emailOptions.cc = ccEmails;

    try {
      MailApp.sendEmail(emailOptions);
    } catch (emailError) {
      Logger.log('salesSanityCheck: Failed to send email - ' + emailError.toString());
      return { error: true, message: 'Failed to send sanity check email: ' + emailError.toString() };
    }

    return {
      success: true,
      message: `Sanity check completed. Found ${issues.length} issue(s). Report sent to ${toEmail}${ccEmails ? ' (CC: ' + ccEmails + ')' : ''}.`,
      issuesCount: issues.length
    };

  } catch (e) {
    Logger.log('Error in salesSanityCheck: ' + e.toString());
    return { error: true, message: 'Failed to perform sanity check: ' + e.toString() };
  }
}

/**
 * Generates HTML email body for sanity check report.
 * @param {Array<object>} issues Array of issue objects
 * @returns {string} HTML email body
 */
function generateSanityCheckEmail(issues) {
  let html = `
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; }
          table { border-collapse: collapse; width: 100%; margin-top: 20px; }
          th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
          th { background-color: #4285f4; color: white; font-weight: bold; }
          tr:nth-child(even) { background-color: #f2f2f2; }
          .checkmark { color: green; font-weight: bold; }
          .xmark { color: red; font-weight: bold; }
          .missing { color: red; font-style: italic; }
          h2 { color: #333; }
          .summary { background-color: #fff3cd; padding: 10px; border-left: 4px solid #ffc107; margin-bottom: 20px; }
        </style>
      </head>
      <body>
        <h2>Sales Data Sanity Check Report</h2>
        <div class="summary">
          <strong>Total Issues Found:</strong> ${issues.length}<br>
          <strong>Report Date:</strong> ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm")}
        </div>
        <table>
          <thead>
            <tr>
              <th>Internal ID</th>
              <th>PO Number</th>
              <th>PO Date</th>
              <th>Despatch Date</th>
              <th>Distributor</th>
              <th>Delivery</th>
              <th>Amount</th>
              <th>GST</th>
              <th>Total</th>
              <th>PO Doc</th>
              <th>Inv Doc</th>
              <th>EWay Doc</th>
              <th>Missing Fields</th>
            </tr>
          </thead>
          <tbody>
  `;

  issues.forEach(issue => {
    const hasPONumber = !issue.missingFields.includes('PO Number');
    const hasPODate = !issue.missingFields.includes('PO Date');
    const hasDespatchDate = !issue.missingFields.includes('Despatch Date');
    const hasDistributor = !issue.missingFields.includes('Distributor');
    const hasDelivery = !issue.missingFields.includes('Delivery');
    const hasAmount = !issue.missingFields.includes('Amount');
    const hasGST = !issue.missingFields.includes('GST');
    const hasTotal = !issue.missingFields.includes('Total');
    const hasPODoc = !issue.missingFields.includes('PO Document');
    const hasInvDoc = !issue.missingFields.includes('Invoice Document');
    const hasEWayDoc = !issue.missingFields.includes('EWay Document');

    html += `
      <tr>
        <td>${issue.internalId}</td>
        <td>${hasPONumber ? issue.poNumber : '<span class="xmark">✗</span>'}</td>
        <td>${hasPODate ? issue.poDate : '<span class="xmark">✗</span>'}</td>
        <td>${hasDespatchDate ? issue.despatchDate : '<span class="xmark">✗</span>'}</td>
        <td>${hasDistributor ? issue.distributor : '<span class="xmark">✗</span>'}</td>
        <td>${hasDelivery ? issue.delivery : '<span class="xmark">✗</span>'}</td>
        <td>${hasAmount ? issue.amount : '<span class="xmark">✗</span>'}</td>
        <td>${hasGST ? issue.gst : '<span class="xmark">✗</span>'}</td>
        <td>${hasTotal ? issue.total : '<span class="xmark">✗</span>'}</td>
        <td>${hasPODoc ? '<span class="checkmark">✓</span>' : '<span class="xmark">✗</span>'}</td>
        <td>${hasInvDoc ? '<span class="checkmark">✓</span>' : '<span class="xmark">✗</span>'}</td>
        <td>${hasEWayDoc ? '<span class="checkmark">✓</span>' : '<span class="xmark">✗</span>'}</td>
        <td class="missing">${issue.missingFields.join(', ')}</td>
      </tr>
    `;
  });

  html += `
          </tbody>
        </table>
        <p style="margin-top: 20px; color: #666;">
          <strong>Note:</strong> <span class="checkmark">✓</span> indicates field is complete, 
          <span class="xmark">✗</span> indicates missing or invalid data.
        </p>
      </body>
    </html>
  `;

  return html;
}
