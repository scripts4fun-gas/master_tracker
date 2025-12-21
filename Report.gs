/**
 * Generates a weekly sales report for a given date range.
 * Creates a new Google Sheet with sales orders dispatched within the date range.
 * Report is saved in Reports/YYYY/MM/DD folder and emailed to REPORT_EMAIL recipients.
 * 
 * @param {string} startDate Start date in YYYY-MM-DD format
 * @param {string} endDate End date in YYYY-MM-DD format
 * @returns {object} { success: true, message: '...', reportUrl: '...' } or { error: true, message: '...' }
 */
function generateWeeklySalesReport(startDate, endDate) {
  try {
    // 1. Parse and validate dates
    if (!startDate || !endDate) {
      return { error: true, message: 'Start date and end date are required.' };
    }

    const startDateObj = new Date(startDate);
    const endDateObj = new Date(endDate);
    
    if (isNaN(startDateObj.getTime()) || isNaN(endDateObj.getTime())) {
      return { error: true, message: 'Invalid date format. Use YYYY-MM-DD.' };
    }

    if (startDateObj > endDateObj) {
      return { error: true, message: 'Start date must be before or equal to end date.' };
    }

    // Reset time to start of day for accurate comparison
    startDateObj.setHours(0, 0, 0, 0);
    endDateObj.setHours(23, 59, 59, 999);

    // Format dates for display
    const startDateFormatted = Utilities.formatDate(startDateObj, Session.getScriptTimeZone(), "dd-MM-yyyy");
    const endDateFormatted = Utilities.formatDate(endDateObj, Session.getScriptTimeZone(), "dd-MM-yyyy");

    // 2. Get Sales data and filter by appointment date (despatch date) range
    const salesSheet = getSheetByName(SALES_SHEET_NAME);
    const salesData = salesSheet.getDataRange().getValues();
    
    if (salesData.length <= 1) {
      return { error: true, message: 'No sales data available.' };
    }

    // Get Vendor map for DC/Customer column
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

    // Get material map for quantity calculation
    const productMap = getMaterialMap();
    const salesHeaders = salesData[0];
    const matIdHeaders = salesHeaders.slice(SALES_COL_FIRST_MATERIAL);
    
    // Material columns are duplicated: first half is PO quantities, second half is Dispatch quantities
    const totalMatCols = matIdHeaders.length;
    const halfPoint = Math.floor(totalMatCols / 2);

    // Filter sales orders within date range
    const filteredSales = [];
    for (let i = 1; i < salesData.length; i++) {
      const row = salesData[i];
      const appointmentDate = row[SALES_COL_APPOINTMENT_DATE];
      
      if (appointmentDate && appointmentDate instanceof Date) {
        const apptDateOnly = new Date(appointmentDate);
        apptDateOnly.setHours(0, 0, 0, 0);
        
        if (apptDateOnly >= startDateObj && apptDateOnly <= endDateObj) {
          // Calculate CB as per PO (first half of material columns) and CB Despatched (second half)
          let cbAsPerPO = 0;
          let cbDespatched = 0;
          
          // First half: PO quantities
          for (let j = 0; j < halfPoint; j++) {
            const quantity = row[SALES_COL_FIRST_MATERIAL + j];
            const numericQty = typeof quantity === 'number' ? quantity : (parseInt(quantity) || 0);
            cbAsPerPO += numericQty;
          }
          
          // Second half: Dispatch quantities
          for (let j = halfPoint; j < totalMatCols; j++) {
            const quantity = row[SALES_COL_FIRST_MATERIAL + j];
            const numericQty = typeof quantity === 'number' ? quantity : (parseInt(quantity) || 0);
            cbDespatched += numericQty;
          }

          const vendorId = row[SALES_COL_VENDOR_ID] ? row[SALES_COL_VENDOR_ID].toString().trim() : '';
          const vendorName = vendorId ? (vendorMap.get(vendorId) || vendorId) : '';

          filteredSales.push({
            poNumber: row[SALES_COL_PO_NUMBER] || '',
            cbAsPerPO: cbAsPerPO,
            despatchDate: Utilities.formatDate(appointmentDate, Session.getScriptTimeZone(), "dd-MM-yyyy"),
            dcCustomer: vendorName,
            invoiceNumber: row[SALES_COL_INVOICE] || '',
            cbDespatched: cbDespatched,
            amount: row[SALES_COL_AMOUNT] || 0,
            gst: row[SALES_COL_GST] || 0,
            total: row[SALES_COL_TOTAL] || 0
          });
        }
      }
    }

    if (filteredSales.length === 0) {
      return { error: true, message: `No sales orders found between ${startDateFormatted} and ${endDateFormatted}.` };
    }

    // 3. Create new spreadsheet
    const reportTitle = `SE_Weekly_Report_${startDateFormatted}_${endDateFormatted}`;
    const newSpreadsheet = SpreadsheetApp.create(reportTitle);
    const reportSheet = newSpreadsheet.getActiveSheet();
    reportSheet.setName('Weekly Sales Report');

    // 4. Set up headers
    const headers = [
      'PO Number',
      'CB as per PO',
      'Despatch Date',
      'DC/Customer',
      'Invoice Number',
      'CB Despatched',
      'Base Amount',
      'GST',
      'Total'
    ];
    reportSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    reportSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');

    // 5. Populate data
    const reportData = filteredSales.map(sale => [
      sale.poNumber,
      sale.cbAsPerPO,
      sale.despatchDate,
      sale.dcCustomer,
      sale.invoiceNumber,
      sale.cbDespatched,
      sale.amount,
      sale.gst,
      sale.total
    ]);

    if (reportData.length > 0) {
      reportSheet.getRange(2, 1, reportData.length, headers.length).setValues(reportData);
      // Set all data cells to left alignment
      reportSheet.getRange(2, 1, reportData.length, headers.length).setHorizontalAlignment('left');
    }

    // 6. Add totals row
    if (filteredSales.length > 0) {
      const totalCbAsPerPO = filteredSales.reduce((sum, sale) => sum + sale.cbAsPerPO, 0);
      const totalCbDespatched = filteredSales.reduce((sum, sale) => sum + sale.cbDespatched, 0);
      const totalAmount = filteredSales.reduce((sum, sale) => sum + sale.amount, 0);
      const totalGst = filteredSales.reduce((sum, sale) => sum + sale.gst, 0);
      const totalTotal = filteredSales.reduce((sum, sale) => sum + sale.total, 0);

      const totalsRow = [
        '',
        totalCbAsPerPO,
        '',
        '',
        'TOTAL',
        totalCbDespatched,
        totalAmount,
        totalGst,
        totalTotal
      ];

      const totalsRowNumber = reportData.length + 2; // +2 because header is row 1, data starts at row 2
      reportSheet.getRange(totalsRowNumber, 1, 1, headers.length).setValues([totalsRow]);
      reportSheet.getRange(totalsRowNumber, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8eaf6').setHorizontalAlignment('left');
    }

    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      reportSheet.autoResizeColumn(i);
    }

    // 6. Move report to Reports folder with date hierarchy
    const reportFile = DriveApp.getFileById(newSpreadsheet.getId());
    const reportsFolderName = (typeof REPORTS_FOLDER !== 'undefined' && REPORTS_FOLDER) ? REPORTS_FOLDER : 'Reports';
    
    // Use current date for folder organization
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = ('0' + (today.getMonth() + 1)).slice(-2);
    const dd = ('0' + today.getDate()).slice(-2);

    // Navigate/create folder hierarchy: Reports/YYYY/MM/DD
    let reportsFolder;
    try {
      reportsFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    } catch (e) {
      Logger.log('generateWeeklySalesReport: Failed to access PARENT_FOLDER_ID, using root folder');
      reportsFolder = DriveApp.getRootFolder();
    }

    const folderPath = [reportsFolderName, yyyy.toString(), mm, dd];
    for (const folderName of folderPath) {
      let found = false;
      const folders = reportsFolder.getFoldersByName(folderName);
      if (folders.hasNext()) {
        reportsFolder = folders.next();
        found = true;
      }
      if (!found) {
        reportsFolder = reportsFolder.createFolder(folderName);
      }
    }

    // Move file to target folder and remove from root
    reportsFolder.addFile(reportFile);
    DriveApp.getRootFolder().removeFile(reportFile);

    const reportUrl = newSpreadsheet.getUrl();

    // 7. Set file sharing permissions and send email notification
    const emailList = (typeof REPORT_EMAIL !== 'undefined' && REPORT_EMAIL) ? REPORT_EMAIL : '';
    if (emailList) {
      const emails = emailList.split(';').map(e => e.trim()).filter(e => e);
      
      if (emails.length > 0) {
        const toEmail = emails[0];
        const ccEmails = emails.slice(1);

        // Set permissions: first email gets VIEW only, rest get EDIT access
        try {
          reportFile.addViewer(toEmail);
          
          for (let i = 1; i < emails.length; i++) {
            reportFile.addEditor(ccEmails[i - 1]);
          }
          
          // Set general sharing to restricted (only people with explicit permissions)
          reportFile.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
        } catch (permError) {
          Logger.log('generateWeeklySalesReport: Failed to set permissions - ' + permError.toString());
          // Fallback to anyone with link can view
          reportFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        }

        const subject = `Weekly Sales Report: ${startDateFormatted} to ${endDateFormatted}`;
        const body = `Dear Team,\n\nPlease find the weekly sales report for the period ${startDateFormatted} to ${endDateFormatted}.\n\nTotal Orders: ${filteredSales.length}\n\nReport Link: ${reportUrl}\n\nBest regards,\nInventory Management System`;

        const emailOptions = { 
          to: toEmail, 
          subject: subject, 
          body: body 
        };
        
        if (ccEmails.length > 0) {
          emailOptions.cc = ccEmails.join(',');
        }
        
        try {
          MailApp.sendEmail(emailOptions);
        } catch (emailError) {
          Logger.log('generateWeeklySalesReport: Failed to send email - ' + emailError.toString());
          // Continue even if email fails
        }
      }
    } else {
      // No email list configured, set default sharing
      reportFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }

    return {
      success: true,
      message: `Report generated successfully with ${filteredSales.length} sales orders.`,
      reportUrl: reportUrl,
      ordersCount: filteredSales.length
    };

  } catch (e) {
    Logger.log('Error in generateWeeklySalesReport: ' + e.toString());
    return { error: true, message: 'Failed to generate report: ' + e.toString() };
  }
}

/**
 * Generates a weekly sales report for the last 7 days (from 7 days ago to today).
 * Convenience function that calculates the date range automatically.
 * 
 * @returns {object} { success: true, message: '...', reportUrl: '...' } or { error: true, message: '...' }
 */
function generateReportToday() {
  try {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const sevenDaysAgo = new Date(today);
    sevenDaysAgo.setDate(today.getDate() - 7);
    
    // Format dates as YYYY-MM-DD
    const startDate = Utilities.formatDate(sevenDaysAgo, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const endDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    return generateWeeklySalesReport(startDate, endDate);
  } catch (e) {
    Logger.log('Error in generateReportToday: ' + e.toString());
    return { error: true, message: 'Failed to generate today\'s report: ' + e.toString() };
  }
}