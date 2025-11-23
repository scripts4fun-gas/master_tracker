/**
 * Handles the user's request for an OTP, generating a new one if necessary, 
 * sending an email, and updating the OTPs sheet.
 * @param {string} email The user's email address.
 * @returns {object} An object containing success status and a message.
 */
function processEmailRequest(email) {
  email = email.toLowerCase().trim();

  try {
    // getSheetByName is defined in Code.gs and is globally available
    const otpSheet = getSheetByName(OTP_SHEET_NAME); 
    const data = otpSheet.getDataRange().getValues();
    data.shift(); // Remove headers

    const today = new Date().toDateString();
    // Use constant for Email column index
    let existingEntry = data.find(row => row[OTP_COL_EMAIL].toString().toLowerCase().trim() === email);

    if (existingEntry) {
      // Use constant for Date column index
      const sentDate = existingEntry[OTP_COL_DATE] ? new Date(existingEntry[OTP_COL_DATE]).toDateString() : null;

      if (sentDate === today) {
        // OTP already sent today
        return { success: true, message: "An OTP was already sent to your email today. Please use the existing code." };
      }
    }

    // Generate new OTP (6 digits)
    const newOtp = Math.floor(100000 + Math.random() * 900000).toString();

    // Send email
    sendEmailWithOTP(email, newOtp);

    // Update/Add entry in OTPs sheet, ordered by constants
    const now = new Date();
    const newRow = [];
    newRow[OTP_COL_EMAIL] = email;
    newRow[OTP_COL_DATE] = now;
    newRow[OTP_COL_OTP] = newOtp;

    if (existingEntry) {
      // Find the row index to update (data index + 2)
      const rowIndex = data.findIndex(row => row[OTP_COL_EMAIL].toString().toLowerCase().trim() === email) + 2;
      // Start writing from the first column (Email), which is index 1 in Sheet Range
      otpSheet.getRange(rowIndex, OTP_COL_EMAIL + 1, 1, newRow.length).setValues([newRow]); 
    } else {
      // Add new row
      otpSheet.appendRow(newRow);
    }

    return { success: true, message: `A new 6-digit OTP has been sent to ${email}.` };

  } catch (e) {
    Logger.log("Error in processEmailRequest: " + e.toString());
    return { success: false, message: "Error processing request. Check logs for details. Make sure you have permission to send emails and sheets exist." };
  }
}

/**
 * Sends the OTP to the specified email address.
 * NOTE: This function requires the script to be authorized to send emails.
 * @param {string} recipient The email address to send to.
 * @param {string} otp The generated OTP.
 */
function sendEmailWithOTP(recipient, otp) {
  MailApp.sendEmail({
    to: recipient,
    subject: 'Inventory Tracker Login OTP',
    body: `Your One-Time Password (OTP) for the Inventory Tracker is: ${otp}. This code is valid for today.`,
  });
}

/**
 * Validates the entered OTP against the one stored in the sheet for today.
 * @param {string} email The user's email.
 * @param {string} otp The OTP entered by the user.
 * @returns {object} An object containing success status and a message.
 */
function validateOTP(email, otp) {
  email = email.toLowerCase().trim();
  otp = otp.trim();
  const today = new Date().toDateString();

  try {
    // getSheetByName is defined in Code.gs and is globally available
    const otpSheet = getSheetByName(OTP_SHEET_NAME); 
    const data = otpSheet.getDataRange().getValues();
    data.shift(); // Remove headers

    // Use constant for Email column index
    const entry = data.find(row => row[OTP_COL_EMAIL].toString().toLowerCase().trim() === email);

    if (!entry) {
      return { success: false, message: "Email not found. Please request an OTP first." };
    }

    // Use constant for Date column index
    const sentDate = new Date(entry[OTP_COL_DATE]).toDateString();
    // Use constant for OTP column index
    const storedOtp = entry[OTP_COL_OTP] ? entry[OTP_COL_OTP].toString().trim() : '';

    if (sentDate !== today) {
      return { success: false, message: "The OTP has expired. Please request a new one." };
    }

    if (storedOtp === otp) {
      return { success: true, message: "Login successful!" };
    } else {
      return { success: false, message: "Invalid OTP. Please check the code and try again." };
    }
  } catch (e) {
    Logger.log("Error in validateOTP: " + e.toString());
    return { success: false, message: "Validation error. Please try again." };
  }
}