// --- Sheet Names ---
const PO_SHEET_NAME = 'Purchase'; // Changed from 'PurchaseOrders'
const SALES_SHEET_NAME = 'Sales'; // Changed from 'SalesOrders'
const MATERIAL_SHEET_NAME = 'Material'; // Changed from 'Materials'
const OTP_SHEET_NAME = 'OTPs'; // Added for consistency
const MANUAL_SHEET_NAME = 'Manual'; // Added for manual stock entries
const COUNTERS_SHEET_NAME = 'Counters'; // Added for tracking internal ID counters

// --- PURCHASE ORDER COLUMN INDICES (0-based) ---
// Note: This structure now supports two PO ID columns.
const P_COL_INTERNAL_ID = 0;      // Script-generated P001, P002... (Column A)
const PO_COL_ID = 1;              // User/Vendor PO ID (Column B)
const PO_COL_DATE = 2;            // PO Date (Column C)
const PO_COL_DESPATCH_DATE = 3;   // Despatch Date (Column D)
const PO_COL_INVOICE = 4;         // Invoice Number (Column E)
const PO_COL_SUPPLIER_ID = 5;     // Supplier ID (Column F)
const PO_COL_PO_LINK = 6;         // PO Link (Column G)
const PO_COL_INV_LINK = 7;        // Invoice Link (Column H)
const PO_COL_EWAY_LINK = 8;       // EWay Link (Column I)
const PO_COL_FIRST_MATERIAL = 9;  // Start of material quantities (Column J onwards)

// --- SALES ORDER COLUMN INDICES (0-based) ---
// Internal ID, PO Number, Date of PO, Appointment Date, Invoice, VendorId, DeliveryId, Amount, GST, Total, PO Link, Invoice Link, EWay Link, First Material
const SALES_COL_INTERNAL_ID = 0;      // Column A
const SALES_COL_PO_NUMBER = 1;        // Column B
const SALES_COL_DATE_PO = 2;          // Column C: Date of PO
const SALES_COL_APPOINTMENT_DATE = 3; // Column D: Appointment Date
const SALES_COL_INVOICE = 4;          // Column E: Invoice Number
const SALES_COL_VENDOR_ID = 5;        // Column F: Vendor ID
const SALES_COL_DELIVERY_ID = 6;      // Column G: Delivery ID
const SALES_COL_AMOUNT = 7;           // Column H: Amount (NEW)
const SALES_COL_GST = 8;              // Column I: GST (NEW)
const SALES_COL_TOTAL = 9;            // Column J: Total (NEW)
const SALES_COL_PO_LINK = 10;         // Column K: PO Link (shifted)
const SALES_COL_INV_LINK = 11;        // Column L: Invoice Link (shifted)
const SALES_COL_EWAY_LINK = 12;       // Column M: EWay Link (shifted)
const SALES_COL_FIRST_MATERIAL = 13;  // Column N: Start of material quantities (shifted)

// --- GOOGLE DRIVE PARENT FOLDER ID ---
const PARENT_FOLDER_ID = '1u-kJQ98zjDaRVijEZuCeqtB8SKzsf8b4';

// New: top-level subfolder name used for Sales uploads (under PARENT_FOLDER_ID)
const SALE_FOLDER = 'Sales';

// New: top-level subfolder name used for Purchase uploads (under PARENT_FOLDER_ID)
const PURCHASE_FOLDER = 'Purchase';

// New: top-level subfolder name used for Reports (under PARENT_FOLDER_ID)
const REPORTS_FOLDER = 'Reports';

// --- MATERIAL COLUMN INDICES (0-based) ---
const MATERIAL_COL_ID = 0;
const MATERIAL_COL_NAME = 1;

// --- OTP COLUMN INDICES (0-based) ---
const OTP_COL_EMAIL = 0;
const OTP_COL_DATE = 1;
const OTP_COL_OTP = 2;

// --- MANUAL STOCK SHEET CONSTANTS (0-based) ---
// Internal ID, Timestamp column and start of material quantities for Manual sheet
const MANUAL_COL_INTERNAL_ID = 0;              // Column A: Internal Manual ID (M001...)
const MANUAL_COL_TIMESTAMP = 1;                // Column B: Timestamp of manual entry
const MANUAL_COL_FIRST_MATERIAL = 2;           // Column C onwards: Material quantities

// --- DATA SHEET CONSTANTS (0-based) ---
// Data sheet holds material headers starting from second column (Column B)
const DATA_SHEET_NAME = 'Data';
const DATA_COL_FIRST_MATERIAL = 1; // Column B onwards contain material IDs/values

// --- VENDOR SHEET CONSTANTS (0-based) ---
const VENDOR_SHEET_NAME = 'Vendor';
const VENDOR_COL_ID = 0;    // Column A: Vendor ID
const VENDOR_COL_NAME = 1;  // Column B: Vendor Name

// --- SUPPLIER SHEET CONSTANTS (0-based) ---
const SUPPLIER_SHEET_NAME = 'Supplier';
const SUPPLIER_COL_ID = 0;    // Column A: Supplier ID
const SUPPLIER_COL_NAME = 1;  // Column B: Supplier Name

// --- DELIVERY SHEET CONSTANTS ---
const DELIVERY_SHEET_NAME = 'Delivery';
const DELIVERY_COL_ID = 0;    // Column A: Delivery ID
const DELIVERY_COL_NAME = 1;  // Column B: Delivery Name
const DELIVERY_COL_VEHICLE = 2; // Column C: Delivery Vehicle (new)

// --- COUNTERS SHEET CONSTANTS ---
const COUNTERS_COL_TYPE = 0;    // Column A: Type (Sales, Purchase, Manual)
const COUNTERS_COL_COUNTER = 1; // Column B: Counter value

// Email(s) to receive vehicle/PO detail notifications (semicolon-separated for multiple recipients)
const NOTIFY_EMAIL = 'schnprabhu@gmail.com;symphonyashwin92@gmail.com;uaprabhu86@gmail.com';

// Email(s) to receive weekly sales reports (semicolon-separated for multiple recipients)
const REPORT_EMAIL = 'schnprabhu@gmail.com;symphonyashwin92@gmail.com;uaprabhu86@gmail.com';