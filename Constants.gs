// --- Sheet Names ---
const PO_SHEET_NAME = 'Purchase'; // Changed from 'PurchaseOrders'
const SALES_SHEET_NAME = 'Sales'; // Changed from 'SalesOrders'
const MATERIAL_SHEET_NAME = 'Material'; // Changed from 'Materials'
const OTP_SHEET_NAME = 'OTPs'; // Added for consistency

// --- PURCHASE ORDER COLUMN INDICES (0-based) ---
// Note: This structure now supports two PO ID columns.
const P_COL_INTERNAL_ID = 0;      // Script-generated P001, P002... (Column A)
const PO_COL_ID = 1;              // User/Vendor PO ID (Column B)
const PO_COL_DATE = 2;            // PO Date (Column C)
const PO_COL_DESPATCH_DATE = 3;   // Despatch Date (Column D)
const PO_COL_INVOICE = 4;         // Invoice Number (Column E)
const PO_COL_FIRST_MATERIAL = 5;  // Start of material quantities (Column F onwards)

// --- SALES ORDER COLUMN INDICES (0-based) ---
// NEW ORDER: Internal ID, PO Number, Date of PO, Appointment Date, Invoice, First Material
const SALES_COL_INTERNAL_ID = 0;      // Column A
const SALES_COL_PO_NUMBER = 1;        // Column B
const SALES_COL_DATE_PO = 2;          // Column C: Date of PO (Shifted from 3)
const SALES_COL_APPOINTMENT_DATE = 3; // Column D: Appointment Date (Shifted from 4)
const SALES_COL_INVOICE = 4;          // Column E: Invoice Number (Shifted from 2)
const SALES_COL_FIRST_MATERIAL = 5;   // Column F: Start of material quantities (Shifted from 5)

// --- MATERIAL COLUMN INDICES (0-based) ---
const MATERIAL_COL_ID = 0;
const MATERIAL_COL_NAME = 1;

// --- OTP COLUMN INDICES (0-based) ---
const OTP_COL_EMAIL = 0;
const OTP_COL_DATE = 1;
const OTP_COL_OTP = 2;