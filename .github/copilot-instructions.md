# Google Apps Script Inventory Tracker - AI Coding Instructions

## Project Architecture

This is a **Google Apps Script Web Application** for inventory management, deployed as a standalone web app with backend spreadsheet integration. Not a Node.js/npm project—all code runs on Google's V8 runtime within Apps Script.

### Core Components
- **`Code.gs`**: Web app entry point (`doGet()`) and shared utilities (`getSheetByName()`)
- **`Constants.gs`**: Global constants for sheet names, column indices, and folder IDs
- **`Purchase.gs`**: Main business logic (Purchase/Sales/Manual stock operations, file uploads, stock computation)
- **`Otp.gs`**: OTP-based authentication (email sending, validation)
- **`index.html`**: Single-page web UI (HTML + inline JavaScript using `google.script.run` API)
- **`appsscript.json`**: Manifest defining OAuth scopes, Drive API, timezone, and webapp config

### Data Model (Google Sheets)
- **`Purchase`**: Purchase orders with internal ID (P001...), vendor PO ID, dates, invoice, material quantities
- **`Sales`**: Sales orders with internal ID (S001...), customer PO, dates, vendor/delivery IDs, Drive links (PO/Invoice/EWay), material quantities
- **`Manual`**: Manual stock adjustments with internal ID (M001...), timestamp, material quantities
- **`Material`**: Master list of materials (Product ID, Product Name)
- **`Vendor`**: Vendor registry (Vendor ID, Vendor Name)
- **`Delivery`**: Delivery registry (Delivery ID, Delivery Name, Vehicle)
- **`Data`**: Computed stock summary (Sales, Purchases, Net, Manual last entry) updated by `computeAndUpdateStock()`
- **`OTPs`**: Email-based OTP storage for authentication

## Critical Patterns

### 1. Column Indexing Convention
**All sheet column access uses 0-based constants from `Constants.gs`**. Never hardcode column numbers:
```javascript
// CORRECT:
newRow[SALES_COL_INVOICE] = invoiceNumber;
sheet.getRange(rowIndex, SALES_COL_VENDOR_ID + 1).setValue(vendorId); // +1 for 1-based Range API

// WRONG:
newRow[4] = invoiceNumber; // Fragile if columns change
```

### 2. Internal ID Generation Pattern
Purchase/Sales/Manual sheets use script-generated IDs (P001, S001, M001):
```javascript
// Find max existing ID, increment, format with zero-padding
const maxNumber = ids.reduce((max, id) => {
  const match = id.match(/^P(\d+)$/);
  return match ? Math.max(max, parseInt(match[1], 10)) : max;
}, 0);
const newId = `P${('000' + (maxNumber + 1)).slice(-3)}`;
```

### 3. Dynamic Material Headers
Material columns are appended dynamically when new products are added. Always call `ensureMaterialHeadersExist(sheet, materialStartIndex)` before inserting data:
```javascript
// Purchase.gs example
ensureMaterialHeadersExist(poSheet, PO_COL_FIRST_MATERIAL);
const finalHeaders = poSheet.getRange(...).getValues()[0];
finalHeaders.forEach((matId, index) => {
  newRow[PO_COL_FIRST_MATERIAL + index] = materials[matId] || 0;
});
```

### 4. Google Drive File Upload Structure
Sales order files are organized hierarchically under `PARENT_FOLDER_ID`:
```
Sales/YYYY/MM/DD/<PO_NUMBER>/
  - PO.pdf
  - Invoice.pdf
  - EWay.pdf
```
`uploadFilesToDrive()` handles base64 decoding, folder creation, and returns URL array preserving input order.

### 5. Client-Server Communication
UI uses `google.script.run` for async server calls:
```javascript
// Client-side (index.html)
google.script.run
  .withSuccessHandler(data => { /* handle response */ })
  .withFailureHandler(error => { /* handle error */ })
  .addPurchaseOrder(poData); // calls server function

// Server-side (Purchase.gs)
function addPurchaseOrder(poData) {
  return { success: true, message: "..." }; // or { error: true, message: "..." }
}
```

## Development Workflows

### Local Testing with clasp
```powershell
# Install clasp globally (one-time setup)
npm install -g @google/clasp

# Login and pull existing project
clasp login
clasp clone <SCRIPT_ID>

# Push changes to Apps Script
clasp push

# Open in browser
clasp open
```

### Deployment
1. **Test Deployment**: Apps Script Editor → Deploy → Test deployments (for development)
2. **Production Deployment**: Deploy → New deployment → Select type: Web app
   - Execute as: User accessing the web app
   - Who has access: Anyone (per `appsscript.json` config)

### Debugging
- **Server-side**: Use `Logger.log()` → View in Apps Script Editor → Executions
- **Client-side**: Browser DevTools console for `index.html` JavaScript
- **OTP emails**: Requires authorized sender permissions (OAuth consent screen)

## Key Constraints

1. **No npm/package.json**: This is Apps Script, not Node.js. Dependencies are Google Advanced Services (Drive, DriveActivity) declared in `appsscript.json`.
2. **No ES6 modules**: Use `function` declarations (not `const foo = () => {}`), all functions are global.
3. **Sheet Range API is 1-based**: When using `sheet.getRange()`, add +1 to 0-based constant indices.
4. **OAuth scopes**: Adding new Google APIs requires updating `oauthScopes` in `appsscript.json`.
5. **File size limits**: Apps Script has 50MB script size limit and execution time limits (6 minutes for web apps).

## Common Tasks

### Adding a New Material Column
1. Add row to `Material` sheet (Product ID, Product Name)
2. Run any Purchase/Sales/Manual operation—`ensureMaterialHeadersExist()` auto-appends header
3. `computeAndUpdateStock()` auto-syncs to `Data` sheet on next stock refresh

### Adding New Sheet Fields
1. Update constants in `Constants.gs` (e.g., `NEW_COL_FIELD = 11`)
2. Update `getSheetByName()` header initialization for new sheets
3. Update relevant functions in `Purchase.gs` to read/write new column

### Modifying Email Notifications
Edit `NOTIFY_EMAIL` constant and `sendVehicleEmail()` in `Purchase.gs`. Ensure `https://www.googleapis.com/auth/script.send_mail` scope is present.

## File Organization Logic

- **`Constants.gs`**: Single source of truth for all magic numbers
- **`Code.gs`**: Entry point and sheet utilities (minimal logic)
- **`Purchase.gs`**: All CRUD operations for Purchase/Sales/Manual, stock computation, file handling
- **`Otp.gs`**: Auth-specific logic isolated
- **`index.html`**: Self-contained UI (no external dependencies except TailwindCSS CDN)

## External Dependencies

- **Google Advanced Services**: Drive v3, DriveActivity v2 (enabled in `appsscript.json`)
- **Client-side**: Tailwind CSS via CDN (`https://cdn.tailwindcss.com`)
- **No external libraries**: Pure Apps Script + vanilla JavaScript

## Gotchas

1. **Sheet creation race condition**: `getSheetByName()` creates sheets on-demand with default headers—ensure consistent header order.
2. **Date handling**: Use `Utilities.formatDate()` for consistent timezone formatting (Asia/Kolkata per `appsscript.json`).
3. **File upload format**: Expect `{ name, mimeType, bytes }` where `bytes` is base64 string (decoded in `uploadFilesToDrive`).
4. **OTP expiration**: OTPs are valid only for the day they're generated (compared via `toDateString()`).
5. **Drive folder fallback**: `uploadFilesToDrive` falls back to root folder if `PARENT_FOLDER_ID` is invalid—check logs.
