# CustomToolkit for Google Sheets & ERPNext Integration

## Overview
CustomToolkit is a collection of Google Apps Script modules designed to automate and streamline data workflows between Google Sheets, ERPNext, Supabase, and email-based CSV imports. It provides tools for exporting/importing data, syncing with ERPNext, managing Google Sheets, and more.

## Features
- **ERPNext Integration:**
  - Post items to ERPNext from Google Sheets
  - Fetch item lists and update Google Sheets
  - Password-based and token-based API support
- **Supabase Integration:**
  - Push shipment and SKU data to Supabase tables
- **Email CSV Import:**
  - Automatically import CSV attachments from Gmail into Sheets
- **Sheet Management:**
  - List, copy, move, rename, and delete sheets in bulk
- **Custom Export Dialogs:**
  - Export various templates and reports as CSV via a user-friendly dialog
- **Status Sync:**
  - Update shipment statuses from external sources
- **Backup & Utilities:**
  - Backup sheets to Google Drive
  - Apply dynamic formulas and clear data ranges

## Setup Instructions
1. **Clone or Download the Repository**
2. **Install clasp (if not already):**
   ```
   npm install -g @google/clasp
   ```
3. **Login to clasp:**
   ```
   clasp login
   ```
4. **Push to Google Apps Script:**
   ```
   clasp push
   ```
5. **Set Script Properties:**
   - In the Apps Script editor, go to `Project Settings` > `Script Properties` and add all required API keys, secrets, URLs, and IDs (e.g., `API_KEY`, `API_SECRET`, `baseUrl`, etc.).
   - **Do NOT store secrets in code.**

## Security
- **No sensitive information is stored in the codebase.**
- All API keys, secrets, and credentials are managed via Google Apps Script Properties.
- `.clasp.json` (which contains your scriptId) is ignored by `.gitignore` and not pushed to GitHub.

## File Structure
- `API-Connects.js` – ERPNext API integration
- `upsertDispatch.js` – Supabase integration for shipments/SKUs
- `importCSV.js` – Gmail CSV import automation
- `sheetManipulator.js` – Bulk sheet management
- `updateShipmentStatus.js` – Status sync from external sources
- `Code.js` – Main menu, export dialogs, and utilities
- `Download.html` – Export dialog UI
- `ERPNext-Webhook.js` – Webhook endpoint for ERPNext
- `password_bassed_api.js` – Password-based ERPNext API integration

## Contributing
Pull requests and suggestions are welcome! Please ensure no secrets are added to code or version control.

## License
[MIT](LICENSE) 