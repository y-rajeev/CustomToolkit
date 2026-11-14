# CustomToolkit

Google Apps Script tools for Google Sheets workflows, ERPNext synchronization, and spreadsheet administration.

## Modules

| File | Purpose |
| --- | --- |
| `spreadsheetMenusAndUtilities.js` | Spreadsheet menus, CSV export, formula helpers, and order-sheet utilities |
| `exportDialog.html` | UI for exporting a sheet range as CSV |
| `erpNextItems.js` | Create ERPNext items and sync the Item List report |
| `erpNextShipmentSync.js` | Sync shipment report rows into Uploaded Invoice |
| `erpNextInvoiceWebhook.js` | Queue and process ERPNext invoice webhook requests |
| `generateBOMUploader.js` | Validate BOM data and generate ERPNext BOM import sheets |
| `shipmentStatusSync.js` | Update shipment status from the configured dispatch spreadsheet |
| `sheetManager.js` | List, copy, move, rename, and delete sheets using the Sheet Manipulator tab |
| `spreadsheetBackup.js` | Keep a dated Drive backup of the active spreadsheet |
| `dateTimeUtils.js` | Shared timestamp formatting used by sync modules |

Apps Script files share one global namespace. Keep top-level function names unique, and use a trailing underscore for private helpers.

## Configuration

Set these values in **Apps Script → Project Settings → Script Properties**:

| Property | Used for |
| --- | --- |
| `ERPNext_API_KEY` | ERPNext token authentication |
| `ERPNext_API_SECRET` | ERPNext token authentication |
| `ERPNext_URL` | ERPNext site base URL |
| `externalSheetID` | Spreadsheet used by shipment status sync |
| `backupFolderId` | Google Drive folder used for spreadsheet backups |

`generateBOMUploader.js` currently targets the spreadsheet ID in its `SPREADSHEET_ID` constant. Update it if the BOM workbook changes.

Sync timestamps are stored in Script Properties as `last_fetch_timestamp` and `last_synced_at`.

Do not commit API keys, secrets, or credentials. `.clasp.json` contains the Apps Script project ID and is intentionally excluded from Git.

## Development

1. Install clasp: `npm install -g @google/clasp`.
2. Authenticate: `clasp login`.
3. Pull the linked Apps Script project: `clasp pull`.
4. Push local source changes: `clasp push`.

The project manifest is `appsscript.json`. The spreadsheet-bound menus are created by `onOpen()` in `spreadsheetMenusAndUtilities.js`.
