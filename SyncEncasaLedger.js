// ------------------------------------------------------
// Sync all shipment data from Encasa Order Ledger sheet
// ------------------------------------------------------

function importSpecificColumns() {
    try {
      const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Encasa Ledger");
      const sourceSpreadsheetId = PropertiesService.getScriptProperties().getProperties("sourceSpreadsheetId")
      const sourceSheet = SpreadsheetApp.openById(sourceSpreadsheetId).getSheetByName("Order Ledger - Encasa");
  
      // Get last row to avoid scanning unnecessary empty rows
      const lastRow = sourceSheet.getLastRow();
      if (lastRow === 0) {
        Logger.log("Source sheet has no data.");
        return;
      }
  
      // Fetch entire data in one go (Columns A to Q)
      const data = sourceSheet.getRange(1, 1, lastRow, 17).getValues();
  
      // Filter out rows where Column B (index 1) is empty
      const filteredData = data.filter(row => row[1] !== "" && row[1] !== null)
                               .map(row => [row[0], row[1], row[2], row[7], row[16]]); // Extract A, B, C, H, Q
  
      if (filteredData.length === 0) {
        Logger.log("No valid data found.");
        return;
      }
  
      // Clear old data & Write all at once (Super fast)
      targetSheet.clear();
      targetSheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
  
      Logger.log(`Successfully imported ${filteredData.length} rows.`);
    } catch (error) {
      Logger.log("Error importing data: " + error.message);
    }
  }