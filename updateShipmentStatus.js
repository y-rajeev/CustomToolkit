// ----------------------------------------------------------------------------------------------
// Update Status of shipment in "Uploaded Invoice" sheet from auto calls "Dispatch Report" sheet
// ----------------------------------------------------------------------------------------------

function updateStatusFrmExtSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const uploadedInvoiceSheet = ss.getSheetByName("Uploaded Invoice")

  // Access external sheet by url
  const externalSheetID = PropertiesService.getScriptProperties().getProperties('externalSheetID');
  const externalFile = SpreadsheetApp.openById(externalSheetID);
  const externalSheet = externalFile.getSheetByName("Dispatched Report");

  // Get data from Uploaded Invoice sheet
  const uploadedInvoiceData = uploadedInvoiceSheet.getDataRange().getValues();
  const uploadedInvoiceHeader = uploadedInvoiceData[0];
  const shipmentNameIndex = uploadedInvoiceHeader.indexOf("shipment_id");
  const statusIndex = uploadedInvoiceHeader.indexOf("status");

  // Get data from Auto-Calls sheet
  const externalSheetData = externalSheet.getDataRange().getValues();

  // Create a map of shipment_id to status from Auto-Calls sheet
  const shipmentStatusMap = {};
  for (let i = 1; i < externalSheetData.length; i++) {
    const [ , shipmentName, , , , , , status] = externalSheetData[i];
    if (shipmentName) {
      shipmentStatusMap[shipmentName] = status;
    }
  }

  // Iterate through the Uploaded Invoice sheet and update the status
  for (let i = 1; i < uploadedInvoiceData.length; i++) {
    const shipmentName = uploadedInvoiceData[i][shipmentNameIndex];
    const currentStatus = uploadedInvoiceData[i][statusIndex];

    if (shipmentName in shipmentStatusMap) {
      // Shipment found in Auto-Calls sheet, check for status update
      const updatedStatus = shipmentStatusMap[shipmentName];
      if (currentStatus !== updatedStatus) {
        uploadedInvoiceSheet.getRange(i + 1, statusIndex + 1).setValue(updatedStatus);
      }
    } else {
      // Shipment not found in Auto-Calls Sheet, set status to "Closed" if not already
      if (currentStatus !== "Closed") {
        uploadedInvoiceSheet.getRange(i+1, statusIndex + 1).setValue("Closed");
      }
    }
  }
  SpreadsheetApp.flush();
  Logger.log("Status update completed!");
}