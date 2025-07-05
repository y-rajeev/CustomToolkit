// -------------------------------------------------------------------------
// Get instant update in 'Uploaded Invoice' sheet after sales invoice submit
// -------------------------------------------------------------------------
function doPost(e) {
  // Open the spreadsheet and get the "Uploaded Invoice" sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Uploaded Invoice");
  
  try {
    // Parse the incoming JSON data
    var invoiceData = JSON.parse(e.postData.contents);
    
    // Extract necessary fields from the incoming data
    var newRow = [
      invoiceData.creation,                   // Column B: doc.posting_date
      invoiceData.vch_no,                     // Column C: doc.name
      invoiceData.shipment_id,                // Column D: doc.custom_shipment_id
      invoiceData.channel_abb,                // Column E: doc.custom_destination
      invoiceData.mode,                       // Column F: doc.custom_shipment_mode
      invoiceData.branch,                     // Column G: doc.branch
      invoiceData.dispatch_date,              // Column H: doc.custom_shipment_date
      invoiceData.eta_date,                   // Column I: doc.custom_eta_date
      invoiceData.repository,                 // Column J: doc.custom_repository
      invoiceData.status,                     // Column K: doc.custom_inbound_status
      invoiceData.total_qty,                  // Column L: doc.total_qty
      invoiceData.net_total,                  // Column M: doc.net_total
      invoiceData.total_taxes_and_charges,    // Column N: doc.total_taxes_and_charges
      invoiceData.grand_total                 // Column O: doc.grand_total
    ];
    
    // Determine the last row to append new data below existing data
    var lastRow = sheet.getLastRow();
    
    // Set the values starting from column B
    sheet.getRange(lastRow + 1, 2, 1, newRow.length).setValues([newRow]);
    
    return ContentService.createTextOutput(JSON.stringify({status: "success"}))
                         .setMimeType(ContentService.MimeType.JSON);
  
  } catch (error) {
    // Handle any errors that occur during the operation
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: error.message}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}