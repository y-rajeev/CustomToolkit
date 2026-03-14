// -------------------------------------------------------------------------
// Webhook: Update 'Uploaded Invoice' sheet when Sales Invoice is submitted
// -------------------------------------------------------------------------
function doPost(e) {
  var lock = LockService.getScriptLock();

  try {
    // Wait up to 30 seconds to get the lock
    lock.waitLock(30000);

    // --- Safety checks ---
    if (!e || !e.postData || !e.postData.contents) {
      return createResponse("error", "No POST data received");
    }

    var invoiceData;
    try {
      invoiceData = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return createResponse("error", "Invalid JSON: " + parseErr);
    }

    // Open spreadsheet & sheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // Or: SpreadsheetApp.openById("SPREADSHEET_ID");
    var sheet = spreadsheet.getSheetByName("Uploaded Invoice");
    if (!sheet) {
      throw new Error("Sheet 'Uploaded Invoice' not found");
    }

    // Build row data (Column A left blank)
    var newRow = [
      "",                                       // Column A
      invoiceData.creation || "",              // Column B
      invoiceData.vch_no || "",                // Column C
      invoiceData.shipment_id || "",           // Column D
      invoiceData.channel_abb || "",           // Column E
      invoiceData.mode || "",                  // Column F
      invoiceData.branch || "",                // Column G
      invoiceData.dispatch_date || "",         // Column H
      invoiceData.eta_date || "",              // Column I
      invoiceData.repository || "",            // Column J
      invoiceData.status || "",                // Column K
      invoiceData.total_qty || "",             // Column L
      invoiceData.net_total || "",             // Column M
      invoiceData.total_taxes_and_charges || "", // Column N
      invoiceData.grand_total || ""            // Column O
    ];

    // Append the row to the sheet
    sheet.appendRow(newRow);

    // Success response
    return createResponse("success", "Invoice row added");

  } catch (error) {
    console.error(error);
    return createResponse("error", String(error.message || error));

  } finally {
    try {
      lock.releaseLock();
    } catch (e2) {
      // ignore
    }
  }
}

// Helper: build JSON response
function createResponse(status, message) {
  var payload = {
    status: status,
    message: message
  };
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
