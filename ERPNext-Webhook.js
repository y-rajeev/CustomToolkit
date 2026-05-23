// -------------------------------------------------------------------------
// Webhook: Update 'Uploaded Invoice' sheet when Sales Invoice is submitted
// -------------------------------------------------------------------------
var UPLOADED_INVOICE_SHEET_NAME = "Uploaded Invoice";
var WEBHOOK_QUEUE_SHEET_NAME = "Webhook Queue";
var WEBHOOK_QUEUE_BATCH_SIZE = 25;

function doPost(e) {
  var rawPayload = "";
  var invoiceData = null;

  try {
    // --- Safety checks ---
    if (!e || !e.postData || !e.postData.contents) {
      return createResponse("error", "No POST data received");
    }

    rawPayload = e.postData.contents;
    try {
      invoiceData = JSON.parse(rawPayload);
    } catch (parseErr) {
      return createResponse("error", "Invalid JSON: " + parseErr);
    }

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var queueRow = enqueueWebhookPayload_(spreadsheet, invoiceData, rawPayload);

    return createResponse("success", "Invoice queued at row " + queueRow);

  } catch (error) {
    console.error(error);
    // Let ERPNext see a failed delivery if the queue write fails, so it can retry.
    throw error;
  }
}

function processWebhookQueue() {
  return processWebhookQueue_(WEBHOOK_QUEUE_BATCH_SIZE);
}

function installWebhookQueueTrigger() {
  getOrCreateQueueSheet_(SpreadsheetApp.getActiveSpreadsheet());

  var handlerName = "processWebhookQueue";
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === handlerName) {
      return;
    }
  }

  ScriptApp.newTrigger(handlerName)
    .timeBased()
    .everyMinutes(1)
    .create();
}

function enqueueWebhookPayload_(spreadsheet, invoiceData, rawPayload) {
  ensureQueueSheetExists_(spreadsheet);
  var invoiceKey = getInvoiceKey_(invoiceData);
  var appendResult = appendQueueRowWithSheetsApi_(spreadsheet.getId(), [
    new Date(),
    invoiceKey,
    "PENDING",
    "",
    0,
    "",
    rawPayload,
    ""
  ]);

  return getRowNumberFromUpdatedRange_(appendResult.updates.updatedRange);
}

function processWebhookQueue_(maxRows) {
  var lock = LockService.getDocumentLock();
  var lockAcquired = false;
  var processedCount = 0;

  try {
    lockAcquired = lock.tryLock(1000);
    if (!lockAcquired) {
      return 0;
    }

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var queueSheet = getOrCreateQueueSheet_(spreadsheet);
    var uploadedSheet = spreadsheet.getSheetByName(UPLOADED_INVOICE_SHEET_NAME);
    if (!uploadedSheet) {
      throw new Error("Sheet '" + UPLOADED_INVOICE_SHEET_NAME + "' not found");
    }

    var lastRow = queueSheet.getLastRow();
    if (lastRow < 2) {
      return 0;
    }

    var queueRows = queueSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    for (var i = 0; i < queueRows.length && processedCount < maxRows; i++) {
      var sheetRow = i + 2;
      var status = String(queueRows[i][2] || "");
      if (status !== "PENDING" && status !== "ERROR") {
        continue;
      }

      var attempts = Number(queueRows[i][4] || 0) + 1;
      var rawPayload = String(queueRows[i][6] || "");

      try {
        var invoiceData = JSON.parse(rawPayload);
        var invoiceKeys = getInvoiceKeys_(invoiceData);

        if (invoiceKeys.length > 0 && hasInvoiceAlreadyBeenSaved_(uploadedSheet, invoiceKeys)) {
          markQueueRow_(queueSheet, sheetRow, "DUPLICATE", "", attempts, "");
        } else {
          var targetRow = appendUploadedInvoiceRows_(uploadedSheet, [buildUploadedInvoiceRow_(invoiceData)]);
          markQueueRow_(queueSheet, sheetRow, "SAVED", targetRow, attempts, "");
        }
      } catch (rowError) {
        markQueueRow_(queueSheet, sheetRow, "ERROR", "", attempts, String(rowError.message || rowError));
      }

      processedCount++;
    }

    return processedCount;
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

function buildUploadedInvoiceRow_(invoiceData) {
  return [
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
}

function appendUploadedInvoiceRows_(sheet, rows) {
  if (!rows || rows.length === 0) {
    return "";
  }

  var targetRow = Math.max(sheet.getLastRow() + 1, 2);
  sheet.getRange(targetRow, 1, rows.length, rows[0].length).setValues(rows);
  SpreadsheetApp.flush();
  return targetRow;
}

function getOrCreateQueueSheet_(spreadsheet) {
  var queueSheet = spreadsheet.getSheetByName(WEBHOOK_QUEUE_SHEET_NAME);
  if (!queueSheet) {
    queueSheet = spreadsheet.insertSheet(WEBHOOK_QUEUE_SHEET_NAME);
    queueSheet
      .getRange(1, 1, 1, 8)
      .setValues([[
        "received_at",
        "invoice_key",
        "status",
        "target_row",
        "attempts",
        "last_error",
        "payload",
        "processed_at"
      ]]);
  }

  return queueSheet;
}

function ensureQueueSheetExists_(spreadsheet) {
  var queueSheet = spreadsheet.getSheetByName(WEBHOOK_QUEUE_SHEET_NAME);
  if (!queueSheet) {
    throw new Error("Sheet '" + WEBHOOK_QUEUE_SHEET_NAME + "' not found. Run installWebhookQueueTrigger once.");
  }
}

function appendQueueRowWithSheetsApi_(spreadsheetId, rowValues) {
  var range = encodeURIComponent("'" + WEBHOOK_QUEUE_SHEET_NAME + "'!A:H");
  var url = "https://sheets.googleapis.com/v4/spreadsheets/" +
    encodeURIComponent(spreadsheetId) +
    "/values/" +
    range +
    ":append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS";

  var response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    payload: JSON.stringify({
      values: [rowValues]
    }),
    muteHttpExceptions: true
  });

  var responseCode = response.getResponseCode();
  var responseBody = response.getContentText();
  if (responseCode < 200 || responseCode >= 300) {
    throw new Error("Queue append failed: HTTP " + responseCode + " " + responseBody);
  }

  return JSON.parse(responseBody);
}

function getRowNumberFromUpdatedRange_(updatedRange) {
  var match = String(updatedRange || "").match(/![A-Z]+(\d+):/);
  return match ? Number(match[1]) : "";
}

function markQueueRow_(queueSheet, rowNumber, status, targetRow, attempts, errorMessage) {
  queueSheet
    .getRange(rowNumber, 3, 1, 4)
    .setValues([[status, targetRow || "", attempts || 0, errorMessage || ""]]);
  queueSheet
    .getRange(rowNumber, 8)
    .setValue(status === "ERROR" ? "" : new Date());
}

function getInvoiceKey_(invoiceData) {
  return getInvoiceKeys_(invoiceData).join(" | ");
}

function getInvoiceKeys_(invoiceData) {
  if (!invoiceData) {
    return [];
  }

  var keys = [invoiceData.name, invoiceData.vch_no, invoiceData.shipment_id]
    .filter(function(value) {
      return value !== null && value !== undefined && String(value).trim() !== "";
    })
    .map(function(value) {
      return String(value);
    });

  return keys.filter(function(value, index) {
    return keys.indexOf(value) === index;
  });
}

function hasInvoiceAlreadyBeenSaved_(sheet, invoiceKeys) {
  if (!invoiceKeys || invoiceKeys.length === 0) {
    return false;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return false;
  }

  var voucherValues = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  var shipmentValues = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
  for (var i = 0; i < voucherValues.length; i++) {
    if (
      invoiceKeys.indexOf(String(voucherValues[i][0])) !== -1 ||
      invoiceKeys.indexOf(String(shipmentValues[i][0])) !== -1
    ) {
      return true;
    }
  }

  return false;
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
