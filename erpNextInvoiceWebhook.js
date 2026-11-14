
var UPLOADED_INVOICE_SHEET_NAME = "Uploaded Invoice";
var WEBHOOK_QUEUE_SHEET_NAME = "Webhook Queue";
var WEBHOOK_QUEUE_BATCH_SIZE = 25;

var INVOICE_FIELD_ORDER = [
  "creation",
  "vch_no",
  "shipment_id",
  "channel_abb",
  "mode",
  "branch",
  "dispatch_date",
  "eta_date",
  "repository",
  "status",
  "total_qty",
  "net_total",
  "total_taxes_and_charges",
  "grand_total"
];


// -------------------------------------------------------------------------
// 1. Webhook Entry Point
// -------------------------------------------------------------------------

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createResponse("error", "No POST data received");
    }

    var invoiceData = JSON.parse(e.postData.contents);
    var orderedData = normalizeInvoiceData_(invoiceData);
    var rawPayload = JSON.stringify(orderedData);

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    var queueRow = enqueueWebhookPayload_(
      spreadsheet,
      orderedData,
      rawPayload
    );

    return createResponse("success", "Invoice queued at row " + queueRow);

  } catch (error) {
    console.error(error);
    throw error;
  }
}


// -------------------------------------------------------------------------
// 2. Normalize JSON Field Order
// -------------------------------------------------------------------------

function normalizeInvoiceData_(data) {
  if (!data || typeof data !== "object" || Array.isArray(data)) {
    throw new Error("Invalid invoice JSON object");
  }

  var ordered = {};

  INVOICE_FIELD_ORDER.forEach(function(key) {
    if (Object.prototype.hasOwnProperty.call(data, key)) {
      ordered[key] = data[key];
    }
  });

  Object.keys(data).forEach(function(key) {
    if (!Object.prototype.hasOwnProperty.call(ordered, key)) {
      ordered[key] = data[key];
    }
  });

  return ordered;
}


// -------------------------------------------------------------------------
// 3. Queue Management
// -------------------------------------------------------------------------

function getOrCreateQueueSheet_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(WEBHOOK_QUEUE_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(WEBHOOK_QUEUE_SHEET_NAME);

    sheet.getRange(1, 1, 1, 8).setValues([[
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

  return sheet;
}


function enqueueWebhookPayload_(spreadsheet, invoiceData, rawPayload) {
  var queueSheet = getOrCreateQueueSheet_(spreadsheet);

  var result = appendQueueRowWithSheetsApi_(
    spreadsheet.getId(),
    [
      new Date(),
      getInvoiceKey_(invoiceData),
      "PENDING",
      "",
      0,
      "",
      rawPayload,
      ""
    ]
  );

  return getRowNumberFromUpdatedRange_(result.updates.updatedRange);
}


// -------------------------------------------------------------------------
// 4. Google Sheets API Queue Append
// -------------------------------------------------------------------------

function appendQueueRowWithSheetsApi_(spreadsheetId, rowValues) {
  var range = encodeURIComponent("'" + WEBHOOK_QUEUE_SHEET_NAME + "'!A:H");

  var url =
    "https://sheets.googleapis.com/v4/spreadsheets/" +
    encodeURIComponent(spreadsheetId) +
    "/values/" + range +
    ":append?valueInputOption=RAW&insertDataOption=INSERT_ROWS";

  var response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    payload: JSON.stringify({ values: [rowValues] }),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();

  if (code < 200 || code >= 300) {
    throw new Error(
      "Queue append failed: HTTP " + code + " " + response.getContentText()
    );
  }

  return JSON.parse(response.getContentText());
}


function getRowNumberFromUpdatedRange_(updatedRange) {
  var match = String(updatedRange || "").match(/![A-Z]+(\d+):/);
  return match ? Number(match[1]) : "";
}


// -------------------------------------------------------------------------
// 5. Queue Processing Trigger
// -------------------------------------------------------------------------

function processWebhookQueue() {
  return processWebhookQueue_(WEBHOOK_QUEUE_BATCH_SIZE);
}


function installWebhookQueueTrigger() {
  getOrCreateQueueSheet_(SpreadsheetApp.getActiveSpreadsheet());

  var handler = "processWebhookQueue";
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === handler) {
      return;
    }
  }

  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyMinutes(1)
    .create();
}


// -------------------------------------------------------------------------
// 6. Process Pending Webhooks
// -------------------------------------------------------------------------

function processWebhookQueue_(maxRows) {
  var lock = LockService.getDocumentLock();

  if (!lock.tryLock(1000)) {
    return 0;
  }

  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var queueSheet = getOrCreateQueueSheet_(spreadsheet);
    var uploadedSheet = spreadsheet.getSheetByName(
      UPLOADED_INVOICE_SHEET_NAME
    );

    if (!uploadedSheet) {
      throw new Error("Uploaded Invoice sheet not found");
    }

    var lastRow = queueSheet.getLastRow();

    if (lastRow < 2) {
      return 0;
    }

    var rows = queueSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    var processed = 0;

    // Build existing invoice lookup once per batch.
    var existingInvoices = getExistingInvoiceKeys_(uploadedSheet);

    for (var i = 0; i < rows.length && processed < maxRows; i++) {
      var row = rows[i];
      var rowNumber = i + 2;
      var status = String(row[2] || "");

      if (status !== "PENDING" && status !== "ERROR") {
        continue;
      }

      var attempts = Number(row[4] || 0) + 1;

      try {
        var invoiceData = JSON.parse(String(row[6] || ""));
        var invoiceKeys = getInvoiceKeys_(invoiceData);

        var duplicate = invoiceKeys.some(function(key) {
          return existingInvoices[key] === true;
        });

        if (duplicate) {
          markQueueRow_(
            queueSheet, rowNumber, "DUPLICATE", "", attempts, ""
          );
        } else {
          var targetRow = appendUploadedInvoiceRows_(
            uploadedSheet,
            [buildUploadedInvoiceRow_(invoiceData)]
          );

          markQueueRow_(
            queueSheet, rowNumber, "SAVED", targetRow, attempts, ""
          );

          // Prevent duplicates within the same processing batch.
          invoiceKeys.forEach(function(key) {
            existingInvoices[key] = true;
          });
        }

      } catch (error) {
        markQueueRow_(
          queueSheet,
          rowNumber,
          "ERROR",
          "",
          attempts,
          String(error.message || error)
        );
      }

      processed++;
    }

    return processed;

  } finally {
    lock.releaseLock();
  }
}


// -------------------------------------------------------------------------
// 7. Uploaded Invoice Operations
// -------------------------------------------------------------------------

function buildUploadedInvoiceRow_(data) {
  return [
    "",
    data.creation || "",
    data.vch_no || "",
    data.shipment_id || "",
    data.channel_abb || "",
    data.mode || "",
    data.branch || "",
    data.dispatch_date || "",
    data.eta_date || "",
    data.repository || "",
    data.status || "",
    data.total_qty || "",
    data.net_total || "",
    data.total_taxes_and_charges || "",
    data.grand_total || ""
  ];
}


function appendUploadedInvoiceRows_(sheet, rows) {
  var targetRow = Math.max(sheet.getLastRow() + 1, 2);

  sheet.getRange(targetRow, 1, rows.length, rows[0].length)
    .setValues(rows);

  return targetRow;
}


// -------------------------------------------------------------------------
// 8. Duplicate Detection
// -------------------------------------------------------------------------

function getInvoiceKeys_(data) {
  var keys = [data.vch_no, data.shipment_id];

  return keys.filter(function(value, index) {
    return (
      value !== null &&
      value !== undefined &&
      String(value).trim() !== "" &&
      keys.indexOf(value) === index
    );
  }).map(function(value) {
    return String(value);
  });
}


function getInvoiceKey_(data) {
  return getInvoiceKeys_(data).join(" | ");
}


function getExistingInvoiceKeys_(sheet) {
  var existing = {};
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return existing;
  }

  var values = sheet.getRange(2, 3, lastRow - 1, 2).getValues();

  values.forEach(function(row) {
    row.forEach(function(value) {
      if (value !== null && value !== undefined && String(value).trim()) {
        existing[String(value)] = true;
      }
    });
  });

  return existing;
}


// -------------------------------------------------------------------------
// 9. Update Queue Status
// -------------------------------------------------------------------------

function markQueueRow_(sheet, row, status, targetRow, attempts, error) {
  sheet.getRange(row, 3, 1, 4).setValues([[
    status,
    targetRow || "",
    attempts,
    error || ""
  ]]);

  sheet.getRange(row, 8).setValue(
    status === "ERROR" ? "" : new Date()
  );
}


// -------------------------------------------------------------------------
// 10. JSON Response
// -------------------------------------------------------------------------

function createResponse(status, message) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: status,
      message: message
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
