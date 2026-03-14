// --------------------------------------
// API Calls to fetch uploaded shipment
// --------------------------------------

function fetchUploadedShipment() {
  var props      = PropertiesService.getScriptProperties();
  var apiKey     = props.getProperty('ERPNext_API_KEY');
  var apiSecret  = props.getProperty('ERPNext_API_SECRET');
  var baseUrl    = props.getProperty('ERPNext_URL');
  var reportName = 'Shipment List';
  var lastSyncAt = props.getProperty('last_synced_at');

  if (!apiKey || !apiSecret || !baseUrl) {
    throw new Error('Missing ERPNext credentials or base URL');
  }

  var requestHeaders = {
    Authorization: 'token ' + apiKey + ':' + apiSecret,
    'Content-Type': 'application/json'
  };

  var sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Uploaded Invoice');

  if (!sheet) {
    throw new Error('Sheet "Uploaded Invoice" not found');
  }

  var options = {
    method: 'GET',
    headers: requestHeaders,
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(
    baseUrl + "/api/method/frappe.desk.query_report.run?report_name=" + reportName,
    options
  );

  if (response.getResponseCode() !== 200) {
    Logger.log(response.getContentText());
    throw new Error('ERPNext API failed');
  }

  var responseData = JSON.parse(response.getContentText());

  var data = responseData.message.result || [];

  // Remove total row & invalid rows
  var filteredData = data.filter(function(row) {
    return row.input_timestamp && row.input_timestamp !== 'Total';
  });

  // Apply incremental sync
  var filteredRows = lastSyncAt
    ? filteredData.filter(function(row) {
        return new Date(row.input_timestamp) > new Date(lastSyncAt);
      })
    : filteredData;

  if (filteredRows.length === 0) {
    Logger.log('No new rows to sync');
    return;
  }

  var displayRows = filteredRows.map(function(row) {
    return [
      row.input_timestamp,
      row.vch_no,
      row.shipment_id,
      row.channel_abb,
      row.mode,
      row.branch,
      row.dispatch_date,
      row.eta_date,
      row.repository,
      row.status,
      row.total_qty,
      row.net_total,
      row.total_taxes_and_charges,
      row.grand_total
    ];
  });

  var lastRow = sheet
    .getRange('B:B')
    .getValues()
    .filter(function(r) {
      return r[0] !== '';
    }).length;

  sheet
    .getRange(lastRow + 1, 2, displayRows.length, displayRows[0].length)
    .setValues(displayRows);

  props.setProperty('last_synced_at', getCurrentISTTimestamp());
}

// --------------------------------------
// Utility: IST Timestamp
// --------------------------------------

function getCurrentISTTimestamp() {
  var now = new Date();
  var istOffset = 5.5 * 60 * 60 * 1000;
  return new Date(now.getTime() + istOffset)
    .toISOString()
    .replace('T', ' ')
    .replace('Z', '');
}
