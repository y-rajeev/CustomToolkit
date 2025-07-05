// -----------------------
// ERP Uloaded Shipment
// -----------------------

function pushShipmentToSupabase() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Uploaded Invoice");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const supabaseUrl = PropertiesService.getScriptProperties().getProperty("URL");
  const supabaseKey = PropertiesService.getScriptProperties().getProperty("SERVICE_KEY");

  const insertData = [];

  rows.forEach((row) => {
    const rowObj = {};
    headers.forEach((header, i) => {
      rowObj[header] = row[i];
    });

    const shipmentId = rowObj["shipment_id"];
    if (!shipmentId) return;  // Skip rows without shipment_id

    insertData.push({
      input_timestamp: new Date().toISOString(),  // Add current timestamp
      shipment_id: shipmentId,
      channel_abb: rowObj["channel_abb"] || null,
      mode: rowObj["mode"] || null,
      branch: rowObj["branch"] || null,
      dispatch_date: rowObj["dispatch_date"] ? formatDate(rowObj["dispatch_date"]) : null,
      eta_date: rowObj["eta_date"] ? formatDate(rowObj["eta_date"]) : null,
      repository: rowObj["repository"] || null,
      status: rowObj["status"] || null,
      total_qty: parseInt(rowObj["total_qty"]) || null,
      net_total: parseInt(rowObj["net_total"]) || null,
      total_taxes_and_charges: parseFloat(rowObj["total_taxes_and_charges"]) || null,
      grand_total: parseFloat(rowObj["grand_total"]),
      month: rowObj["month"] ? formatDate(rowObj["month"]) : null,
    });
  });

  if (insertData.length === 0) {
    Logger.log("No valid shipment data found to upsert.");
    return;
  }

  const options = {
    method: "POST",
    contentType: "application/json",
    headers: {
      apikey: supabaseKey,
      Authorization: `Bearer ${supabaseKey}`,
      Prefer: "resolution=merge-duplicates"
    },
    payload: JSON.stringify(insertData),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(`${supabaseUrl}/rest/v1/tab_shipment_meta?on_conflict=shipment_id`, options);
    Logger.log("Response code: " + response.getResponseCode());
    Logger.log("Response body: " + response.getContentText());
  } catch (err) {
    Logger.log("Error posting to Supabase: " + err);
  }
}

function formatDate(dateValue) {
  // Convert from Google Sheets date to YYYY-MM-DD format
  if (Object.prototype.toString.call(dateValue) === '[object Date]') {
    return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return null;
}


// -----------------
// SKU Mapping
// -----------------

function pushSkuToSupabase() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SKU Mapping");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const supabaseUrl = PropertiesService.getScriptProperties().getProperty("URL");
  const supabaseKey = PropertiesService.getScriptProperties().getProperty("SERVICE_KEY");

  const insertData = [];

  rows.forEach((row) => {
    const rowObj = {};
    headers.forEach((header, i) => {
      rowObj[header] = row[i];
    });

    const skuId = rowObj["SKU"];
    if (!skuId) return;  // Skip rows without skuId

    insertData.push({
      sku: skuId,
      line: rowObj["Line"] || null,
      design: rowObj["Color/Design"] || null,
      size: rowObj["Size"] || null,
      pcs_pack: rowObj["Pcs / pack"] || null,
      product: rowObj["Product"] || null,
      production: rowObj["Production"] || null
    });
  });

  if (insertData.length === 0) {
    Logger.log("No valid sku data found to upsert.");
    return;
  }

  const options = {
    method: "POST",
    contentType: "application/json",
    headers: {
      apikey: supabaseKey,
      Authorization: `Bearer ${supabaseKey}`,
      Prefer: "resolution=merge-duplicates"
    },
    payload: JSON.stringify(insertData),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(`${supabaseUrl}/rest/v1/tab_sku_mapping?on_conflict=sku`, options);
    Logger.log("Response code: " + response.getResponseCode());
    Logger.log("Response body: " + response.getContentText());
  } catch (err) {
    Logger.log("Error posting to Supabase: " + err);
  }
}