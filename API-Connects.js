// ----------------------------------------------------
// API Post Method (Add item in ERPNext)
// ----------------------------------------------------
function postItemToERPNext() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('API POST');
  
  if (!sheet) {
    Logger.log("Error: Sheet 'API POST' not found");
    return;
  }

  const config = {
    apiKey: PropertiesService.getScriptProperties().getProperty('API_KEY'),
    apiSecret: PropertiesService.getScriptProperties().getProperty('API_SECRET'),
    baseUrl: PropertiesService.getScriptProperties().getProperties('baseUrl'),
    logDetails: false
  };

  if (!config.apiKey || !config.apiSecret) {
    Logger.log("Error: API credentials not found in Script Properties");
    return;
  }

  // Determine the range dynamically
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  if (lastRow < 2) {
    Logger.log("Error: No data available for posting");
    return;
  }

  // Get headers from row 2
  var headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  
  // Find existing Status column or create new one
  var statusColumnIndex = headers.findIndex(header => header === "Status");
  if (statusColumnIndex === -1) {
    // Status column doesn't exist, create it
    statusColumnIndex = lastCol + 1;
    sheet.getRange(2, statusColumnIndex).setValue("Status");
  } else {
    // Status column exists, convert from 0-based to 1-based index
    statusColumnIndex += 1;
  }

  // Get all data starting from row 3
  var allData = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
  
  // Filter out completely empty rows
  var data = allData.filter(row => {
    return row.some(cell => cell !== "" && cell !== null);
  });
  
  if (data.length === 0) {
    Logger.log("No valid data rows found to process");
    return;
  }
  
  Logger.log(`Found ${data.length} valid rows to process`);

  // Process each non-empty row
  data.forEach((row, index) => {
    var actualRowNumber = allData.findIndex(r => r.every((cell, i) => cell === row[i])) + 3;
    
    // Build payload object from headers and row data
    var payload = {};
    var taxes = [];

    headers.forEach((header, colIndex) => {
      if (header && row[colIndex]) {
        let value = row[colIndex].toString().trim();
        header = header.trim();
        
        if (header.toLowerCase() === 'item_tax_template') {
          if (value) {
            payload.item_tax_template = value;
            taxes.push({
              item_tax_template: value,
              tax_category: ""
            });
          }
        } else {
          payload[header] = value;
        }
      }
    });

    if (taxes.length > 0) {
      payload.taxes = taxes;
    }

    var options = {
      method: "POST",
      contentType: "application/json",
      headers: {
        "Authorization": "token " + config.apiKey + ":" + config.apiSecret,
        "Accept": "application/json"
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    // Simplified logging
    if (config.logDetails) {
      Logger.log(`Processing row ${actualRowNumber}: ${payload.item_code}`);
    }

    try {
      var response = UrlFetchApp.fetch(config.baseUrl + '/api/resource/Item', options);
      var responseCode = response.getResponseCode();
      var responseData = JSON.parse(response.getContentText());

      if (responseCode === 200 || responseCode === 201) {
        let status = "Success";
        sheet.getRange(actualRowNumber, statusColumnIndex).setValue(status);
        
        // Simplified success logging
        Logger.log(`✓ Row ${actualRowNumber}: Created ${payload.item_code} | ${payload.item_name}`);
        
        // Log tax template info
        if (payload.item_tax_template) {
          Logger.log(`  └─ Tax Template: ${payload.item_tax_template}`);
        }

      } else {
        let errorMessage = responseData.exception ? 
          responseData.exception.split(":").pop().trim() : 
          `Error ${responseCode}`;
        
        sheet.getRange(actualRowNumber, statusColumnIndex).setValue("Failed: " + errorMessage);
        Logger.log(`✗ Row ${actualRowNumber}: Failed to create ${payload.item_code} - ${errorMessage}`);
      }

    } catch (error) {
      let errorMessage = "Error: " + error.message;
      sheet.getRange(actualRowNumber, statusColumnIndex).setValue(errorMessage);
      Logger.log(`✗ Row ${actualRowNumber}: Error processing ${payload.item_code} - ${error.toString()}`);
    }

    Utilities.sleep(1000);
  });
  
  // Final summary
  Logger.log("\n=== Summary ===");
  Logger.log(`Total rows processed: ${data.length}`);
  Logger.log("Processing completed");
}

// ------------------------------------
// API Calls to fetch uploaded item
// ------------------------------------

function fetchERPItem() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('API_KEY');
  var apiSecret = PropertiesService.getScriptProperties().getProperty('API_SECRET');

  var baseUrl = PropertiesService.getScriptProperties().getProperties('baseUrl');
  var reportName = "Item List";

  var headers = {
    "Authorization": "token " + apiKey + ":" + apiSecret,
    "Content-Type": "application/json"
  };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Item List");

  // Fetch the last fetch timestamp from script properties (if available)
  var lastFetchTimestamp = PropertiesService.getScriptProperties().getProperty('last_fetch_timestamp');

  // Make the API request to fetch data
  var options = {
    "method": "GET",
    "headers": headers,
    "muteHttpExceptions": true
  };

  var response = UrlFetchApp.fetch(
    baseUrl + "/api/method/frappe.desk.query_report.run?report_name=" + reportName,
    options
  );
  var responseData = JSON.parse(response.getContentText());

  // Check if the request was successful
  if (response.getResponseCode() === 200) {
    var data = responseData.message.result;
    var columns = responseData.message.columns;

    // Extract headers from columns
    var headers = columns.map(function(col) {
      return col.label;
    });

    // Filter out the "Total" row and invalid rows
    var filteredData = data.filter(function(row) {
      return row.input_timestamp !== "Total" && typeof row === 'object';
    });

    // Filter rows only if the last fetch timestamp exists
    var filteredRows = lastFetchTimestamp
      ? filteredData.filter(function(row) {
          return new Date(row.input_timestamp) > new Date(lastFetchTimestamp);
        })
      : filteredData; // Fetch all data if no timestamp exists

    // Prepare rows for writing to the sheet
    var displayRows = filteredRows.map(function(row) {
      return [
        row.input_timestamp,
        row.sku,
        row.item_name,
        row.description,
        row.item_group,
        row.quality,
        row.stock_uom,
        row.gst_hsn_code,
        row.item_tax_template,
        row.is_purchase_item,
        row.is_sales_item,
        row.is_sub_contracted_item,
        row.include_item_in_manufacturing
      ];
    });

    // Get the last row with actual content (ignoring any empty rows)
    var lastRow = sheet.getRange("B:B").getValues().filter(String).length;

    // Append the filtered rows starting directly after the last row with data
    if (displayRows.length > 0) {
      sheet.getRange(lastRow + 1, 2, displayRows.length, displayRows[0].length).setValues(displayRows);
    }

    // Store the current timestamp in IST format as the new 'last_fetch_timestamp'
    var currentISTTimestamp = getCurrentISTTimestamp();
    PropertiesService.getScriptProperties().setProperty('last_fetch_timestamp', currentISTTimestamp);
  } else {
    Logger.log("Unable to fetch data. Response code: " + response.getResponseCode());
  }
}

// Function to get the current timestamp in IST
function getCurrentISTTimestamp() {
  var now = new Date(); // Current UTC time
  var istOffset = 5.5 * 60 * 60 * 1000; // IST offset in milliseconds (UTC + 5:30)
  var istTime = new Date(now.getTime() + istOffset);
  return istTime.toISOString().replace('T', ' ').replace('Z', '');
}