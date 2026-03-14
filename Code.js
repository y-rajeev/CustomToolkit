function onOpen() {
    var ui = SpreadsheetApp.getUi();
    
    ui.createMenu('Export')
        .addItem('Export Item Template', 'openDialogForNewItemTemplate')
        .addItem('Export Sales Order', 'openDialogForSalesOrder')
        .addItem('Export Sales Invoice', 'openDialogForSalesInvoice')
        .addItem('Export Auto Invoicing', 'openDialogForAutoInvoicing')
        .addItem('Export Schedule Template', 'openDialogForScheduleTemplate')
        .addItem('Export Stock Entry', 'openDialogForStockEntry')
        .addItem('Export Price List', 'openDialogForPriceList')
        .addToUi();

    ui.createMenu('ERP Sync')
        .addItem('Sync Items', 'fetchERPItem')
        .addItem('Sync Shipment', 'fetchUploadedShipment')
        .addItem('Push Item', 'postItemToERPNext')
        .addToUi();

    // ui.createMenu('Sheet Sync')
    //     .addItem('Refresh Order Ledger', 'importSpecificColumns')
    //     .addToUi();
}

function openDialogForNewItemTemplate() {
    openDialog('Item Template');
}

function openDialogForSalesOrder() {
    openDialog('ERP import - Sales Order');
}

function openDialogForSalesInvoice() {
    openDialog('ERP Import - Sales Invoice');
}


function openDialogForScheduleTemplate() {
    openDialog('Amazon - Schedule Template');
}

function openDialogForStockEntry() {
  openDialog('Stock Entry');
}

function openDialogForAutoInvoicing() {
  openDialog('Auto Invoicing');
} 

function openDialogForPriceList() {
  openDialog('erp_price_list');
}

// Open a modal dialog with a custom title and content from 'Download.html'
function openDialog(sheetName) {
    // Create HTML output for the dialog
    var html = HtmlService.createHtmlOutputFromFile('Download')
        .setWidth(300)
        .setHeight(150);
    
    // Set the sheetName in a hidden input field within the HTML
    html.append('<script>document.getElementById("sheetName").value = "' + sheetName + '";</script>');
    
    // Set the title of the dialog
    var dialogTitle = 'Exporting - ' + sheetName;
    
    // Show the dialog to the user
    SpreadsheetApp.getUi().showModalDialog(html, dialogTitle);
}

function exportToCSV(sheetName, filename) {
    if (!filename) {
        return { success: false, message: 'No filename entered. Export cancelled.' };
    }

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var range, sheet;

    // Determine the range to export based on sheetName
    sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
        return { success: false, message: 'Sheet not found: ' + sheetName };
    }

    // Set range based on sheetName using switch statement
    switch (sheetName) {
        case 'Item Template':
            range = sheet.getRange('B2:AO');
            break;
        case 'ERP import - Sales Order':
            range = sheet.getRange('A2:AM');
            break;
        case 'ERP Import - Sales Invoice':
            range = sheet.getRange('D2:CA');
            break;
        case 'Amazon - Schedule Template':
            range = sheet.getRange('F2:O');
            break;
        case 'ERP import - Delivery Note':
            range = sheet.getRange('A2:BD');
            break;
        case 'Stock Entry':
            range = sheet.getRange('A2:O');
            break;
        case 'Auto Invoicing':
            range = sheet.getRange('D3:BZ');
            break;
        case 'erp_price_list':
            range = sheet.getRange('E2:U');
            break;
        default:
            range = sheet.getRange('A:K'); // Default range for other sheets
    }

    // Convert the data range to CSV format
    var data = range.getValues();
    var csvContent = "";

    data.forEach(function(infoArray, index) {
        var dataString = infoArray.map(function(cell) {
            if (cell instanceof Date) {
                // Format the date to dd-MM-yyyy
                return '"' + Utilities.formatDate(cell, Session.getScriptTimeZone(), 'dd-MM-yyyy') + '"';
            }
            // Convert cell to string and escape quotes
            return '"' + (cell === null ? '' : cell.toString().replace(/"/g, '""')) + '"';
        }).join(",");
        csvContent += dataString + (index < data.length - 1 ? "\n" : "");
    });

    // Create a Blob with the CSV data
    var blob = Utilities.newBlob(csvContent, 'text/csv;charset=utf-8', filename + '.csv');

    // Create a URL for the Blob and return it
    var url = 'data:text/csv;base64,' + Utilities.base64Encode(blob.getBytes());

    return { success: true, url: url };
}
// -----------------------------------------------------Exporting End --------------------------------------------------------

// Clear the data in the "Order Formatter" sheet
function clearDataInOrderReview() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("Order Formatter");
    if (sheet) {
      // Define the range to clear
      var range = sheet.getRange("A3:E");
      range.clearContent();
    }
}

// Clear the data in the "ERP Import - Sales Invoice" sheet, excluding column AX:BA
function clearDataInERPImport_SalesInvoice() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("ERP Import - Sales Invoice");
  if (sheet) {
    // Define the range to clear, excluding column AX & AY
    var rangeToClearBeforeAX = sheet.getRange("A3:CB");
    rangeToClearBeforeAX.clearContent();
  }
}

// Moves the cursor to a specific cell
function moveToCell() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Order Formatter');
    var cellAddress = sheet.getRange('Z12').getValue();
    var cell = sheet.getRange(cellAddress);
    SpreadsheetApp.getActiveSpreadsheet().setActiveSelection(cell);
}

// Delete row and shift up 
function deleteRangeAndShiftUp() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Order Formatter');
  
  // Get the address from cell Z12
  var targetCellAddress = sheet.getRange('Z12').getValue();
  
  // Extract the row number from the address
  var targetCell = sheet.getRange(targetCellAddress);
  var targetRow = targetCell.getRow();
  
  // Define the range to delete (A to B in the target row)
  var rangeToDelete = sheet.getRange('A' + targetRow + ':B' + targetRow);
  
  // Clear content in the range
  rangeToDelete.clearContent();
  
  // Get the last row of the sheet
  var lastRow = sheet.getLastRow();
  
  // Shift rows up if needed
  if (targetRow < lastRow) {
    // Get the range of rows from A to B starting from the row below the deleted row
    var rangeToShift = sheet.getRange('A' + (targetRow + 1) + ':B' + lastRow);
    
    // Move the range up by one row
    rangeToShift.moveTo(sheet.getRange('A' + targetRow + ':B' + (lastRow - 1)));
    
    // Clear the last row in columns A and B which is now empty
    sheet.getRange('A' + lastRow + ':B' + lastRow).clearContent();
  }
}

function applyDynamicFormulaBasedOnColumnI() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ERP Import - Sales Invoice');
  if (!sheet) {
    Logger.log("ERP Import - Sales Invoice not found!");
    return;
  }

  var sheetData = sheet.getDataRange().getValues();
  var sheetDataHeader = sheetData[1];
  var fixedColIndex = sheetDataHeader.indexOf("Account Head (Sales Taxes and Charges)");
  var taxCategoryColIndexManual = sheetDataHeader.indexOf("Tax Category");
  var lastRow = sheet.getLastRow();
  var rangeI = sheet.getRange(3, 12, lastRow - 2, 1).getValues();

  for (var i = 0; i < rangeI.length; i++) {
    var rowIndex = i + 3;
    var valueI = rangeI[i][0];

    if (valueI !== "") {
      var formula = `=ARRAYFORMULA(IF(R${rowIndex}C${taxCategoryColIndexManual + 1}="In-State",fixed_data!B3:B4,IF(R${rowIndex}C${taxCategoryColIndexManual + 1}="Out-State",fixed_data!B7, "")))`;
      sheet.getRange(rowIndex, fixedColIndex + 1).setFormula(formula);
    } else {
      sheet.getRange(rowIndex, fixedColIndex + 1).clearContent();
    }
  }
}

function applyDynamicFormulaInAutoInvoicing() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Auto Invoicing');
  if (!sheet) {
    Logger.log("Auto Invoicing not found!");
    return;
  }

  var sheetDataAuto = sheet.getDataRange().getValues();
  var sheetDataHeaderAuto = sheetDataAuto[2];
  var fixedColIndexAuto = sheetDataHeaderAuto.indexOf("Account Head (Sales Taxes and Charges)");
  var taxCategoryColIndex = sheetDataHeaderAuto.indexOf("Tax Category");
  var lastRow = sheet.getLastRow();
  var rangeI = sheet.getRange(4, 12, lastRow - 2, 1).getValues();

  if (fixedColIndexAuto === 0) {
    Logger.log("Column 'Account Head (Sales Taxes and Charges)' not found.");
    return;
  }

  for (var i = 0; i < rangeI.length; i++) {
    var rowIndex = i + 4;
    var valueI = rangeI[i][0];

    if (valueI !== "") {
      var formula = `=ARRAYFORMULA(IF(R${rowIndex}C${taxCategoryColIndex + 1}="In-State",fixed_data!B3:B4,IF(R${rowIndex}C${taxCategoryColIndex + 1}="Out-State",fixed_data!B7,"")))`;
      sheet.getRange(rowIndex, fixedColIndexAuto + 1).setFormula(formula);
    } else {
      sheet.getRange(rowIndex, fixedColIndexAuto + 1).clearContent();
    }
  }
}