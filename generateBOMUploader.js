/*** BOM Tools — generates ERPNext BOM import data from the "BOM File" sheet.
 *   Lives in the spreadsheet (Apps Script), NOT in the GSheet MCP server.
 *   Menu: BOM Tools → Generate BOM Import
 ***/

var SPREADSHEET_ID = '1lbESZU51c2jkW_A19lLo3g1P3SyPRQ-JWcj49fDHgp8';

var BOM_SHEET        = 'BOM File';
var ITEM_LIST_SHEET  = 'Item List';
var VALIDATION_SHEET = 'BOM Validation';
var L0_SHEET         = 'BOM Import Generated (L0)';
var L1_SHEET         = 'BOM Import Generated (L1)';

var FIRST_DATA_ROW = 3;     // rows 1-2 are headers
var NUM_BLOCKS     = 9;     // up to 9 material blocks
var TOTAL_COLS     = 7 + NUM_BLOCKS * 4;  // 43

// Toggles
var VALIDATE_AGAINST_ITEM_LIST = true;   // flag codes missing from Item List.sku
var EXCLUDE_INVALID_FROM_IMPORT = false; // false = include all rows (matches old MCP); true = skip invalid SKUs

var IMPORT_HEADER = ['Item','Quantity','Is Active','Is Default','Finished Goods based Operating Cost',
  'Item Code (Items)','Qty (Items)','Item Code (Scrap Items)','Qty (Scrap Items)',
  'Do Not Explode (Items)','Rate Of Materials Based On','Price List'];

function generateBOMUploader() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(BOM_SHEET);
  var lastRow = sheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) { _alert('No data rows in "' + BOM_SHEET + '".'); return; }

  var values = sheet.getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, TOTAL_COLS).getValues();

  // Pass 1: collect all SKUs (for level + do-not-explode inference)
  var skuSet = {};
  values.forEach(function(row) {
    var sku = String(row[0] || '').trim();
    if (sku) skuSet[sku] = true;
  });

  // Optional: item codes known to ERPNext (Item List.sku, column C)
  var itemSet = {};
  if (VALIDATE_AGAINST_ITEM_LIST) {
    var il = ss.getSheetByName(ITEM_LIST_SHEET);
    if (il && il.getLastRow() >= 2) {
      il.getRange(2, 3, il.getLastRow() - 1, 1).getValues().forEach(function(r) {
        var c = String(r[0] || '').trim();
        if (c) itemSet[c] = true;
      });
    }
  }

  // Pass 2a: group physical rows into SKU records.
  // A row with a SKU starts a new record; a blank-SKU row is a CONTINUATION
  // of the record above it (its material blocks belong to that same SKU).
  var records = [];
  var current = null;
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rowNum = FIRST_DATA_ROW + i;
    var sku = String(row[0] || '').trim();

    // Gather components present in THIS physical row (across all 9 blocks)
    var rowComps = [];
    for (var b = 0; b < NUM_BLOCKS; b++) {
      var base = 7 + b * 4;                       // col H = index 7
      var code = String(row[base] || '').trim();
      if (!code) continue;
      rowComps.push({ code: code, qty: row[base + 1], dne: row[base + 3] });  // col K = base+3
    }

    if (sku) {
      current = { sku: sku, rowNum: rowNum, rate: row[5], priceList: row[6], comps: rowComps };
      records.push(current);
    } else if (current) {
      current.comps = current.comps.concat(rowComps);  // continuation row
    }
    // else: leading blank row before any SKU — ignore
  }

  // Pass 2b: validate + build outputs from the grouped records
  var validationRows = [];
  var l0 = [], l1 = [];
  var seenSku = {};

  for (var i = 0; i < records.length; i++) {
    var rec = records[i];
    var rowNum = rec.rowNum;
    var sku = rec.sku;
    var rate = rec.rate, priceList = rec.priceList;

    var comps = rec.comps;
    comps.forEach(function(c, idx) { c.n = idx + 1; });  // sequential material number

    var level = (comps.length && skuSet[comps[0].code]) ? 1 : 0;

    // Validation
    var issues = [];
    if (seenSku[sku]) issues.push('Duplicate SKU "' + sku + '"');
    seenSku[sku] = true;
    if (comps.length === 0) issues.push('No materials defined');
    comps.forEach(function(c) {
      if (c.qty === '' || c.qty === null) {
        issues.push('Material ' + c.n + ' (' + c.code + '): quantity is required when item code is present');
      }
      if (VALIDATE_AGAINST_ITEM_LIST && !itemSet[c.code] && !skuSet[c.code]) {
        issues.push('Material ' + c.n + ' (' + c.code + '): item code not found in Item List');
      }
    });
    var valid = issues.length === 0;

    validationRows.push([rowNum, sku, level, comps.length, valid, issues.join('; ')]);

    if (!valid && EXCLUDE_INVALID_FROM_IMPORT) continue;

    // Build import rows (first component shares the parent row)
    var target = (level === 1) ? l1 : l0;
    for (var j = 0; j < comps.length; j++) {
      var c = comps[j];
      // Use the user-entered "Do Not Explode" value from the BOM File when present;
      // fall back to inference (a component that is itself a SKU defaults to 1) only when blank.
      var dne = _hasValue(c.dne) ? _toBool01(c.dne) : (skuSet[c.code] ? 1 : 0);
      if (j === 0) {
        // Parent row carries the BOM-level fields (rate / price list).
        target.push([sku, 1, 1, 1, '', c.code, c.qty, '', '', dne, rate, priceList]);
      } else {
        // Child rows only carry the component fields; parent-level fields stay blank.
        target.push(['', '', '', '', '', c.code, c.qty, '', '', dne, '', '']);
      }
    }
  }

  // Write outputs
  _writeSheet(ss, VALIDATION_SHEET, ['Row','SKU','Level','Material Count','Valid','Issues'], validationRows);
  _writeSheet(ss, L0_SHEET, IMPORT_HEADER, l0);
  _writeSheet(ss, L1_SHEET, IMPORT_HEADER, l1);

  var invalid = validationRows.filter(function(r) { return r[4] === false; }).length;
  _alert('BOM import generated.\n\n' +
         'SKUs: ' + validationRows.length + '  (invalid: ' + invalid + ')\n' +
         'L0 rows: ' + l0.length + '   L1 rows: ' + l1.length +
         (invalid ? '\n\nFix invalid rows in "BOM Validation" before uploading.' : ''));
}

function _writeSheet(ss, name, header, rows) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clearContents();
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
}

function _alert(msg) { SpreadsheetApp.getUi().alert(msg); }

// True when a cell actually holds a value (0 / "0" / false count as present).
function _hasValue(v) { return v !== '' && v !== null && v !== undefined; }

// Normalise a "Do Not Explode" cell to 1/0. Accepts 1/0, true/false, yes/no, y/n, x/✓.
function _toBool01(v) {
  if (v === true) return 1;
  if (v === false) return 0;
  if (typeof v === 'number') return v ? 1 : 0;
  var s = String(v).trim().toLowerCase();
  return (s === '1' || s === 'yes' || s === 'y' || s === 'true' || s === 'x' || s === '✓') ? 1 : 0;
}
