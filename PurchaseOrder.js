function generatePurchaseOrderPDF() {

  const sheetName = "Encasa - Purchase Order";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  const startRow = 3;
  const lastRow = sheet.getLastRow();

  const header = sheet.getRange(startRow,4,1,5).getValues()[0];

  const supplier = header[0];
  const branch = header[1];
  const poNumber = header[2];
  const date = formatDate(header[3]);
  const requiredBy = formatDate(header[4]);

  const isIGST = branch === "Karur";

  const rows = sheet.getRange(startRow,10,lastRow-2,19).getValues();

  const items = [];
  const gstSummary = {};

  let totalTaxable = 0;
  let totalGST = 0;
  let grandTotal = 0;
  let totalQty = 0;

  rows.forEach(r => {

    if (!r[0]) return;

    const taxable = Number(r[7]);
    const hsn = r[8];

    const igstAmt = Number(r[13]);
    const cgstAmt = Number(r[15]);
    const sgstAmt = Number(r[17]);

    const gstAmount = isIGST ? igstAmt : cgstAmt + sgstAmt;

    const total = Number(r[18]);

    totalTaxable += taxable;
    totalGST += gstAmount;
    grandTotal += total;
    totalQty += Number(r[5]);

    items.push({
      sku:r[0],
      line:r[1],
      design:r[2],
      size:r[3],
      pcs:r[4],
      qty:r[5],
      rate:fmt(r[6]),
      amount:fmt(r[7])
    });

    if(!gstSummary[hsn]){

      gstSummary[hsn] = {
        hsn:hsn,
        taxable:0,
        igstRate:r[12],
        igstAmt:0,
        cgstRate:r[14],
        cgstAmt:0,
        sgstRate:r[16],
        sgstAmt:0
      };

    }

    gstSummary[hsn].taxable += taxable;
    gstSummary[hsn].igstAmt += igstAmt;
    gstSummary[hsn].cgstAmt += cgstAmt;
    gstSummary[hsn].sgstAmt += sgstAmt;

  });

  const gstRows = Object.values(gstSummary).map(x=>({

    hsn:x.hsn,
    taxable:fmt(x.taxable),
    igstRate:x.igstRate,
    igstAmt:fmt(x.igstAmt),
    cgstRate:x.cgstRate,
    cgstAmt:fmt(x.cgstAmt),
    sgstRate:x.sgstRate,
    sgstAmt:fmt(x.sgstAmt)

  }));

  const template = HtmlService.createTemplateFromFile("po_template");

  template.supplier = supplier;
  template.branch = branch;
  template.poNumber = poNumber;
  template.date = date;
  template.requiredBy = requiredBy;

  template.items = items;
  template.gstRows = gstRows;

  const FIRST_PAGE = 35;
  const OTHER_PAGES = 40;
  const pageGroups = [];
  pageGroups.push(items.slice(0, FIRST_PAGE));
  for (let i = FIRST_PAGE; i < items.length; i += OTHER_PAGES) {
    pageGroups.push(items.slice(i, i + OTHER_PAGES));
  }
  template.pageGroups = pageGroups;

  template.totalTaxable = fmt(totalTaxable);
  template.totalGST = fmt(totalGST);
  template.grandTotal = fmt(grandTotal);

  template.isIGST = isIGST;
  template.totalQty = totalQty;
  template.totalItems = items.length;

  const html = template.evaluate().getContent();

  const blob = Utilities.newBlob(html,"text/html").getAs("application/pdf");

  const folder = DriveApp.getFolderById("1kW1Tp_ZkV24HS6gu2BRI0Hcj97oBSbYH");
  const fileName = `PO_${poNumber}.pdf`;

  const existing = folder.getFilesByName(fileName);
  while(existing.hasNext()) existing.next().setTrashed(true);

  folder.createFile(blob.setName(fileName));

}


function fmt(v){
  if(v === "" || v === null) return "";
  return Number(v).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2});
}


function formatDate(v){

  if(!v) return "";

  if(v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "dd-MMM-yyyy");
  }

  var str = String(v);
  var parts = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(parts){
    return Utilities.formatDate(
      new Date(Number(parts[1]), Number(parts[2]) - 1, Number(parts[3])),
      Session.getScriptTimeZone(),
      "dd-MMM-yyyy"
    );
  }

  return Utilities.formatDate(new Date(v), Session.getScriptTimeZone(), "dd-MMM-yyyy");

}