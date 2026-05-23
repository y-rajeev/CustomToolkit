function generatePurchaseOrderPDF() {

  const sheetName = "Encasa - Purchase Order";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  const HEADER_ROW = 2;
  const DATA_ROW = 3;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const col = {};
  headers.forEach((h, i) => { if (h) col[String(h).trim()] = i; });

  const data = sheet.getRange(DATA_ROW, 1, lastRow - DATA_ROW + 1, lastCol).getValues();

  // Group rows by purchase order number (column can be anywhere, we use the header)
  const orders = {};
  data.forEach(r => {
    const po = r[col["Purchase Order"]];
    if (!po) return;
    if (!orders[po]) orders[po] = [];
    orders[po].push(r);
  });

  const FIRST_PAGE = 33;
  const OTHER_PAGES = 40;
  const folder = DriveApp.getFolderById("1kW1Tp_ZkV24HS6gu2BRI0Hcj97oBSbYH");

  Object.keys(orders).forEach(poNumber => {

    const rows = orders[poNumber];
    const info = rows[0];

    const supplier = info[col["Supplier Name"]];
    const branch = info[col["Branch"]];
    const destination = info[col["Destination"]];
    const date = formatDate(info[col["Date"]]);
    const requiredBy = formatDate(info[col["Required By"]]);

    const isIGST = branch === "Karur";

    const items = [];
    const gstSummary = {};

    let totalTaxable = 0;
    let totalGST = 0;
    let grandTotal = 0;
    let totalQty = 0;

    rows.forEach(r => {

      if (!r[col["SKU"]]) return;

      const taxable = Number(r[col["Amount"]]);
      const hsn = r[col["HSN/SAC"]];

      const igstAmt = Number(r[col["IGST Amount"]]);
      const cgstAmt = Number(r[col["CGST Amount"]]);
      const sgstAmt = Number(r[col["SGST Amount"]]);

      const gstAmount = isIGST ? igstAmt : cgstAmt + sgstAmt;

      const total = Number(r[col["Total Amount"]]);

      totalTaxable += taxable;
      totalGST += gstAmount;
      grandTotal += total;
      totalQty += Number(r[col["Quantity (Sets)"]]);

      items.push({
        sku: r[col["SKU"]],
        line: r[col["Line"]],
        design: r[col["Design"]],
        size: r[col["Size"]],
        pcs: r[col["Pcs/Pack"]],
        qty: r[col["Quantity (Sets)"]],
        rate: fmt(r[col["Rate"]]),
        amount: fmt(r[col["Amount"]])
      });

      if (!gstSummary[hsn]) {

        gstSummary[hsn] = {
          hsn: hsn,
          taxable: 0,
          igstRate: r[col["IGST Rate"]],
          igstAmt: 0,
          cgstRate: r[col["CGST Rate"]],
          cgstAmt: 0,
          sgstRate: r[col["SGST Rate"]],
          sgstAmt: 0
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
    template.destination = destination;
    template.date = date;
    template.requiredBy = requiredBy;

    template.items = items;
    template.gstRows = gstRows;

    const pageGroups = [];
    pageGroups.push(items.slice(0, FIRST_PAGE));
    for (let i = FIRST_PAGE; i < items.length; i += OTHER_PAGES) {
      pageGroups.push(items.slice(i, i + OTHER_PAGES));
    }
    template.pageGroups = pageGroups;

    const pageOffsets = [];
    let offset = 0;
    for (let g = 0; g < pageGroups.length; g++) {
      pageOffsets.push(offset);
      offset += pageGroups[g].length;
    }
    template.pageOffsets = pageOffsets;

    template.totalTaxable = fmt(totalTaxable);
    template.totalGST = fmt(totalGST);
    template.grandTotal = fmt(grandTotal);

    const advance = grandTotal * 0.5;
    const balance = grandTotal - advance;
    template.advanceAmount = fmt(advance);
    template.balanceAmount = fmt(balance);

    template.isIGST = isIGST;
    template.totalQty = totalQty;
    template.totalItems = items.length;

    const html = template.evaluate().getContent();
    const blob = Utilities.newBlob(html,"text/html").getAs("application/pdf");

    const fileName = `PO_${poNumber}.pdf`;
    const existing = folder.getFilesByName(fileName);
    while(existing.hasNext()) existing.next().setTrashed(true);
    folder.createFile(blob.setName(fileName));

  });

}


function fmt(v){
  if(v === "" || v === null) return "";
  return Number(v).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2});
}


function formatDate(v){

  if(!v) return "";

  var d = (v instanceof Date) ? v : new Date(v);

  var str = String(v);
  var parts = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(parts) d = new Date(Number(parts[1]), Number(parts[2]) - 1, Number(parts[3]));

  var months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  var day = d.getDate();
  return (day < 10 ? "0" + day : day) + "-" + months[d.getMonth()] + "-" + d.getFullYear();

}