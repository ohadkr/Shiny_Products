import { strToU8, zipSync } from "fflate";

export type CartExportRow = {
  sku: string;
  description: string;
  barcode: string;
  capacity: string;
  quantity: number;
  unitPrice: number | null;
  lineTotal: number | null;
};

function escapeXml(value: string) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function columnName(index: number) {
  let name = "";
  let number = index + 1;

  while (number > 0) {
    const remainder = (number - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    number = Math.floor((number - 1) / 26);
  }

  return name;
}

function textCell(reference: string, value: string) {
  return `<c r="${reference}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
}

function numberCell(reference: string, value: number) {
  return `<c r="${reference}"><v>${value}</v></c>`;
}

export function createCartWorkbook(rows: CartExportRow[], total: number) {
  const headers = [
    "מק״ט",
    "תיאור מוצר",
    "ברקוד",
    "אריזה",
    "כמות",
    "מחיר ליחידה (₪)",
    "סה״כ (₪)",
  ];
  const xmlRows = [
    `<row r="1">${headers.map((value, index) => textCell(`${columnName(index)}1`, value)).join("")}</row>`,
    ...rows.map((row, rowIndex) => {
      const excelRow = rowIndex + 2;
      const cells = [
        textCell(`A${excelRow}`, row.sku),
        textCell(`B${excelRow}`, row.description),
        textCell(`C${excelRow}`, row.barcode),
        textCell(`D${excelRow}`, row.capacity),
        numberCell(`E${excelRow}`, row.quantity),
        row.unitPrice === null
          ? textCell(`F${excelRow}`, "לא זמין")
          : numberCell(`F${excelRow}`, row.unitPrice),
        row.lineTotal === null
          ? textCell(`G${excelRow}`, "לא זמין")
          : numberCell(`G${excelRow}`, row.lineTotal),
      ];

      return `<row r="${excelRow}">${cells.join("")}</row>`;
    }),
    `<row r="${rows.length + 2}">${textCell(`F${rows.length + 2}`, "סה״כ לפריטים עם מחיר (₪)")}${numberCell(`G${rows.length + 2}`, total)}</row>`,
  ];

  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        "</Types>",
    ),
    "_rels/.rels": strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        "</Relationships>",
    ),
    "xl/workbook.xml": strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="הזמנה" sheetId="1" r:id="rId1"/></sheets></workbook>',
    ),
    "xl/_rels/workbook.xml.rels": strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        "</Relationships>",
    ),
    "xl/worksheets/sheet1.xml": strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        `<sheetData>${xmlRows.join("")}</sheetData></worksheet>`,
    ),
  };

  return zipSync(files, { level: 6 });
}

function downloadFile(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function getFileDate() {
  return new Date().toISOString().slice(0, 10);
}

export function downloadCartExcel(rows: CartExportRow[], total: number) {
  const data = createCartWorkbook(rows, total);
  downloadFile(
    new Blob([data], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    `shiny-order-${getFileDate()}.xlsx`,
  );
}

export function printCartAsPdf() {
  window.print();
}
