import { mkdir, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import path from "node:path";
import readXlsxFile from "read-excel-file/node";

const projectRoot = fileURLToPath(new URL("../", import.meta.url));
const workbookPath = path.join(projectRoot, "DATABASE_SHINY.xlsx");
const outputDirectory = path.join(projectRoot, "src", "data");
const outputPath = path.join(outputDirectory, "products.json");

const sheets = await readXlsxFile(workbookPath);
const firstSheet = sheets[0];

if (!firstSheet || firstSheet.data.length === 0) {
  throw new Error("The Excel workbook does not contain any product rows.");
}

const rows = firstSheet.data;
const headers = rows[0].map((header) => String(header ?? "").trim());
const headerIndexes = new Map(
  headers
    .map((header, index) => [String(header ?? "").trim(), index])
    .filter(([header]) => header),
);
const requiredHeaders = ["SN", "Description", "CAPACITY", "Price", "BARCODE"];
const missingHeaders = requiredHeaders.filter(
  (header) => !headerIndexes.has(header),
);

if (missingHeaders.length > 0) {
  throw new Error(`Missing Excel column headers: ${missingHeaders.join(", ")}`);
}

const cellText = (row, header) =>
  String(row[headerIndexes.get(header)] ?? "").trim();

const products = rows
  .slice(1)
  .map((row) => ({
    sku: cellText(row, "SN"),
    description: cellText(row, "Description"),
    capacity: cellText(row, "CAPACITY"),
    price: cellText(row, "Price"),
    barcode: cellText(row, "BARCODE"),
  }))
  .filter((product) => product.sku || product.description);

if (products.length === 0) {
  throw new Error("No products were found. Check the workbook's column headers.");
}

await mkdir(outputDirectory, { recursive: true });
await writeFile(outputPath, `${JSON.stringify(products, null, 2)}\n`, "utf8");
console.log(`Imported ${products.length} products from DATABASE_SHINY.xlsx.`);
