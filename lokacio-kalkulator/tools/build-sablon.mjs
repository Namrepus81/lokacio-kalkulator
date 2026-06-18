import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = path.resolve("D:/Web/Projektek/lokacio-kalkulator");
const workbook = Workbook.create();

const dataSheet = workbook.worksheets.add("Adatok");
const helpSheet = workbook.worksheets.add("Súgó");

const headers = [
  "cikkszám",
  "megnevezés",
  "lokáció",
  "min",
  "max",
  "forgás",
  "tároló",
  "súly_kategória",
  "kapacitás",
  "megjegyzés",
];

const sampleRows = [
  ["1013818", "Minta csavar", "A-A-1-1", 500, 3000, "sűrű", "kis KLT", "könnyű", 1200, "jó helyen van"],
  ["1013818", "Minta csavar", "A-A-1-2", 500, 3000, "sűrű", "kis KLT", "könnyű", 800, "második alsó szint"],
  ["1013818", "Minta csavar", "A-A-1-3", 500, 3000, "sűrű", "kis KLT", "könnyű", 1500, "felső tartalékhely"],
  ["1022455", "Minta fedél", "C-A-1-2", 1000, 6000, "sűrű", "XL láda", "nehéz", 2000, "sűrű és nehéz anyag rossz helyen"],
  ["1088891", "Minta konzol", "B-A-1-1", 200, 1500, "lassú", "nagy KLT", "közepes", 1800, "lassú anyag elöl"],
  ["1099911", "Hiányos adat", "B-A-1-4", null, 2200, "közepes", "raklap", "nehéz", 1200, "hiányzó minimum"],
];

dataSheet.getRange("A1:J1").values = [headers];
dataSheet.getRange("A2:J7").values = sampleRows;
dataSheet.getRange("A1:J1").format = {
  fill: "#0F6B55",
  font: { bold: true, color: "#FFFFFF" },
};
dataSheet.getRange("A1:J7").format.borders = { preset: "all", style: "thin", color: "#D8DFDA" };
dataSheet.getRange("A:J").format.columnWidthPx = 130;
dataSheet.getRange("B:B").format.columnWidthPx = 170;
dataSheet.getRange("C:C").format.columnWidthPx = 150;
dataSheet.getRange("J:J").format.columnWidthPx = 250;
dataSheet.freezePanes.freezeRows(1);

helpSheet.getRange("A1:D1").values = [["Mező", "Ajánlott értékek", "Kötelező?", "Megjegyzés"]];
helpSheet.getRange("A2:D9").values = [
  ["cikkszám", "szöveg vagy szám", "igen", "Vonalkódolvasóval is kereshető."],
  ["lokáció", "pl. A-A-1-1, A-A-1-2", "igen", "Az utolsó szám a szint. Max 4. szint."],
  ["min", "szám", "igen", "Minimum készlet."],
  ["max", "szám", "igen", "Maximum készlet."],
  ["forgás", "sűrű, közepes, lassú", "igen", "A sűrű anyag lehetőleg elöl legyen."],
  ["tároló", "kis KLT, nagy KLT, XL láda, raklap", "igen", "Nagy vagy XL láda ne kerüljön magasra."],
  ["súly_kategória", "könnyű, közepes, nehéz", "igen", "Nehéz anyag ne legyen 2. vagy magasabb szinten."],
  ["kapacitás", "szám, pl. 2000", "igen", "Az adott lokáció/szint befogadóképessége darabban."],
];
helpSheet.getRange("A1:D1").format = {
  fill: "#235F9F",
  font: { bold: true, color: "#FFFFFF" },
};
helpSheet.getRange("A1:D10").format.borders = { preset: "all", style: "thin", color: "#D8DFDA" };
helpSheet.getRange("A:D").format.columnWidthPx = 210;
helpSheet.getRange("D:D").format.columnWidthPx = 330;
helpSheet.freezePanes.freezeRows(1);

const inspect = await workbook.inspect({
  kind: "table",
  range: "Adatok!A1:J7",
  include: "values",
  tableMaxRows: 8,
  tableMaxCols: 10,
});
console.log(inspect.ndjson);

const preview = await workbook.render({ sheetName: "Adatok", range: "A1:J7", scale: 1, format: "png" });
await fs.writeFile(path.join(root, "sablon-preview.png"), new Uint8Array(await preview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(path.join(root, "lokacio-kalkulator-sablon.xlsx"));
