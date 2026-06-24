import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/brodrigues1/Downloads/escala-culto-2026-07 (9).xlsx";
const outputDir = "C:/Users/brodrigues1/Documents/Codex/2026-06-23/n/outputs/excel-repair";
const outputPath = `${outputDir}/escala-culto-2026-07-reparado.xlsx`;

await fs.mkdir(outputDir, { recursive: true });

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const summary = await workbook.inspect({
  kind: "workbook,sheet,table",
  maxChars: 3000,
  tableMaxRows: 6,
  tableMaxCols: 8,
});
console.log(summary.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

const preview = await workbook.render({
  sheetName: "Escala do Culto",
  autoCrop: "all",
  scale: 1,
  format: "png",
});
await fs.writeFile(
  `${outputDir}/escala-culto-2026-07-preview.png`,
  new Uint8Array(await preview.arrayBuffer()),
);

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
