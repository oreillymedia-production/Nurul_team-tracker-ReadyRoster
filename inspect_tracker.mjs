import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "/Users/nishak/Desktop/Final November tracker.xlsx";
const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const summary = await workbook.inspect({
  kind: "workbook,sheet,table,region",
  maxChars: 18000,
  tableMaxRows: 50,
  tableMaxCols: 40,
  tableMaxCellChars: 120,
});
console.log(summary.ndjson);

for (const sheetName of ["Tracker", "Sheet1", "November", "Events"]) {
  try {
    const render = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
    const fs = await import("node:fs/promises");
    await fs.writeFile(`outputs/speaker_outreach/${sheetName}.png`, new Uint8Array(await render.arrayBuffer()));
    console.log(`RENDERED ${sheetName}`);
    break;
  } catch {}
}
