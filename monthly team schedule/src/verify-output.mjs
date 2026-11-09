import fs from 'node:fs/promises';
const { FileBlob, SpreadsheetFile } = await import(process.env.ARTIFACT_TOOL_MODULE ?? '@oai/artifact-tool');
const outputPath = process.env.OUTPUT_XLSX ?? 'artifacts/Nurul_December_2026_Team_Schedule.xlsx';
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(outputPath));
console.log((await workbook.inspect({
  kind: 'table', sheetId: 'NI schedule', range: 'A5:L24', include: 'values,formulas',
  tableMaxRows: 25, tableMaxCols: 12, maxChars: 12000,
})).ndjson);
console.log((await workbook.inspect({
  kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',
  options: { useRegex: true, maxResults: 300 }, summary: 'final formula error scan',
})).ndjson);
const image = await workbook.render({ sheetName: 'NI schedule', range: 'A1:L28', scale: 2, format: 'png' });
await fs.writeFile('artifacts/preview.png', new Uint8Array(await image.arrayBuffer()));
