import fs from 'node:fs/promises';
import path from 'node:path';
const { FileBlob, SpreadsheetFile } = await import(process.env.ARTIFACT_TOOL_MODULE ?? '@oai/artifact-tool');

const schedulePath = process.env.SCHEDULE_XLSX ?? 'data/schedule.xlsx';
const contactsPath = process.env.CONTACTS_XLSX ?? 'data/contacts.xlsx';
const on24LinksPath = process.env.ON24_LINKS_JSON ?? 'data/on24-presenter-links.json';
const outputDir = process.env.OUTPUT_DIR ?? 'artifacts';
const normal = (value) => String(value ?? '')
  .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .replace(/\([^)]*\)/g, '')
  .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
const emails = (value) => [...new Set(String(value ?? '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) ?? [])];
const isoFromExcel = (serial) => new Date(Date.UTC(1899, 11, 30) + Number(serial) * 86400000).toISOString().slice(0, 10);
const splitCourseAndInstructor = (value) => {
  for (const delimiter of [' — ', ' - ']) {
    const position = String(value).lastIndexOf(delimiter);
    if (position > 0) return [String(value).slice(0, position), String(value).slice(position + delimiter.length)];
  }
  return [value, null];
};

let on24Links = {};
try {
  on24Links = JSON.parse(await fs.readFile(on24LinksPath, 'utf8'));
} catch (error) {
  if (error.code !== 'ENOENT') throw error;
}

const scheduleWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(schedulePath));
const scheduleRows = scheduleWorkbook.worksheets.getItem('2026 December Schedule').getRange('A5:G125').values;
const contactsWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(contactsPath));
const contactRows = contactsWorkbook.worksheets.getItemAt(0).getRange('A2:B1233').values;
const contactsByName = new Map(contactRows
  .filter(([name, email]) => name && emails(email).length)
  .map(([name, email]) => [normal(name), emails(email)]));

const merges = scheduleRows.filter((row) => row[4] === 'NI').map((row) => {
  const [courseTitle, speakerName] = splitCourseAndInstructor(row[2]);
  const courseDate = isoFromExcel(row[1]);
  const key = `${courseDate}|${courseTitle}`;
  const speakerNames = String(speakerName ?? '').split(',').map((name) => name.trim()).filter(Boolean);
  const to = [...new Set(speakerNames.flatMap((name) => contactsByName.get(normal(name)) ?? []))];
  const on24PresenterLink = on24Links[key] ?? null;
  const missing = [
    ...(to.length ? [] : ['speaker email']),
    ...(row[6] ? [] : ['registration course link']),
    ...(on24PresenterLink ? [] : ['ON24 presenter link']),
  ];
  return {
    eventKey: key,
    speakerName,
    courseTitle,
    courseDate,
    registrationCourseLink: row[6] ?? null,
    on24PresenterLink,
    to,
    draftReady: missing.length === 0,
    missing,
  };
});

await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(path.join(outputDir, 'speaker-email-merge.json'), JSON.stringify({ merges }, null, 2));
console.log(JSON.stringify({ events: merges.length, readyForDraft: merges.filter((merge) => merge.draftReady).length, output: path.join(outputDir, 'speaker-email-merge.json') }, null, 2));
