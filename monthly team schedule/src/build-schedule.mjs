import fs from 'node:fs/promises';
import path from 'node:path';
const { FileBlob, SpreadsheetFile } = await import(process.env.ARTIFACT_TOOL_MODULE ?? '@oai/artifact-tool');

const args = new Set(process.argv.slice(2));
const argumentValue = (flag) => {
  const index = process.argv.indexOf(flag);
  return index >= 0 ? process.argv[index + 1] : undefined;
};
const asOf = argumentValue('--as-of') ?? new Date().toISOString().slice(0, 10);
const readinessOnly = args.has('--readiness-only');
const schedulePath = process.env.SCHEDULE_XLSX ?? 'data/schedule.xlsx';
const contactsPath = process.env.CONTACTS_XLSX ?? 'data/contacts.xlsx';
const templatePath = process.env.TEMPLATE_XLTX;
const outputDir = process.env.OUTPUT_DIR ?? 'artifacts';

if (!templatePath && !readinessOnly) throw new Error('Set TEMPLATE_XLTX to the monthly schedule template path.');

const canonicalName = (value) => String(value ?? '')
  .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
  .replace(/\([^)]*\)/g, '')
  .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
const extractEmails = (value) => [...new Set(String(value ?? '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) ?? [])];
const isoFromExcel = (serial) => new Date(Date.UTC(1899, 11, 30) + serial * 86400000).toISOString().slice(0, 10);
const splitCourseAndInstructor = (text) => {
  for (const delimiter of [' — ', ' - ']) {
    const index = text.lastIndexOf(delimiter);
    if (index > 0) return [text.slice(0, index), text.slice(index + delimiter.length)];
  }
  return [text, null];
};

const sourceWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(schedulePath));
const sourceRows = sourceWorkbook.worksheets.getItem('2026 December Schedule').getRange('A5:G125').values;
const contactsWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(contactsPath));
const contactRows = contactsWorkbook.worksheets.getItemAt(0).getRange('A2:B1233').values;
const emailByInstructor = new Map(contactRows
  .filter(([name, email]) => name && extractEmails(email).length)
  .map(([name, email]) => [canonicalName(name), extractEmails(email)]));

const events = sourceRows.filter((row) => row[4] === 'NI').map((row) => {
  const [title, instructor] = splitCourseAndInstructor(row[2]);
  const names = (instructor ?? '').split(',').map((name) => name.trim()).filter(Boolean);
  const missingContacts = [...new Set(names.filter((name) => !emailByInstructor.has(canonicalName(name))))];
  const contacts = [...new Set(names.flatMap((name) => emailByInstructor.get(canonicalName(name)) ?? []))];
  return {
    date: row[1], dateIso: isoFromExcel(row[1]), title, instructor, time: row[3], notes: row[5], link: row[6],
    contacts, missingContacts,
  };
});

await Promise.all(events.filter((event) => event.link).map(async (event) => {
  try {
    const response = await fetch(event.link, { redirect: 'follow' });
    const page = await response.text();
    event.liveEventStatus = response.ok && page.includes('Live event sections') ? 'verified' : 'not_verified';
  } catch (error) {
    event.liveEventStatus = 'check_error';
    event.liveEventError = error.message;
  }
}));
for (const event of events.filter((event) => !event.link)) event.liveEventStatus = 'not_published';

const dueDate = new Date(`${asOf}T00:00:00Z`);
dueDate.setUTCDate(dueDate.getUTCDate() + 3);
const dueIso = dueDate.toISOString().slice(0, 10);
for (const event of events) {
  event.editorAssigned = false; // This template leaves editor assignment as a manual workflow step.
  event.readinessGaps = [
    ...(event.missingContacts.length ? [`missing contact: ${event.missingContacts.join(', ')}`] : []),
    ...(!event.link ? ['live-event link missing'] : []),
    ...(!event.editorAssigned ? ['editor not assigned'] : []),
  ];
  event.templateReady = event.readinessGaps.length === 0;
  event.checkDue = event.dateIso === dueIso;
}
const readinessReport = {
  asOf,
  dueDate: dueIso,
  totalEvents: events.length,
  dueEvents: events.filter((event) => event.checkDue),
  allEvents: events,
};
await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(path.join(outputDir, `readiness-${asOf}.json`), JSON.stringify(readinessReport, null, 2));

if (!readinessOnly) {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(templatePath));
  const sheet = workbook.worksheets.getItem('NI schedule');
  sheet.getRange('B5').values = [[46357]];
  sheet.getRange(`A14:E${13 + events.length}`).values = events.map((event) => [
    event.date, event.time, event.title, event.instructor, event.contacts.join('; ') || null,
  ]);
  sheet.getRange(`H14:J${13 + events.length}`).values = events.map((event) => [event.link, null, event.notes]);
  events.forEach((event, index) => {
    if (!event.link) sheet.getRange(`A${14 + index}:L${14 + index}`).format.fill = '#FCE4D6';
  });
  workbook.recalculate();
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(path.join(outputDir, 'Nurul_December_2026_Team_Schedule.xlsx'));
}

console.log(JSON.stringify({ outputDir, asOf, dueDate: dueIso, events: events.length, due: readinessReport.dueEvents.length }, null, 2));
