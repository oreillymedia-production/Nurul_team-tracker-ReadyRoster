import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const sourcePath = "/Users/nishak/Desktop/Final November tracker.xlsx";
const outputDir = "outputs/speaker_outreach";
const outputPath = `${outputDir}/Speaker outreach tracker - flow ready.xlsx`;

const source = await FileBlob.load(sourcePath);
const workbook = await SpreadsheetFile.importXlsx(source);
const tracker = workbook.worksheets.getItem("NI event tracker");

// The source table starts at row 1, which makes the title the connector header.
// Replace only that table with a connector-safe table whose headers are on row 4.
for (const table of tracker.tables.items) table.delete();

const headers = [
  "Event date", "Time", "Course title", "Instructor", "Contact details",
  "Slides status", "Slides deadline", "Last date hosted", "ON24 presenter link",
  "Editor name", "Notes", "Event ID", "Outreach status", "To email", "CC email",
  "Course type", "Registration page", "Draft created on", "Test expectation",
  "Follow-up status", "Follow-up draft created on"
];
tracker.getRange("A4:U4").values = [headers];

const sourceRows = tracker.getRange("A5:K23").values;
const augmentedRows = sourceRows.map((row, index) => {
  const contact = String(row[4] ?? "");
  const emails = [...contact.matchAll(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi)].map(m => m[0]);
  const isBlocked = /Needs instructor match/i.test(String(row[3] ?? "")) ||
    /Needs contact match/i.test(contact) ||
    /Hold\s*\/\s*confirm before outreach/i.test(String(row[10] ?? ""));
  const courseType = String(row[7] ?? "").trim().toLowerCase() === "new course" ? "New course" : "Returning course";
  const testReady = index < 3 && !isBlocked && emails.length > 0;
  const deadline = Number(row[6]);
  const todaySerial = 46323; // 2026-10-28, the current local date for this tracker review.
  const materialsNotReceived = !/received|complete|submitted/i.test(String(row[5] ?? ""));
  const followUpReady = !isBlocked && emails.length > 0 && materialsNotReceived && deadline < todaySerial;
  return [
    ...row,
    `EVT-${String(index + 1).padStart(3, "0")}`,
    testReady ? "Ready for draft" : "Not ready",
    emails[0] ?? "",
    emails.slice(1).join("; "),
    courseType,
    "",
    "",
    testReady ? "Test: eligible; expect Outlook draft" : (isBlocked ? "Skip: incomplete or hold" : "Not queued"),
    followUpReady ? "Follow-up ready" : (isBlocked ? "Skip: incomplete or hold" : "Not due"),
    ""
  ];
});
tracker.getRange("A5:U23").values = augmentedRows;

const eventTable = tracker.tables.add("A4:U23", true, "EventTracker");
eventTable.style = "TableStyleMedium2";
eventTable.showBandedColumns = false;

tracker.getRange("A4:U4").format = {
  fill: "#1F4E78",
  font: { color: "#FFFFFF", bold: true },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
};
tracker.getRange("A5:U23").format.verticalAlignment = "center";
tracker.getRange("A5:A23").format.numberFormat = "mm/dd/yy";
tracker.getRange("G5:G23").format.numberFormat = "mm/dd/yy";
tracker.getRange("R5:R23").format.numberFormat = "mm/dd/yy h:mm AM/PM";
tracker.getRange("U5:U23").format.numberFormat = "mm/dd/yy h:mm AM/PM";
tracker.getRange("M5:M23").dataValidation = { rule: { type: "list", values: ["Not ready", "Ready for draft", "Draft created"] } };
tracker.getRange("T5:T23").dataValidation = { rule: { type: "list", values: ["Not due", "Follow-up ready", "Follow-up draft created"] } };
tracker.getRange("M5:M23").conditionalFormats.add("containsText", {
  text: "Ready for draft", format: { fill: "#FFF2CC", font: { color: "#7F6000", bold: true } }
});
tracker.getRange("M5:M23").conditionalFormats.add("containsText", {
  text: "Draft created", format: { fill: "#E2F0D9", font: { color: "#375623", bold: true } }
});
tracker.getRange("S5:S23").conditionalFormats.add("containsText", {
  text: "Skip:", format: { fill: "#FCE4D6", font: { color: "#C00000", bold: true } }
});
tracker.getRange("T5:T23").conditionalFormats.add("containsText", {
  text: "Follow-up ready", format: { fill: "#FCE4D6", font: { color: "#C00000", bold: true } }
});
tracker.getRange("T5:T23").conditionalFormats.add("containsText", {
  text: "Follow-up draft created", format: { fill: "#E2F0D9", font: { color: "#375623", bold: true } }
});
tracker.freezePanes.freezeRows(4);
tracker.getRange("A:S").format.font = { name: "Arial", size: 10 };

const widths = {
  A: 12, B: 18, C: 34, D: 25, E: 34, F: 15, G: 14, H: 16, I: 38,
  J: 22, K: 36, L: 12, M: 18, N: 28, O: 28, P: 17, Q: 38, R: 20, S: 31, T: 22, U: 24
};
for (const [column, width] of Object.entries(widths)) tracker.getRange(`${column}:${column}`).format.columnWidth = width;
tracker.getRange("A1:K2").format.wrapText = true;
tracker.getRange("A1:K2").format.font = { name: "Arial", size: 10, italic: true, color: "#595959" };
tracker.getRange("A1:K1").format.font = { name: "Arial", size: 14, bold: true, color: "#1F1F1F" };

const setup = workbook.worksheets.add("Power Automate setup");
setup.showGridLines = false;
setup.getRange("A1:F1").merge();
setup.getRange("A1").values = [["Speaker outreach: manual draft flow setup"]];
setup.getRange("A1").format = { font: { name: "Arial", size: 16, bold: true, color: "#1F1F1F" }, verticalAlignment: "center" };
setup.getRange("A3:B3").values = [["Before you run", "What to do"]];
setup.getRange("A4:B7").values = [
  ["Store the workbook", "Upload this workbook to OneDrive for Business or SharePoint. Power Automate cannot use the Desktop copy."],
  ["Complete each row", "Enter Registration page, To email, and CC email where needed. Confirm Instructor and ON24 presenter link."],
  ["Queue rows", "Set Outreach status to Ready for draft only after every field has been reviewed."],
  ["Review in Outlook", "The flow creates drafts only. It contains no Send email action. Review and send each draft yourself."],
];
setup.getRange("A9:D9").values = [["Flow step", "Connector / action", "Configuration", "Safety check"]];
setup.getRange("A10:D17").values = [
  ["1", "Instant cloud flow — Manually trigger a flow", "Name: Create speaker outreach drafts", "Manual trigger only"],
  ["2", "Excel Online (Business) — List rows present in a table", "Choose the uploaded file and table EventTracker. Turn on pagination if there will be more than 256 rows.", "Read the table, not a worksheet range"],
  ["3", "Filter array", "From: value. Advanced mode: @equals(item()?['Outreach status'],'Ready for draft')", "Prevents drafts for unqueued rows"],
  ["4", "Apply to each + Condition", "Advanced mode: @and(not(empty(item()?['To email'])),not(empty(item()?['Instructor'])),not(startsWith(item()?['Instructor'],'Needs')),not(contains(toLower(coalesce(item()?['Notes'],'')),'hold / confirm before outreach')))", "No draft and no status change when incomplete or held"],
  ["5", "Microsoft 365 Outlook — Draft an email message", "To = To email. CC = CC email. Add a Course type condition and use the matching subject/body below.", "Do not add Send an email (V2)"],
  ["6", "Excel Online (Business) — Update a row", "Key column: Event ID. Set Outreach status = Draft created and Draft created on = utcNow().", "Runs only after the draft action succeeds"],
  ["7", "Test", "Run with the first three rows already marked Ready for draft. Confirm three drafts, then confirm three status updates.", "Never click Send in Outlook"],
  ["8", "Reuse", "For future events, fill required fields, set only reviewed rows to Ready for draft, then run the same flow manually.", "Do not reuse Draft created rows"],
];
setup.getRange("A20:B20").values = [["New course subject", "Slides and materials request — @{items('Apply_to_each')?['Course title']}"]];
setup.getRange("A21:B21").values = [["New course body", "Hi @{items('Apply_to_each')?['Instructor']},\n\nI’m looking forward to your upcoming OLT, @{items('Apply_to_each')?['Course title']}, scheduled for @{formatDateTime(items('Apply_to_each')?['Event date'],'MMMM d, yyyy')} at @{items('Apply_to_each')?['Time']}.\n\nBecause this is a new course, please send your final slides and any supporting materials by @{formatDateTime(items('Apply_to_each')?['Slides deadline'],'MMMM d, yyyy')} so we can upload and review them before the session.\n\n@if(empty(items('Apply_to_each')?['Registration page']),'',concat('Registration page: ',items('Apply_to_each')?['Registration page'],'\\n'))Presenter link: @{items('Apply_to_each')?['ON24 presenter link']}\n\nPlease let me know if you need support preparing for the live event.\n\nBest,\n[Your name]"]];
setup.getRange("A23:B23").values = [["Returning course subject", "Slides and materials request — @{items('Apply_to_each')?['Course title']}"]];
setup.getRange("A24:B24").values = [["Returning course body", "Hi @{items('Apply_to_each')?['Instructor']},\n\nThank you for returning to lead @{items('Apply_to_each')?['Course title']} on @{formatDateTime(items('Apply_to_each')?['Event date'],'MMMM d, yyyy')} at @{items('Apply_to_each')?['Time']}.\n\nPlease send your final slides and any updated supporting materials by @{formatDateTime(items('Apply_to_each')?['Slides deadline'],'MMMM d, yyyy')} so we can upload and review them before the session.\n\n@if(empty(items('Apply_to_each')?['Registration page']),'',concat('Registration page: ',items('Apply_to_each')?['Registration page'],'\\n'))Presenter link: @{items('Apply_to_each')?['ON24 presenter link']}\n\nPlease let me know if anything has changed since the last time you presented this course.\n\nBest,\n[Your name]"]];
setup.getRange("A27:D27").values = [["Follow-up drafts", "How they were selected", "Result", "Reuse"]];
setup.getRange("A28:D28").values = [["Draft only; never send automatically", "Slides status is not received/complete, Slides deadline is before October 28, 2026, a To email and valid instructor exist, and no hold note is present.", "Two rows qualify: EVT-001 and EVT-002. Rows with an October 28 deadline are not yet overdue.", "For a second manual flow, filter Follow-up status = Follow-up ready; after Draft an email message succeeds, update it to Follow-up draft created and set the timestamp."]];
setup.getRange("A27:D27").format = { fill: "#1F4E78", font: { color: "#FFFFFF", bold: true }, horizontalAlignment: "center" };
setup.getRange("A27:D28").format.borders = { preset: "all", style: "thin", color: "#D9D9D9" };
setup.getRange("A28:D28").format.wrapText = true;
setup.getRange("A28:D28").format.verticalAlignment = "top";
setup.getRange("A28:D28").format.rowHeight = 68;
setup.getRange("A3:B3").format = { fill: "#1F4E78", font: { color: "#FFFFFF", bold: true }, horizontalAlignment: "center" };
setup.getRange("A9:D9").format = { fill: "#1F4E78", font: { color: "#FFFFFF", bold: true }, horizontalAlignment: "center" };
setup.getRange("A20:B20").format = { fill: "#D9EAF7", font: { bold: true } };
setup.getRange("A23:B23").format = { fill: "#D9EAF7", font: { bold: true } };
setup.getRange("A3:B7").format.borders = { preset: "all", style: "thin", color: "#D9D9D9" };
setup.getRange("A9:D17").format.borders = { preset: "all", style: "thin", color: "#D9D9D9" };
setup.getRange("A20:B24").format.wrapText = true;
setup.getRange("A21:B21").format.verticalAlignment = "top";
setup.getRange("A24:B24").format.verticalAlignment = "top";
setup.getRange("A:A").format.columnWidth = 25;
setup.getRange("B:B").format.columnWidth = 110;
setup.getRange("C:C").format.columnWidth = 50;
setup.getRange("D:D").format.columnWidth = 36;
setup.getRange("A21:B21").format.rowHeight = 205;
setup.getRange("A24:B24").format.rowHeight = 205;
setup.freezePanes.freezeRows(3);

const followUps = workbook.worksheets.add("Follow-up drafts");
followUps.showGridLines = false;
followUps.getRange("A1:H1").merge();
followUps.getRange("A1").values = [["Overdue materials follow-up drafts — review before sending"]];
followUps.getRange("A1").format = { font: { name: "Arial", size: 16, bold: true, color: "#1F1F1F" }, verticalAlignment: "center" };
followUps.getRange("A3:H3").merge();
followUps.getRange("A3").values = [["Selection rule: slides are not received, the deadline is before October 28, 2026, the row is not on hold, and the speaker/contact is complete. These are draft bodies only; no messages were sent."]];
followUps.getRange("A3").format = { fill: "#FFF2CC", font: { name: "Arial", size: 10, color: "#7F6000", italic: true }, wrapText: true, verticalAlignment: "center" };
followUps.getRange("A5:H5").values = [["Event ID", "To", "CC", "Speaker", "Course title", "Slides deadline", "Subject", "Draft body"]];
const formatDate = serial => {
  const date = new Date(Date.UTC(1899, 11, 30 + Number(serial)));
  return new Intl.DateTimeFormat("en-US", { month: "long", day: "numeric", year: "numeric", timeZone: "UTC" }).format(date);
};
const followUpRows = augmentedRows
  .filter(row => row[19] === "Follow-up ready")
  .map(row => {
    const eventDate = formatDate(row[0]);
    const deadline = formatDate(row[6]);
    const registration = row[16] ? `\nRegistration page: ${row[16]}` : "";
    const body = `Hi ${row[3]},\n\nA quick follow-up on the final slides and materials for ${row[2]}. The materials deadline of ${deadline} has passed, and we need them to complete the upload and review before the ${eventDate} live session.${registration}\n\nPresenter link: ${row[8]}\n\nPlease reply with the final materials as soon as possible, or let me know if you need support.\n\nBest,\n[Your name]`;
    return [row[11], row[13], row[14], row[3], row[2], row[6], `Follow-up: slides and materials needed — ${row[2]}`, body];
  });
followUps.getRange(`A6:H${5 + followUpRows.length}`).values = followUpRows;
const followUpTable = followUps.tables.add(`A5:H${5 + followUpRows.length}`, true, "FollowUpDrafts");
followUpTable.style = "TableStyleMedium2";
followUps.getRange("A5:H5").format = { fill: "#1F4E78", font: { color: "#FFFFFF", bold: true }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };
followUps.getRange(`F6:F${5 + followUpRows.length}`).format.numberFormat = "mm/dd/yy";
followUps.getRange(`G6:G${5 + followUpRows.length}`).format.wrapText = true;
followUps.getRange(`H6:H${5 + followUpRows.length}`).format.wrapText = true;
followUps.getRange(`A6:H${5 + followUpRows.length}`).format.verticalAlignment = "top";
followUps.getRange(`A6:H${5 + followUpRows.length}`).format.rowHeight = 156;
for (const [column, width] of Object.entries({ A: 12, B: 30, C: 25, D: 24, E: 34, F: 15, G: 48, H: 78 })) {
  followUps.getRange(`${column}:${column}`).format.columnWidth = width;
}
followUps.freezePanes.freezeRows(5);

workbook.recalculate();
const check = await workbook.inspect({
  kind: "table,region",
  sheetId: "NI event tracker",
  range: "A4:U10",
  include: "values,formulas",
  tableMaxRows: 10,
  tableMaxCols: 20,
  maxChars: 10000,
});
console.log(check.ndjson);
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 50 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);
for (const sheetName of ["NI event tracker", "Power Automate setup", "Follow-up drafts"]) {
  const image = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/${sheetName.replaceAll(" ", "_")}.png`, new Uint8Array(await image.arrayBuffer()));
}
await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`SAVED ${outputPath}`);
