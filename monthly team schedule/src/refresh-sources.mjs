import fs from 'node:fs/promises';

const outputDir = process.env.SOURCE_DIR ?? 'data';
const token = process.env.GOOGLE_ACCESS_TOKEN;
const files = {
  schedule: process.env.SCHEDULE_SHEET_ID ?? '1FK58zn4BemhJ8-Qg6pc0ZFZE_Yy86eRgceJbkmTShRw',
  contacts: process.env.CONTACTS_SHEET_ID ?? '1AA__cjBjtvLeH5mJzAKaG4Hs3vvm3vF4dhj3sozfi3Y',
};

if (!token) throw new Error('Set GOOGLE_ACCESS_TOKEN with read-only access to both Google Sheets before refreshing source files.');
await fs.mkdir(outputDir, { recursive: true });
for (const [name, id] of Object.entries(files)) {
  const url = `https://www.googleapis.com/drive/v3/files/${id}/export?mimeType=${encodeURIComponent('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}`;
  const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!response.ok) throw new Error(`Could not export ${name}: ${response.status} ${response.statusText}`);
  await fs.writeFile(`${outputDir}/${name}.xlsx`, new Uint8Array(await response.arrayBuffer()));
  console.log(`Refreshed ${name}.xlsx`);
}
