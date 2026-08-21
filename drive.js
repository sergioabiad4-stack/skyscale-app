import { google } from 'googleapis';

let driveClient = null;

function getClient() {
  if (driveClient) return driveClient;
  const keyRaw = process.env.GOOGLE_SERVICE_ACCOUNT;
  if (!keyRaw) return null;
  try {
    const key = JSON.parse(keyRaw);
    const auth = new google.auth.GoogleAuth({
      credentials: key,
      scopes: ['https://www.googleapis.com/auth/drive.readonly'],
    });
    driveClient = google.drive({ version: 'v3', auth });
    return driveClient;
  } catch { return null; }
}

// Search Drive for files matching a client name, return top 5
export async function searchClientFiles(clientName) {
  const drive = getClient();
  if (!drive) return [];
  try {
    const q = `name contains '${clientName.replace(/'/g, "\\'")}' and trashed = false`;
    const res = await drive.files.list({
      q,
      pageSize: 8,
      fields: 'files(id, name, mimeType)',
      orderBy: 'modifiedTime desc',
    });
    return res.data.files || [];
  } catch { return []; }
}

// Read a Google Sheet and return rows as array-of-arrays
export async function readSheet(fileId) {
  const drive = getClient();
  if (!drive) return [];
  try {
    const sheets = google.sheets({ version: 'v4', auth: drive._options.auth });
    const meta = await sheets.spreadsheets.get({ spreadsheetId: fileId });
    const sheetName = meta.data.sheets?.[0]?.properties?.title || 'Sheet1';
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: fileId,
      range: sheetName,
    });
    return res.data.values || [];
  } catch { return []; }
}

// Export a Google Doc/Sheet as plain text for context
export async function exportAsText(fileId, mimeType) {
  const drive = getClient();
  if (!drive) return '';
  try {
    let exportMime = 'text/plain';
    if (mimeType === 'application/vnd.google-apps.spreadsheet') exportMime = 'text/csv';
    const res = await drive.files.export({ fileId, mimeType: exportMime }, { responseType: 'text' });
    return String(res.data).substring(0, 6000); // cap context size
  } catch { return ''; }
}

// Pull Drive context for a client: search files, read top ones, return combined text
export async function getClientContext(clientName) {
  const files = await searchClientFiles(clientName);
  if (!files.length) return '';

  const parts = [`[DRIVE CONTEXT for ${clientName}]`, `Files found: ${files.map(f => f.name).join(', ')}`];

  for (const f of files.slice(0, 3)) {
    const isExportable = f.mimeType?.includes('google-apps');
    if (!isExportable) continue;
    const text = await exportAsText(f.id, f.mimeType);
    if (text) parts.push(`\n--- ${f.name} ---\n${text}`);
  }

  return parts.join('\n');
}
