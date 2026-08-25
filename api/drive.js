import { google } from 'googleapis';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { query } = req.body;
  if (!query) return res.status(400).json({ error: 'Query required' });

  try {
    const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);

    const auth = new google.auth.JWT(
      credentials.client_email,
      null,
      credentials.private_key,
      ['https://www.googleapis.com/auth/drive.readonly']
    );

    const drive = google.drive({ version: 'v3', auth });

    const safeQuery = query.replace(/\\/g, '\\\\').replace(/'/g, "\\'");

    const response = await drive.files.list({
            q: `(name contains '${safeQuery}' or fullText contains '${safeQuery}') and trashed = false`,
      fields: 'files(id, name, mimeType, webViewLink, modifiedTime)',
      pageSize: 15,
      orderBy: 'modifiedTime desc',
    });

    const files = (response.data.files || []).map(f => ({
      name: f.name,
      type: f.mimeType.split('.').pop().replace('vnd.google-apps.', ''),
      url: f.webViewLink,
      modified: f.modifiedTime ? new Date(f.modifiedTime).toLocaleDateString('en-GB') : '',
    }));

    return res.status(200).json({ files });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
}
