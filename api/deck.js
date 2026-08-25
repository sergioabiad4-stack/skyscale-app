import PptxGenJS from 'pptxgenjs';
import { google } from 'googleapis';

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { text, client } = req.body;
  if (!text) return res.status(400).json({ error: 'Text required' });

  let driveFiles = [];
  try {
    const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth = new google.auth.JWT(
      credentials.client_email, null, credentials.private_key,
      ['https://www.googleapis.com/auth/drive.readonly']
    );
    const drive = google.drive({ version: 'v3', auth });
    const safeQuery = text.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
    const response = await drive.files.list({
      q: "(name contains '" + safeQuery + "' or fullText contains '" + safeQuery + "') and trashed = false",
      fields: 'files(id, name, mimeType, webViewLink, modifiedTime)',
      pageSize: 10, orderBy: 'modifiedTime desc',
    });
    driveFiles = (response.data.files || []).map(f => ({
      name: f.name,
      type: f.mimeType.split('.').pop().replace('vnd.google-apps.', ''),
    }));
  } catch (e) { console.error('Drive error:', e.message); }

  const driveContext = driveFiles.length > 0
    ? '\nDRIVE FILES FOUND:\n' + driveFiles.map(f => '- ' + f.name + ' (' + f.type + ')').join('\n')
    : '';

  const today = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'long', year: 'numeric' });
  const aiPrompt = 'Create a professional PowerPoint deck for: ' + text + '\nClient: ' + (client || 'Skyscale Media') + '\nToday: ' + today + driveContext + '\n\nReturn ONLY valid JSON:\n{\n  "title": "Deck title here",\n  "subtitle": "Client | Campaign | Date",\n  "slides": [\n    { "title": "SLIDE TITLE", "bullets": ["Point one", "Point two", "Point three"] }\n  ]\n}\nRules: 6-9 slides max. Each slide max 5 bullets. Include Executive Summary, strategy slides, Next Steps.';

  let deckData;
  try {
    const aiRes = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': process.env.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({ model: 'claude-sonnet-4-5', max_tokens: 2048, messages: [{ role: 'user', content: aiPrompt }] }),
    });
    const aiData = await aiRes.json();
    const raw = aiData.content?.[0]?.text || '';
    const match = raw.match(/\{[\s\S]*\}/);
    if (!match) throw new Error('No JSON returned');
    deckData = JSON.parse(match[0]);
  } catch (e) { return res.status(500).json({ error: 'AI failed: ' + e.message }); }

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  const BG = '080A0F', SURFACE = '0D1017', ACCENT = 'E8371E', WHITE = 'FFFFFF', MUTED = 'C8CDD8', DIM = '444C5C';

  const ts = pptx.addSlide();
  ts.background = { color: BG };
  ts.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.5, h: 7.5, fill: { color: ACCENT } });
  ts.addShape(pptx.ShapeType.rect, { x: 0, y: 6.5, w: 13.33, h: 1, fill: { color: SURFACE } });
  ts.addText('SKYSCALE MEDIA', { x: 0.8, y: 0.9, w: 11, h: 0.5, fontSize: 11, color: ACCENT, bold: true, charSpacing: 5 });
  ts.addText(deckData.title, { x: 0.8, y: 1.6, w: 11, h: 2, fontSize: 38, color: WHITE, bold: true });
  ts.addText(deckData.subtitle, { x: 0.8, y: 3.8, w: 11, h: 0.5, fontSize: 15, color: MUTED });
  ts.addText('CONFIDENTIAL  |  skyscalemedia.com', { x: 0.8, y: 6.7, w: 12, h: 0.35, fontSize: 9, color: DIM, charSpacing: 2 });

  for (const slide of deckData.slides || []) {
    const s = pptx.addSlide();
    s.background = { color: BG };
    s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.12, h: 7.5, fill: { color: ACCENT } });
    s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 1.1, fill: { color: SURFACE } });
    s.addText((slide.title || '').toUpperCase(), { x: 0.35, y: 0.2, w: 12.5, h: 0.7, fontSize: 20, color: WHITE, bold: true });
    const bullets = (slide.bullets || []).map(b => ({ text: b, options: { color: MUTED, fontSize: 14, paraSpaceAfter: 10, bullet: { indent: 15 } } }));
    if (bullets.length > 0) s.addText(bullets, { x: 0.5, y: 1.25, w: 12.3, h: 5.8, valign: 'top' });
    s.addText('SKYSCALE MEDIA  |  CONFIDENTIAL', { x: 0.35, y: 7.1, w: 12.5, h: 0.28, fontSize: 8, color: DIM, charSpacing: 2 });
  }

  const cs = pptx.addSlide();
  cs.background = { color: BG };
  cs.addShape(pptx.ShapeType.rect, { x: 0, y: 2.8, w: 13.33, h: 2, fill: { color: SURFACE } });
  cs.addShape(pptx.ShapeType.rect, { x: 0, y: 2.8, w: 0.5, h: 2, fill: { color: ACCENT } });
  cs.addText('SKYSCALE', { x: 0, y: 2.95, w: 13.33, h: 0.9, fontSize: 52, color: WHITE, bold: true, align: 'center', charSpacing: 8 });
  cs.addText('MEDIA', { x: 0, y: 3.75, w: 13.33, h: 0.6, fontSize: 18, color: ACCENT, align: 'center', charSpacing: 14 });
  cs.addText('skyscalemedia.com', { x: 0, y: 4.7, w: 13.33, h: 0.4, fontSize: 13, color: MUTED, align: 'center' });

  const buf = await pptx.write({ outputType: 'nodebuffer' });
  const filename = (client || 'Skyscale').replace(/[^a-zA-Z0-9]/g, '_') + '_Deck_' + new Date().toISOString().split('T')[0] + '.pptx';
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  res.setHeader('Content-Disposition', 'attachment; filename="' + filename + '"');
  res.send(buf);
                                                          }
