import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import PptxGenJS from 'pptxgenjs';
import * as XLSX from 'xlsx';

const app = express();
const __dirname = path.dirname(fileURLToPath(import.meta.url));

app.use(express.json());
app.use(express.static(__dirname));

const ANTHROPIC_URL = 'https://api.anthropic.com/v1/messages';

async function callClaude(system, prompt, json = false) {
  const body = {
    model: 'claude-sonnet-4-5',
    max_tokens: 4096,
    system,
    messages: [{ role: 'user', content: prompt }],
  };
  const res = await fetch(ANTHROPIC_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': process.env.ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01',
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(await res.text());
  const data = await res.json();
  const text = data.content?.[0]?.text || '';
  if (json) {
    const match = text.match(/```(?:json)?\s*([\s\S]*?)```/) || [null, text];
    return JSON.parse(match[1].trim());
  }
  return text;
}

// Plain chat
app.post('/api/chat', async (req, res) => {
  const { prompt, system } = req.body;
  if (!prompt) return res.status(400).json({ error: 'Prompt required' });
  try {
    const text = await callClaude(system || '', prompt);
    res.json({ content: text });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// DECK → .pptx
app.post('/api/deck', async (req, res) => {
  const { prompt, client } = req.body;
  if (!prompt) return res.status(400).json({ error: 'Prompt required' });

  const system = `You are a presentation architect for Skyscale Media, a Dubai-based OOH media agency.
Return ONLY valid JSON (no markdown, no explanation) matching this exact schema:
{
  "title": "string",
  "subtitle": "string",
  "client": "string",
  "date": "string",
  "slides": [
    {
      "title": "string",
      "bullets": ["string", "string"],
      "note": "string"
    }
  ]
}
Generate 8-12 slides. Bullets should be punchy, data-led, max 12 words each. No fabricated numbers — use ranges or qualitative descriptions when data is unknown.`;

  try {
    const deck = await callClaude(system, `Create a deck for: ${prompt}\nClient: ${client || 'Skyscale'}`, true);

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.theme = { headFontFace: 'Calibri', bodyFontFace: 'Calibri' };

    // Title slide
    const titleSlide = pptx.addSlide();
    titleSlide.background = { color: '080a0f' };
    titleSlide.addText(deck.title || 'Skyscale Deck', {
      x: 0.8, y: 2.2, w: 10, h: 1.4,
      fontSize: 36, bold: true, color: 'FFFFFF', fontFace: 'Calibri',
    });
    titleSlide.addText(deck.subtitle || '', {
      x: 0.8, y: 3.7, w: 10, h: 0.6,
      fontSize: 16, color: '00c2ff', fontFace: 'Calibri',
    });
    titleSlide.addText(`${deck.client || ''} · ${deck.date || new Date().toLocaleDateString('en-GB')}`, {
      x: 0.8, y: 4.4, w: 10, h: 0.4,
      fontSize: 12, color: '888888', fontFace: 'Calibri',
    });

    // Content slides
    for (const slide of (deck.slides || [])) {
      const s = pptx.addSlide();
      s.background = { color: '0d1017' };

      // Title bar
      s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 0.9, fill: { color: '080a0f' } });
      s.addText(slide.title || '', {
        x: 0.5, y: 0.12, w: 12, h: 0.65,
        fontSize: 22, bold: true, color: 'FFFFFF', fontFace: 'Calibri',
      });

      // Accent line
      s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 1.0, w: 1.2, h: 0.06, fill: { color: '00c2ff' } });

      // Bullets
      const bullets = (slide.bullets || []).map(b => ({ text: b, options: { bullet: { type: 'bullet' }, color: 'E8EAF0', fontSize: 15, fontFace: 'Calibri', paraSpaceAfter: 8 } }));
      if (bullets.length) {
        s.addText(bullets, { x: 0.5, y: 1.2, w: 12, h: 4.5, valign: 'top' });
      }

      // Speaker note
      if (slide.note) s.addNotes(slide.note);
    }

    const buf = await pptx.write({ outputType: 'nodebuffer' });
    const filename = `${(deck.client || 'Skyscale').replace(/\s+/g, '_')}_Deck_${Date.now()}.pptx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buf);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// PLAN → .xlsx
app.post('/api/plan', async (req, res) => {
  const { prompt, client } = req.body;
  if (!prompt) return res.status(400).json({ error: 'Prompt required' });

  const system = `You are a media planning expert for Skyscale Media, a Dubai-based OOH media agency.
Return ONLY valid JSON (no markdown, no explanation) matching this exact schema:
{
  "campaign": "string",
  "client": "string",
  "objective": "string",
  "period": "string",
  "markets": [
    {
      "market": "string",
      "format": "string",
      "locations": "string",
      "screens": number,
      "duration": "string",
      "cpm_usd": "string",
      "budget_usd": "string",
      "notes": "string"
    }
  ],
  "total_budget": "string",
  "kpis": ["string"],
  "exclusions": ["string"]
}
Include realistic OOH market data for the requested regions. Use ranges for budgets. No fabricated specifics — use industry-standard estimates.`;

  try {
    const plan = await callClaude(system, `Create a media plan for: ${prompt}\nClient: ${client || 'Skyscale'}`, true);

    const wb = XLSX.utils.book_new();

    // ── Sheet 1: Campaign Overview ──
    const overview = [
      ['SKYSCALE MEDIA — MEDIA PLAN', ''],
      ['', ''],
      ['Campaign', plan.campaign || ''],
      ['Client', plan.client || client || ''],
      ['Objective', plan.objective || ''],
      ['Period', plan.period || ''],
      ['Total Budget', plan.total_budget || ''],
      ['Generated', new Date().toLocaleDateString('en-GB')],
    ];
    const wsOverview = XLSX.utils.aoa_to_sheet(overview);
    wsOverview['!cols'] = [{ wch: 20 }, { wch: 50 }];
    XLSX.utils.book_append_sheet(wb, wsOverview, 'Overview');

    // ── Sheet 2: Market Plan ──
    const headers = ['Market', 'Format', 'Key Locations', 'Screens', 'Duration', 'CPM (USD)', 'Budget (USD)', 'Notes'];
    const rows = (plan.markets || []).map(m => [
      m.market, m.format, m.locations,
      m.screens, m.duration, m.cpm_usd, m.budget_usd, m.notes,
    ]);
    const wsMarkets = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    wsMarkets['!cols'] = [
      { wch: 22 }, { wch: 20 }, { wch: 35 },
      { wch: 10 }, { wch: 14 }, { wch: 12 }, { wch: 18 }, { wch: 30 },
    ];
    XLSX.utils.book_append_sheet(wb, wsMarkets, 'Market Plan');

    // ── Sheet 3: KPIs & Exclusions ──
    const kpiRows = [
      ['KPIs', ''], ['──────', ''],
      ...(plan.kpis || []).map(k => [k, '']),
      ['', ''],
      ['Exclusions', ''], ['──────', ''],
      ...(plan.exclusions || []).map(e => [e, '']),
    ];
    const wsKpi = XLSX.utils.aoa_to_sheet(kpiRows);
    wsKpi['!cols'] = [{ wch: 50 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, wsKpi, 'KPIs & Exclusions');

    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const filename = `${(plan.client || 'Skyscale').replace(/\s+/g, '_')}_MediaPlan_${Date.now()}.xlsx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buf);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
