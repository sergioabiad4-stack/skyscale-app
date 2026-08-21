import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import PptxGenJS from 'pptxgenjs';
import ExcelJS from 'exceljs';
import { getClientContext } from './drive.js';

const app = express();
const __dirname = path.dirname(fileURLToPath(import.meta.url));

app.use(express.json());
app.use(express.static(__dirname));

const ANTHROPIC_URL = 'https://api.anthropic.com/v1/messages';

// ─── AI helper ────────────────────────────────────────────────────────────────
async function callClaude(system, prompt, wantJson = false) {
  const res = await fetch(ANTHROPIC_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': process.env.ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01',
    },
    body: JSON.stringify({
      model: 'claude-sonnet-4-5',
      max_tokens: 4096,
      system,
      messages: [{ role: 'user', content: prompt }],
    }),
  });
  if (!res.ok) throw new Error(await res.text());
  const data = await res.json();
  const text = data.content?.[0]?.text || '';
  if (!wantJson) return text;
  const match = text.match(/```(?:json)?\s*([\s\S]*?)```/) || [null, text];
  return JSON.parse(match[1].trim());
}

// ─── Plain chat ───────────────────────────────────────────────────────────────
app.post('/api/chat', async (req, res) => {
  const { prompt, system } = req.body;
  if (!prompt) return res.status(400).json({ error: 'Prompt required' });
  try {
    res.json({ content: await callClaude(system || '', prompt) });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── DECK → .pptx ─────────────────────────────────────────────────────────────
app.post('/api/deck', async (req, res) => {
  const { prompt, client } = req.body;
  if (!prompt) return res.status(400).json({ error: 'Prompt required' });

  // Pull Drive context if available
  const driveCtx = client ? await getClientContext(client) : '';
  const driveNote = driveCtx
    ? `Use the following real data from Google Drive files:\n${driveCtx}`
    : 'No live Drive connection — use realistic OOH industry data and ranges.';

  const system = `You are a senior OOH media strategist and presentation designer for Skyscale Media, Dubai.
${driveNote}

Return ONLY valid JSON (no markdown, no explanation) with this exact schema:
{
  "title": "string",
  "subtitle": "string",
  "client": "string",
  "date": "string",
  "slides": [
    {
      "section": "string",
      "title": "string",
      "bullets": [
        { "text": "string", "stat": "string or null" }
      ],
      "callout": "string or null",
      "note": "string"
    }
  ]
}

Generate 10-14 slides covering: Executive Summary, Market Opportunity, Campaign Objectives, Target Audience, Market-by-Market Plan (one slide per key market with real/estimated figures), Format Strategy, Budget & Timeline, KPIs & Measurement, Next Steps.
Each bullet must be specific and data-led (use figures, percentages, market sizes). stat = the key number/metric for that bullet (e.g. "$2.4M", "38%", "12 screens"). callout = one punchy headline stat for the slide.`;

  try {
    const deck = await callClaude(system, `Create a comprehensive OOH media deck for:\nCLIENT: ${client || 'Skyscale'}\nREQUEST: ${prompt}`, true);

    // ── Build PPTX ──
    const RED = 'CC0000';
    const BLACK = '000000';
    const WHITE = 'FFFFFF';
    const LIGHTGRAY = 'F2F2F2';
    const DARKGRAY = '1A1A1A';

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // 13.33 × 7.5 inches
    pptx.author = 'Skyscale Media';
    pptx.company = 'Skyscale Media';
    pptx.subject = deck.title || 'OOH Media Deck';

    // ── Title slide ──
    const ts = pptx.addSlide();
    ts.background = { color: BLACK };
    // Full-bleed red bar left edge
    ts.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.22, h: 7.5, fill: { color: RED } });
    // Red accent line
    ts.addShape(pptx.ShapeType.rect, { x: 0.22, y: 3.3, w: 12.5, h: 0.04, fill: { color: RED } });
    ts.addText('SKYSCALE MEDIA', {
      x: 0.5, y: 1.4, w: 12, h: 0.5,
      fontSize: 11, bold: true, color: RED, fontFace: 'Calibri', charSpacing: 4,
    });
    ts.addText(deck.title || 'OOH Proposal', {
      x: 0.5, y: 2.0, w: 11.5, h: 1.4,
      fontSize: 40, bold: true, color: WHITE, fontFace: 'Calibri',
    });
    ts.addText(deck.subtitle || '', {
      x: 0.5, y: 3.5, w: 11.5, h: 0.6,
      fontSize: 16, color: LIGHTGRAY, fontFace: 'Calibri', italic: true,
    });
    ts.addText(`${deck.client || client || ''} · ${deck.date || new Date().toLocaleDateString('en-GB', { month: 'long', year: 'numeric' })}`, {
      x: 0.5, y: 6.5, w: 12, h: 0.4,
      fontSize: 11, color: RED, fontFace: 'Calibri',
    });
    // Confidential tag
    ts.addText('CONFIDENTIAL', {
      x: 10, y: 6.5, w: 3, h: 0.4,
      fontSize: 9, color: '555555', fontFace: 'Calibri', align: 'right', charSpacing: 2,
    });

    // ── Content slides ──
    for (const slide of (deck.slides || [])) {
      const s = pptx.addSlide();
      s.background = { color: BLACK };

      // Red sidebar
      s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.22, h: 7.5, fill: { color: RED } });

      // Header band
      s.addShape(pptx.ShapeType.rect, { x: 0.22, y: 0, w: 13.11, h: 1.1, fill: { color: DARKGRAY } });

      // Section label
      if (slide.section) {
        s.addText(slide.section.toUpperCase(), {
          x: 0.5, y: 0.08, w: 11, h: 0.3,
          fontSize: 8, bold: true, color: RED, fontFace: 'Calibri', charSpacing: 3,
        });
      }

      // Slide title
      s.addText(slide.title || '', {
        x: 0.5, y: 0.32, w: 11, h: 0.72,
        fontSize: 24, bold: true, color: WHITE, fontFace: 'Calibri',
      });

      // Callout box (right side)
      if (slide.callout) {
        s.addShape(pptx.ShapeType.rect, { x: 10.0, y: 1.3, w: 3.1, h: 5.8, fill: { color: RED } });
        s.addText(slide.callout, {
          x: 10.05, y: 1.4, w: 3.0, h: 5.6,
          fontSize: 13, bold: true, color: WHITE, fontFace: 'Calibri',
          align: 'center', valign: 'middle', wrap: true,
        });
      }

      // Bullets
      const bulletW = slide.callout ? 9.2 : 12.5;
      const bullets = (slide.bullets || []).map(b => {
        const statStr = b.stat ? `  [${b.stat}]` : '';
        return {
          text: [
            { text: '▶  ', options: { color: RED, fontSize: 11 } },
            { text: b.text || '', options: { color: WHITE, fontSize: 13, fontFace: 'Calibri' } },
            ...(b.stat ? [{ text: `  ${b.stat}`, options: { color: RED, fontSize: 13, bold: true, fontFace: 'Calibri' } }] : []),
          ],
          options: { paraSpaceAfter: 10 },
        };
      });

      if (bullets.length) {
        s.addText(bullets.map(b => b.text).flat().map((seg, i) => ({ text: seg.text, options: { ...seg.options } })), {
          x: 0.5, y: 1.25, w: bulletW, h: 5.9,
          valign: 'top', fontFace: 'Calibri', lineSpacingMultiple: 1.3,
        });
      }

      // Speaker notes
      if (slide.note) s.addNotes(slide.note);

      // Slide footer
      s.addShape(pptx.ShapeType.rect, { x: 0.22, y: 7.2, w: 13.11, h: 0.3, fill: { color: DARKGRAY } });
      s.addText('SKYSCALE MEDIA · CONFIDENTIAL', {
        x: 0.5, y: 7.22, w: 12, h: 0.22,
        fontSize: 7, color: '555555', fontFace: 'Calibri', charSpacing: 2,
      });
    }

    // ── End slide ──
    const es = pptx.addSlide();
    es.background = { color: RED };
    es.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.22, h: 7.5, fill: { color: BLACK } });
    es.addText('THANK YOU', {
      x: 1, y: 2.8, w: 11, h: 1.5,
      fontSize: 48, bold: true, color: WHITE, fontFace: 'Calibri',
    });
    es.addText('skyscalemedia.com · sergio@skyscalemedia.com', {
      x: 1, y: 4.5, w: 11, h: 0.5,
      fontSize: 13, color: WHITE, fontFace: 'Calibri', italic: true,
    });

    const buf = await pptx.write({ outputType: 'nodebuffer' });
    const filename = `${(deck.client || client || 'Skyscale').replace(/\s+/g, '_')}_Deck_${Date.now()}.pptx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buf);
  } catch (e) {
    console.error('DECK error:', e);
    res.status(500).json({ error: e.message });
  }
});

// ─── PLAN → .xlsx ─────────────────────────────────────────────────────────────
app.post('/api/plan', async (req, res) => {
  const { prompt, client } = req.body;
  if (!prompt) return res.status(400).json({ error: 'Prompt required' });

  const driveCtx = client ? await getClientContext(client) : '';
  const driveNote = driveCtx
    ? `Use the following real data from Google Drive files:\n${driveCtx}`
    : 'No live Drive connection — use realistic OOH industry rates and budgets.';

  const system = `You are a senior media planner for Skyscale Media, Dubai.
${driveNote}

Return ONLY valid JSON (no markdown) with this exact schema:
{
  "campaign": "string",
  "client": "string",
  "objective": "string",
  "period": "string",
  "markets": [
    {
      "tier": "1|2|3",
      "market": "string",
      "city": "string",
      "format": "string",
      "operator": "string",
      "locations": "string",
      "screens": number,
      "spots_per_hour": number,
      "duration_weeks": number,
      "impressions_000": number,
      "cpm_usd": number,
      "net_rate_usd": number,
      "quantity": number,
      "total_usd": number,
      "notes": "string"
    }
  ],
  "summary": {
    "total_markets": number,
    "total_screens": number,
    "total_impressions_000": number,
    "total_budget_usd": number,
    "blended_cpm_usd": number
  },
  "kpis": ["string"],
  "exclusions": ["string"],
  "approved_by": ""
}
All numbers must be realistic OOH industry figures. Use actual market rates for the regions requested.`;

  try {
    const plan = await callClaude(system, `Create a detailed media plan for:\nCLIENT: ${client || 'Skyscale'}\nREQUEST: ${prompt}`, true);

    const wb = new ExcelJS.Workbook();
    wb.creator = 'Skyscale Media Intelligence Surface';
    wb.created = new Date();

    const RED = 'FFCC0000';
    const BLACK = 'FF000000';
    const WHITE = 'FFFFFFFF';
    const LIGHTRED = 'FFFFE5E5';
    const GRAY = 'FFF5F5F5';
    const DARKGRAY = 'FF3A3A3A';
    const BORDER_COLOR = 'FFCCCCCC';

    const thinBorder = {
      top: { style: 'thin', color: { argb: BORDER_COLOR } },
      left: { style: 'thin', color: { argb: BORDER_COLOR } },
      bottom: { style: 'thin', color: { argb: BORDER_COLOR } },
      right: { style: 'thin', color: { argb: BORDER_COLOR } },
    };
    const thickBorder = {
      top: { style: 'medium', color: { argb: 'FFCC0000' } },
      left: { style: 'medium', color: { argb: 'FFCC0000' } },
      bottom: { style: 'medium', color: { argb: 'FFCC0000' } },
      right: { style: 'medium', color: { argb: 'FFCC0000' } },
    };

    function styleHeader(cell, bgArgb = RED, fgArgb = WHITE) {
      cell.font = { bold: true, color: { argb: fgArgb }, name: 'Calibri', size: 10 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.border = thinBorder;
    }

    function styleCell(cell, numeric = false, shade = false) {
      cell.font = { name: 'Calibri', size: 10, color: { argb: BLACK } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: shade ? GRAY : WHITE } };
      cell.alignment = { vertical: 'middle', horizontal: numeric ? 'right' : 'left', wrapText: true };
      cell.border = thinBorder;
    }

    // ── Sheet 1: Cover ──────────────────────────────────────────────────────
    const cover = wb.addWorksheet('Cover');
    cover.columns = [{ width: 30 }, { width: 50 }];

    const coverData = [
      ['SKYSCALE MEDIA', ''],
      ['MEDIA PLAN', ''],
      ['', ''],
      ['Campaign', plan.campaign || ''],
      ['Client', plan.client || client || ''],
      ['Objective', plan.objective || ''],
      ['Campaign Period', plan.period || ''],
      ['Date Generated', new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'long', year: 'numeric' })],
      ['', ''],
      ['SUMMARY', ''],
      ['Total Markets', plan.summary?.total_markets ?? ''],
      ['Total Screens', plan.summary?.total_screens ?? ''],
      ['Total Impressions (000)', plan.summary?.total_impressions_000?.toLocaleString() ?? ''],
      ['Total Budget (USD)', `$${(plan.summary?.total_budget_usd ?? 0).toLocaleString()}`],
      ['Blended CPM (USD)', `$${plan.summary?.blended_cpm_usd ?? ''}`],
      ['', ''],
      ['Prepared by', 'Skyscale Media'],
      ['Approved by', plan.approved_by || ''],
    ];

    coverData.forEach((row, i) => {
      const r = cover.addRow(row);
      r.height = 20;
      const [a, b] = [r.getCell(1), r.getCell(2)];

      if (i === 0) { // SKYSCALE MEDIA
        a.font = { bold: true, size: 18, color: { argb: 'FFCC0000' }, name: 'Calibri' };
        a.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BLACK } };
        b.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BLACK } };
        r.height = 36;
      } else if (i === 1) { // MEDIA PLAN
        a.font = { bold: true, size: 14, color: { argb: WHITE }, name: 'Calibri' };
        a.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCC0000' } };
        b.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCC0000' } };
        r.height = 28;
      } else if (row[0] === 'SUMMARY') {
        a.font = { bold: true, size: 11, color: { argb: WHITE }, name: 'Calibri' };
        a.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3A3A3A' } };
        b.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3A3A3A' } };
      } else if (row[0] && row[1] !== '') {
        a.font = { bold: true, size: 10, name: 'Calibri', color: { argb: 'FF3A3A3A' } };
        a.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GRAY } };
        b.font = { size: 10, name: 'Calibri' };
        b.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: WHITE } };
        [a, b].forEach(c => c.border = thinBorder);
      }
    });

    // ── Sheet 2: Market Plan ────────────────────────────────────────────────
    const ws = wb.addWorksheet('Market Plan');
    ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 3 }];

    // Row 1: campaign title banner
    ws.mergeCells('A1:O1');
    const banner = ws.getCell('A1');
    banner.value = `${plan.client || client || 'Skyscale Media'} — ${plan.campaign || 'Media Plan'} · ${plan.period || ''}`;
    banner.font = { bold: true, size: 13, color: { argb: WHITE }, name: 'Calibri' };
    banner.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BLACK } };
    banner.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    ws.getRow(1).height = 28;

    // Row 2: sub-header
    ws.mergeCells('A2:O2');
    const sub = ws.getCell('A2');
    sub.value = `Generated by Skyscale Intelligence Surface · ${new Date().toLocaleDateString('en-GB')}`;
    sub.font = { size: 9, color: { argb: WHITE }, italic: true, name: 'Calibri' };
    sub.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCC0000' } };
    sub.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    ws.getRow(2).height = 18;

    // Row 3: column headers
    const headers = [
      'Tier', 'Market', 'City', 'Format', 'Operator',
      'Key Locations', 'Screens', 'Spots/Hr', 'Weeks',
      'Impressions (000)', 'CPM (USD)', 'Net Rate (USD)', 'Qty', 'Total (USD)', 'Notes',
    ];
    const headerRow = ws.getRow(3);
    headers.forEach((h, i) => {
      const cell = headerRow.getCell(i + 1);
      cell.value = h;
      styleHeader(cell);
    });
    headerRow.height = 32;

    // Column widths
    ws.columns = [
      { width: 7 },  // Tier
      { width: 18 }, // Market
      { width: 14 }, // City
      { width: 18 }, // Format
      { width: 16 }, // Operator
      { width: 28 }, // Locations
      { width: 9 },  // Screens
      { width: 9 },  // Spots/Hr
      { width: 8 },  // Weeks
      { width: 16 }, // Impressions
      { width: 12 }, // CPM
      { width: 14 }, // Net Rate
      { width: 7 },  // Qty
      { width: 15 }, // Total
      { width: 30 }, // Notes
    ];

    // Data rows
    let lastTier = '';
    (plan.markets || []).forEach((m, idx) => {
      const shade = idx % 2 === 1;
      const r = ws.addRow([
        m.tier, m.market, m.city, m.format, m.operator,
        m.locations, m.screens, m.spots_per_hour, m.duration_weeks,
        m.impressions_000, m.cpm_usd, m.net_rate_usd, m.quantity, m.total_usd, m.notes,
      ]);
      r.height = 22;
      r.eachCell({ includeEmpty: true }, (cell, colNum) => {
        const numeric = [7, 8, 9, 10, 11, 12, 13, 14].includes(colNum);
        styleCell(cell, numeric, shade);
        // Format numbers
        if ([11, 12, 14].includes(colNum) && typeof cell.value === 'number') {
          cell.numFmt = '"$"#,##0';
        }
        if (colNum === 10 && typeof cell.value === 'number') {
          cell.numFmt = '#,##0';
        }
        // Tier colour coding
        if (colNum === 1) {
          const tierColors = { '1': 'FFCC0000', '2': 'FF3A3A3A', '3': 'FF777777' };
          cell.font = { bold: true, color: { argb: WHITE }, name: 'Calibri', size: 10 };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: tierColors[m.tier] || 'FF777777' } };
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
        }
      });
    });

    // Totals row
    const totals = ws.addRow([
      '', 'TOTAL', '', '', '', '',
      plan.summary?.total_screens ?? '',
      '', '',
      plan.summary?.total_impressions_000?.toLocaleString() ?? '',
      `$${plan.summary?.blended_cpm_usd ?? ''}`,
      '', '',
      plan.summary?.total_budget_usd ?? '',
      '',
    ]);
    totals.height = 26;
    totals.eachCell({ includeEmpty: true }, (cell, colNum) => {
      cell.font = { bold: true, size: 10, name: 'Calibri', color: { argb: WHITE } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BLACK } };
      cell.border = thickBorder;
      cell.alignment = { vertical: 'middle', horizontal: colNum > 6 ? 'right' : 'left' };
      if (colNum === 14 && typeof cell.value === 'number') cell.numFmt = '"$"#,##0';
    });

    // ── Sheet 3: KPIs ───────────────────────────────────────────────────────
    const kpiWs = wb.addWorksheet('KPIs & Exclusions');
    kpiWs.columns = [{ width: 6 }, { width: 55 }];

    const kpiHeader = kpiWs.addRow(['#', 'KPI']);
    [kpiHeader.getCell(1), kpiHeader.getCell(2)].forEach(c => styleHeader(c));
    kpiWs.getRow(1).height = 26;

    (plan.kpis || []).forEach((k, i) => {
      const r = kpiWs.addRow([i + 1, k]);
      r.height = 20;
      styleCell(r.getCell(1), true, i % 2 === 1);
      styleCell(r.getCell(2), false, i % 2 === 1);
    });

    kpiWs.addRow([]);

    const exHeader = kpiWs.addRow(['#', 'Exclusion']);
    [exHeader.getCell(1), exHeader.getCell(2)].forEach(c => styleHeader(c, 'FF3A3A3A', WHITE));
    exHeader.height = 26;

    (plan.exclusions || []).forEach((e, i) => {
      const r = kpiWs.addRow([i + 1, e]);
      r.height = 20;
      styleCell(r.getCell(1), true, i % 2 === 1);
      styleCell(r.getCell(2), false, i % 2 === 1);
    });

    const buf = await wb.xlsx.writeBuffer();
    const filename = `${(plan.client || client || 'Skyscale').replace(/\s+/g, '_')}_MediaPlan_${Date.now()}.xlsx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buf);
  } catch (e) {
    console.error('PLAN error:', e);
    res.status(500).json({ error: e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
