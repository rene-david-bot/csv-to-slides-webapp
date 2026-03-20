const fileInput = document.getElementById('fileInput');
const outputNameInput = document.getElementById('outputName');
const generateBtn = document.getElementById('generateBtn');
const statusEl = document.getElementById('status');

const STOPWORDS = new Set([
  'the','a','an','and','or','but','if','then','than','that','this','these','those','to','of','in','on','for','with','by','from','at','as',
  'is','are','was','were','be','been','being','it','its','their','them','they','he','she','his','her','we','our','you','your','i','me','my',
  'not','no','yes','can','could','should','would'
]);

const tokenRegex = /[A-Za-zÀ-ÖØ-öø-ÿ0-9']+/g;
const logoMapPromise = fetch('./assets/lance-logos/logo-map.json').then(r => r.ok ? r.json() : {}).catch(() => ({}));
const logoDataCache = new Map();

function setStatus(text, isError = false) {
  statusEl.textContent = text;
  statusEl.style.color = isError ? '#ff8e8e' : '#a7a7a7';
}

function sanitizeFilename(name) {
  const base = (name || 'news-slides.pptx').trim().replace(/[^a-zA-Z0-9._-]/g, '-');
  return base.toLowerCase().endsWith('.pptx') ? base : `${base}.pptx`;
}

function normalizeLanceKey(value) {
  return String(value || '')
    .toUpperCase()
    .replace(/&/g, 'AND')
    .replace(/[‐‑–—]/g, '-')
    .replace(/[^A-Z0-9]+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
}

async function resolveLogoData(lance) {
  const key = normalizeLanceKey(lance);
  if (!key) return null;

  if (logoDataCache.has(key)) return logoDataCache.get(key);

  const candidates = [key];
  if (key.includes('_AND_')) candidates.push(key.replaceAll('_AND_', '_'));
  if (key.endsWith('_AND')) candidates.push(key.slice(0, -4));

  const map = await logoMapPromise;

  for (const cand of candidates) {
    const relPath = map[cand] || `assets/lance-logos/${cand}.png`;
    try {
      const res = await fetch(`./${relPath}`);
      if (!res.ok) continue;
      const blob = await res.blob();
      const dataUri = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      });

      const dims = await new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => resolve({ width: img.naturalWidth || img.width || 1, height: img.naturalHeight || img.height || 1 });
        img.onerror = reject;
        img.src = dataUri;
      });

      const out = { data: dataUri, width: dims.width, height: dims.height };
      logoDataCache.set(key, out);
      return out;
    } catch {
      // try next candidate
    }
  }

  logoDataCache.set(key, null);
  return null;
}

function trimToWordLimit(text, maxWords = 145) {
  const matches = [...text.matchAll(tokenRegex)];
  if (matches.length <= maxWords) return text;
  const cutAt = matches[maxWords - 1].index + matches[maxWords - 1][0].length;
  return text.slice(0, cutAt).trim().replace(/[,:;\-]+$/, '') + '…';
}

function decodeHtmlEntities(value) {
  const str = String(value || '');
  if (!str.includes('&')) return str;
  const textarea = document.createElement('textarea');
  textarea.innerHTML = str;
  return textarea.value;
}

function parseSemicolonCSV(input) {
  const rows = [];
  let row = [];
  let value = '';
  let inQuotes = false;

  for (let i = 0; i < input.length; i++) {
    const ch = input[i];
    const next = input[i + 1];

    if (ch === '"') {
      if (inQuotes && next === '"') {
        value += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (ch === ';' && !inQuotes) {
      row.push(value);
      value = '';
    } else if ((ch === '\n' || ch === '\r') && !inQuotes) {
      if (ch === '\r' && next === '\n') i++;
      row.push(value);
      if (row.some(cell => (cell || '').trim() !== '')) rows.push(row);
      row = [];
      value = '';
    } else {
      value += ch;
    }
  }

  if (value.length || row.length) {
    row.push(value);
    if (row.some(cell => (cell || '').trim() !== '')) rows.push(row);
  }

  if (!rows.length) return [];

  const headers = rows[0].map(h => String(h || '').trim());
  return rows.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, idx) => {
      if (!h) return;
      obj[h] = String(r[idx] ?? '').trim();
    });
    return obj;
  });
}

function pick(row, names) {
  const entries = Object.entries(row || {});
  for (const wanted of names) {
    const exact = row[wanted];
    if (exact !== undefined && exact !== null && String(exact).trim() !== '') return String(exact).trim();
    const found = entries.find(([k]) => k.toLowerCase().trim() === wanted.toLowerCase().trim());
    if (found && String(found[1]).trim() !== '') return String(found[1]).trim();
  }
  return '';
}

function cleanHtmlText(raw) {
  if (!raw) return '';
  const doc = new DOMParser().parseFromString(`<div>${raw}</div>`, 'text/html');
  const root = doc.body;

  root.querySelectorAll('br').forEach(el => el.replaceWith('\n'));
  root.querySelectorAll('p,div,li,ul,ol,h1,h2,h3,h4,h5,h6').forEach(el => {
    el.insertAdjacentText('afterend', '\n\n');
  });

  let text = root.textContent || '';
  text = text.replace(/\r/g, '');
  text = text.replace(/\n{3,}/g, '\n\n');
  return text.trim();
}

function formatDate(value) {
  if (!value) return '-';
  const raw = String(value).trim();
  if (!raw) return '-';

  if (/^\d+(\.\d+)?$/.test(raw)) {
    let num = Number(raw);
    if (num > 1e12) num = num / 1000;
    const d = new Date(num * 1000);
    if (!Number.isNaN(d.getTime())) {
      return new Intl.DateTimeFormat('en-US', { month: 'long', day: '2-digit', year: 'numeric' }).format(d);
    }
  }

  const d = new Date(raw);
  if (!Number.isNaN(d.getTime())) {
    return new Intl.DateTimeFormat('en-US', { month: 'long', day: '2-digit', year: 'numeric' }).format(d);
  }
  return raw;
}

function summarizeWholeText(text, maxWords = 145, targetWords = 140) {
  const cleaned = (text || '').replace(/\s+/g, ' ').trim();
  if (!cleaned) return '';

  const words = cleaned.match(tokenRegex) || [];
  if (words.length <= maxWords) return cleaned;

  const sentences = cleaned.split(/(?<=[.!?])\s+/).map(s => s.trim()).filter(Boolean);
  if (sentences.length <= 2) return trimToWordLimit(cleaned, maxWords);

  const freq = Object.create(null);
  for (const w of (cleaned.toLowerCase().match(tokenRegex) || [])) {
    if (w.length <= 2 || STOPWORDS.has(w)) continue;
    freq[w] = (freq[w] || 0) + 1;
  }

  const scores = sentences.map((sentence, idx) => {
    const sentWords = sentence.toLowerCase().match(tokenRegex) || [];
    const score = sentWords.length
      ? sentWords.reduce((sum, w) => sum + (freq[w] || 0), 0) / sentWords.length
      : 0;
    return { idx, score };
  });

  const selected = new Set();
  const n = sentences.length;
  const thirds = [
    [0, Math.floor(n / 3)],
    [Math.floor(n / 3), Math.floor((2 * n) / 3)],
    [Math.floor((2 * n) / 3), n]
  ];

  for (const [start, end] of thirds) {
    const band = scores.filter(s => s.idx >= start && s.idx < end);
    if (!band.length) continue;
    band.sort((a, b) => b.score - a.score);
    selected.add(band[0].idx);
  }

  const ranked = [...scores].sort((a, b) => b.score - a.score);
  const countWords = indexes => [...indexes].reduce((acc, i) => acc + (sentences[i].match(tokenRegex) || []).length, 0);

  for (const item of ranked) {
    if (selected.has(item.idx)) continue;
    const projected = countWords(new Set([...selected, item.idx]));
    if (projected <= maxWords) selected.add(item.idx);
    if (projected >= targetWords) break;
  }

  const ordered = [...selected].sort((a, b) => a - b);
  const summary = ordered.map(i => sentences[i]).join(' ').trim();
  return trimToWordLimit(summary, maxWords);
}

function dynamicTitleSize(title) {
  const len = (title || '').length;
  if (len <= 70) return 28;
  if (len <= 110) return 24;
  return 21;
}

async function readRowsFromFile(file) {
  const ext = (file.name.split('.').pop() || '').toLowerCase();
  const buffer = await file.arrayBuffer();

  if (ext === 'csv' || ext === 'txt') {
    const text = new TextDecoder('utf-8').decode(buffer);
    return parseSemicolonCSV(text);
  }

  if (ext === 'xls') {
    const maybeText = new TextDecoder('utf-8').decode(buffer.slice(0, 300));
    if (maybeText.includes(';') && maybeText.toLowerCase().includes('title')) {
      const text = new TextDecoder('utf-8').decode(buffer);
      return parseSemicolonCSV(text);
    }
  }

  const workbook = XLSX.read(buffer, { type: 'array' });
  const ws = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: '' });
}

async function generateDeck() {
  const file = fileInput.files?.[0];
  if (!file) {
    setStatus('Please select a file first.', true);
    return;
  }

  generateBtn.disabled = true;
  try {
    setStatus('Reading input file…');
    const rows = await readRowsFromFile(file);
    if (!rows.length) throw new Error('No data rows found in file.');

    setStatus(`Generating ${rows.length} slide(s)…`);

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
    pptx.author = 'csv_to_slides web app';
    pptx.subject = 'CSV/XLSX to PPTX conversion';

    const margin = 0.5;
    const gap = 0.35;
    const slideW = 13.333;
    const slideH = 7.5;
    const fullContentW = slideW - margin * 2;
    const columnsW = fullContentW - gap;
    const leftW = columnsW * 0.6;
    const rightW = columnsW - leftW;
    const xLeft = margin;
    const xRight = xLeft + leftW + gap;
    const headerH = 0.92;
    const titleRatio = 0.80;
    const titleW = fullContentW * titleRatio;
    const logoW = fullContentW - titleW;
    const logoX = xLeft + titleW;
    const metaH = 0.95;
    const contentY = margin + headerH + metaH + 0.12;
    const contentH = slideH - contentY - margin;

    for (const row of rows) {
      const projectId = pick(row, ['Id', 'ID', 'id']);
      const title = pick(row, ['Title']) || '(Untitled)';
      const lance = decodeHtmlEntities(pick(row, ['Associated Lance', 'Associated Lances', 'Associated Lens']) || '-').trim() || '-';
      const deliverable = pick(row, ['Associated Deliverable', 'Deliverable']) || '-';
      const publicationDate = formatDate(pick(row, ['Publication Date', 'Date'])) || '-';
      const bodyRaw = cleanHtmlText(pick(row, ['Text', 'Description', 'Body']));
      const body = summarizeWholeText(bodyRaw, 145, 140);

      const slide = pptx.addSlide();
      slide.background = { color: '000000' };

      // Header: title (left 80%) + logo (right 20%)
      slide.addText(title, {
        x: xLeft,
        y: margin,
        w: titleW,
        h: headerH,
        fontSize: dynamicTitleSize(title),
        bold: true,
        color: 'FFFFFF',
        valign: 'top',
        align: 'left'
      });

      const logoData = await resolveLogoData(lance);
      if (logoData?.data) {
        const scale = Math.min(logoW / logoData.width, headerH / logoData.height);
        const drawW = logoData.width * scale;
        const drawH = logoData.height * scale;
        const drawX = logoX + (logoW - drawW) / 2;
        const drawY = margin + (headerH - drawH) / 2;

        slide.addImage({
          data: logoData.data,
          x: drawX,
          y: drawY,
          w: drawW,
          h: drawH
        });
      }

      // Metadata
      slide.addText(`Associated Lance: ${lance}`, {
        x: xLeft,
        y: margin + headerH,
        w: leftW,
        h: 0.27,
        fontSize: 12,
        color: 'CCCCCC',
        align: 'left'
      });

      slide.addText([
        { text: 'Associated Deliverable: ', options: { color: 'CCCCCC', bold: false } },
        {
          text: deliverable,
          options: {
            color: 'CCCCCC',
            underline: { color: 'CCCCCC', style: 'sng' },
            hyperlink: projectId ? { url: `https://guilds.reply.com/news/${projectId}` } : undefined
          }
        }
      ], {
        x: xLeft,
        y: margin + headerH + 0.3,
        w: leftW,
        h: 0.27,
        fontSize: 12,
        align: 'left'
      });

      slide.addText(`Publication Date: ${publicationDate}`, {
        x: xLeft,
        y: margin + headerH + 0.6,
        w: leftW,
        h: 0.27,
        fontSize: 12,
        color: 'CCCCCC',
        align: 'left'
      });

      // Body text
      slide.addText(body || '', {
        x: xLeft,
        y: contentY,
        w: leftW,
        h: contentH,
        fontSize: 16,
        color: 'FFFFFF',
        align: 'left',
        valign: 'top',
        breakLine: true,
        paraSpaceAfterPt: 5,
        lineSpacingMultiple: 1.2
      });

      // Right placeholder
      slide.addShape(pptx.ShapeType.rect, {
        x: xRight,
        y: contentY,
        w: rightW,
        h: contentH,
        fill: { color: '333333' },
        line: { color: '555555', pt: 1 }
      });

      slide.addText('Image Placeholder', {
        x: xRight,
        y: contentY,
        w: rightW,
        h: contentH,
        fontSize: 18,
        bold: true,
        color: 'CCCCCC',
        align: 'center',
        valign: 'mid'
      });
    }

    const outName = sanitizeFilename(outputNameInput.value);
    setStatus('Preparing download…');
    await pptx.writeFile({ fileName: outName });
    setStatus(`Done. Generated ${rows.length} slide(s): ${outName}`);
  } catch (err) {
    console.error(err);
    setStatus(`Error: ${err.message || err}`, true);
  } finally {
    generateBtn.disabled = false;
  }
}

generateBtn.addEventListener('click', generateDeck);
