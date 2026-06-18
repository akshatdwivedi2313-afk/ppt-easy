/* SRDM PPT Image to Editable Excel/PPT - v4
   - GitHub Pages static app
   - Local PptxGenJS, JSZip, XLSX libraries
   - OCR via Tesseract.js in browser
*/

const els = {
  pptFile: document.getElementById('pptFile'),
  fileName: document.getElementById('fileName'),
  convertBtn: document.getElementById('convertBtn'),
  resetBtn: document.getElementById('resetBtn'),
  testPptBtn: document.getElementById('testPptBtn'),
  progressWrap: document.getElementById('progressWrap'),
  progressBar: document.getElementById('progressBar'),
  status: document.getElementById('status'),
  previewBox: document.getElementById('previewBox'),
  downloadPanel: document.getElementById('downloadPanel'),
  downloadExcelBtn: document.getElementById('downloadExcelBtn'),
  downloadPptBtn: document.getElementById('downloadPptBtn'),
};

let slideResults = [];
let lastBaseName = 'converted';

els.pptFile.addEventListener('change', () => {
  const f = els.pptFile.files && els.pptFile.files[0];
  els.fileName.textContent = f ? f.name : 'PPTX choose karo';
});
els.convertBtn.addEventListener('click', convertNow);
els.resetBtn.addEventListener('click', resetAll);
els.testPptBtn.addEventListener('click', makeTestPpt);
els.downloadExcelBtn.addEventListener('click', () => downloadExcel(lastBaseName));
els.downloadPptBtn.addEventListener('click', () => downloadPpt(lastBaseName));

window.addEventListener('load', () => {
  const missing = [];
  if (!window.JSZip) missing.push('JSZip');
  if (!window.XLSX) missing.push('XLSX');
  if (!window.PptxGenJS) missing.push('PptxGenJS');
  if (!window.Tesseract) missing.push('Tesseract');
  if (missing.length) setStatus(0, 'Library missing: ' + missing.join(', '));
});

function resetAll() {
  els.pptFile.value = '';
  els.fileName.textContent = 'PPTX choose karo';
  slideResults = [];
  lastBaseName = 'converted';
  els.previewBox.className = 'empty';
  els.previewBox.innerHTML = 'PPT upload karke Convert / OCR Now दबाओ.';
  els.downloadPanel.classList.add('hidden');
  els.progressWrap.classList.add('hidden');
  setStatus(0, 'Ready');
}

function setStatus(percent, msg) {
  els.progressWrap.classList.remove('hidden');
  els.progressBar.style.width = Math.max(0, Math.min(100, percent)) + '%';
  els.status.textContent = msg;
}

function requireLibraries() {
  const missing = [];
  if (!window.JSZip) missing.push('JSZip');
  if (!window.XLSX) missing.push('XLSX');
  if (!window.PptxGenJS) missing.push('PptxGenJS');
  if (!window.Tesseract) missing.push('Tesseract');
  if (missing.length) {
    throw new Error('Required local libraries load nahi hui: ' + missing.join(', ') + '. Repo me libs folder check karo.');
  }
}

async function convertNow() {
  const file = els.pptFile.files && els.pptFile.files[0];
  if (!file) {
    alert('Pehle PPTX upload karo.');
    return;
  }

  els.convertBtn.disabled = true;
  els.downloadPanel.classList.add('hidden');
  slideResults = [];
  els.previewBox.className = '';
  els.previewBox.innerHTML = '';
  lastBaseName = cleanBaseName(file.name);

  try {
    requireLibraries();
    setStatus(3, 'PPTX read ho rahi hai...');
    const zip = await JSZip.loadAsync(await file.arrayBuffer());

    setStatus(8, 'Slides ke images find ho rahe hain...');
    const imagePolicy = document.getElementById('imagePolicy').value;
    const slideImages = imagePolicy === 'all-media'
      ? await extractAllMediaImages(zip)
      : await extractMainSlideImages(zip);

    if (!slideImages.length) throw new Error('Is PPTX me slide image nahi mili. Ye tool image-only PPT ke liye hai.');

    const lang = document.getElementById('ocrLang').value;
    const gridCols = Number(document.getElementById('gridCols').value || 28);
    const excelMode = document.getElementById('excelMode').value;

    for (let i = 0; i < slideImages.length; i++) {
      const img = slideImages[i];
      const basePct = 10 + Math.round(i * 74 / slideImages.length);
      setStatus(basePct, `Slide ${i + 1}/${slideImages.length} OCR chal raha hai...`);
      const imageSize = await getImageSize(img.dataUrl);
      const ocr = await runOCR(img.dataUrl, lang, (m) => {
        if (m && m.status && typeof m.progress === 'number') {
          const p = basePct + Math.round(m.progress * (74 / slideImages.length));
          setStatus(p, `Slide ${i + 1}/${slideImages.length}: ${m.status} ${Math.round(m.progress * 100)}%`);
        }
      });

      const lineBoxes = groupWordsIntoLines(ocr.words);
      const visualRows = buildVisualGridRows(ocr.words, imageSize.width, gridCols);
      const autoRows = groupWordsAsAutoTable(ocr.words);
      const lineRows = lineBoxes.map(l => [l.text]);
      const rows = excelMode === 'lines' ? lineRows : excelMode === 'auto' ? autoRows : visualRows;
      const text = ocr.text || lineBoxes.map(l => l.text).join('\n');

      const result = {
        slideNo: img.slideNo || i + 1,
        imageName: img.imageName || `slide_${i + 1}`,
        dataUrl: img.dataUrl,
        imageWidth: imageSize.width,
        imageHeight: imageSize.height,
        words: ocr.words,
        text,
        lineBoxes,
        rows,
        gridCols,
        title: detectTitle(lineBoxes, i + 1)
      };
      slideResults.push(result);
      addPreview(result, slideResults.length - 1);
    }

    els.downloadPanel.classList.remove('hidden');
    setStatus(90, 'OCR complete. Output generate ho raha hai...');

    const out = document.getElementById('outputType').value;
    if (document.getElementById('autoDownload').checked) {
      if (out === 'excel' || out === 'both') downloadExcel(lastBaseName);
      if (out === 'ppt' || out === 'both') await downloadPpt(lastBaseName);
    }
    setStatus(100, 'Done. Download buttons ready hain.');
  } catch (err) {
    console.error(err);
    setStatus(100, 'Error: ' + (err.message || err));
  } finally {
    els.convertBtn.disabled = false;
  }
}

function cleanBaseName(name) {
  return String(name || 'converted')
    .replace(/\.pptx$/i, '')
    .replace(/[^a-zA-Z0-9_\-]+/g, '_')
    .slice(0, 90) || 'converted';
}

async function runOCR(dataUrl, lang, logger) {
  const options = {
    logger,
    workerPath: './libs/worker.min.js',
    corePath: './libs/tesseract-core-simd-lstm.wasm.js',
    // traineddata language files are downloaded by browser; cached after first run
    langPath: 'https://tessdata.projectnaptha.com/4.0.0'
  };
  const result = await Tesseract.recognize(dataUrl, lang, options);
  return {
    text: result?.data?.text || '',
    words: normalizeWords(result?.data?.words || [])
  };
}

function normalizeWords(words) {
  return words.map(w => {
    const b = w.bbox || {};
    const x0 = Number(b.x0 ?? b.left ?? 0);
    const y0 = Number(b.y0 ?? b.top ?? 0);
    const x1 = Number(b.x1 ?? b.right ?? 0);
    const y1 = Number(b.y1 ?? b.bottom ?? 0);
    return {
      text: String(w.text || '').trim(),
      confidence: Number(w.confidence || 0),
      x0, y0, x1, y1,
      cx: (x0 + x1) / 2,
      cy: (y0 + y1) / 2
    };
  }).filter(w => w.text && (w.x1 - w.x0) > 1 && (w.y1 - w.y0) > 1);
}

async function extractMainSlideImages(zip) {
  const slidePaths = Object.keys(zip.files)
    .filter(p => /^ppt\/slides\/slide\d+\.xml$/i.test(p))
    .sort((a, b) => slideNumber(a) - slideNumber(b));
  const out = [];

  for (const slidePath of slidePaths) {
    const no = slideNumber(slidePath);
    const relPath = `ppt/slides/_rels/slide${no}.xml.rels`;
    const slideXml = await zip.file(slidePath)?.async('text');
    const relXml = await zip.file(relPath)?.async('text');
    if (!slideXml || !relXml) continue;

    const relMap = parseRelMap(relXml);
    const imgs = extractImageRids(slideXml).map(img => ({ ...img, path: normalizeTargetPath('ppt/slides', relMap[img.rid] || '') }))
      .filter(x => /^ppt\/media\//i.test(x.path) && zip.file(x.path));
    imgs.sort((a, b) => (b.area || 0) - (a.area || 0));
    const chosen = imgs[0];
    if (chosen) {
      const blob = await zip.file(chosen.path).async('blob');
      out.push({ slideNo: no, imageName: chosen.path.split('/').pop(), dataUrl: await blobToDataURL(blob) });
    }
  }

  if (!out.length) return await extractAllMediaImages(zip);
  return out;
}

async function extractAllMediaImages(zip) {
  const media = Object.keys(zip.files)
    .filter(p => /^ppt\/media\/.*\.(png|jpg|jpeg)$/i.test(p))
    .sort(naturalSort);
  const out = [];
  for (let i = 0; i < media.length; i++) {
    const blob = await zip.file(media[i]).async('blob');
    out.push({ slideNo: i + 1, imageName: media[i].split('/').pop(), dataUrl: await blobToDataURL(blob) });
  }
  return out;
}

function parseRelMap(xml) {
  const map = {};
  const re = /<Relationship\s+[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*>/gi;
  let m;
  while ((m = re.exec(xml))) map[m[1]] = m[2];
  return map;
}
function extractImageRids(xml) {
  const list = [];
  const re = /<a:blip[^>]*(?:r:embed|r:link)="([^"]+)"[^>]*>/gi;
  let m;
  while ((m = re.exec(xml))) {
    const around = xml.slice(Math.max(0, m.index - 2000), Math.min(xml.length, m.index + 2500));
    const ext = around.match(/<a:ext[^>]*cx="(\d+)"[^>]*cy="(\d+)"/i);
    list.push({ rid: m[1], area: ext ? Number(ext[1]) * Number(ext[2]) : 0 });
  }
  return list;
}
function normalizeTargetPath(base, target) {
  if (!target) return '';
  if (target.startsWith('/')) return target.slice(1);
  const parts = `${base}/${target}`.split('/');
  const stack = [];
  for (const p of parts) {
    if (!p || p === '.') continue;
    if (p === '..') stack.pop();
    else stack.push(p);
  }
  return stack.join('/');
}
function slideNumber(path) { const m = path.match(/slide(\d+)\.xml/i); return m ? Number(m[1]) : 0; }
function naturalSort(a, b) { return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }); }
function blobToDataURL(blob) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = reject;
    r.readAsDataURL(blob);
  });
}
function getImageSize(dataUrl) {
  return new Promise(resolve => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth || img.width || 1600, height: img.naturalHeight || img.height || 900 });
    img.onerror = () => resolve({ width: 1600, height: 900 });
    img.src = dataUrl;
  });
}

function median(arr) {
  if (!arr.length) return 0;
  const a = [...arr].sort((x, y) => x - y);
  const mid = Math.floor(a.length / 2);
  return a.length % 2 ? a[mid] : (a[mid - 1] + a[mid]) / 2;
}

function groupWordsIntoLines(words) {
  const sorted = [...words].sort((a, b) => (a.cy - b.cy) || (a.x0 - b.x0));
  const heights = sorted.map(w => Math.max(8, w.y1 - w.y0));
  const yTol = Math.max(8, (median(heights) || 14) * 0.72);
  const groups = [];
  for (const w of sorted) {
    let best = null, bestD = Infinity;
    for (const g of groups) {
      const d = Math.abs(g.cy - w.cy);
      if (d < yTol && d < bestD) { best = g; bestD = d; }
    }
    if (best) {
      best.words.push(w);
      best.cy = best.words.reduce((s, x) => s + x.cy, 0) / best.words.length;
    } else groups.push({ cy: w.cy, words: [w] });
  }
  return groups.map(g => {
    const ws = g.words.sort((a, b) => a.x0 - b.x0);
    return {
      text: ws.map(w => w.text).join(' ').replace(/\s+/g, ' ').trim(),
      words: ws,
      x0: Math.min(...ws.map(w => w.x0)),
      y0: Math.min(...ws.map(w => w.y0)),
      x1: Math.max(...ws.map(w => w.x1)),
      y1: Math.max(...ws.map(w => w.y1)),
      cx: ws.reduce((s, w) => s + w.cx, 0) / ws.length,
      cy: g.cy
    };
  }).filter(l => l.text).sort((a, b) => (a.y0 - b.y0) || (a.x0 - b.x0));
}

function buildVisualGridRows(words, imageWidth, colCount) {
  const lines = groupWordsIntoLines(words);
  const rows = [];
  for (const line of lines) {
    const row = Array.from({ length: colCount }, () => '');
    for (const w of line.words) {
      let col = Math.floor((w.cx / Math.max(1, imageWidth)) * colCount);
      col = Math.max(0, Math.min(colCount - 1, col));
      row[col] = row[col] ? `${row[col]} ${w.text}` : w.text;
    }
    trimSparseRow(row);
    rows.push(row.map(cleanCell));
  }
  return rows;
}

function trimSparseRow(row) {
  // keep the same column count for visual alignment; do not trim fully
  return row;
}

function groupWordsAsAutoTable(words) {
  const lines = groupWordsIntoLines(words);
  const rows = [];
  for (const line of lines) {
    const ws = line.words;
    if (!ws.length) continue;
    const gaps = [];
    for (let i = 1; i < ws.length; i++) gaps.push(Math.max(0, ws[i].x0 - ws[i - 1].x1));
    const medGap = median(gaps.filter(x => x > 2)) || 18;
    const splitGap = Math.max(24, medGap * 2.15);
    const cells = [];
    let cur = [];
    for (let i = 0; i < ws.length; i++) {
      if (i > 0) {
        const gap = Math.max(0, ws[i].x0 - ws[i - 1].x1);
        if (gap > splitGap) { cells.push(cur.map(x => x.text).join(' ')); cur = []; }
      }
      cur.push(ws[i]);
    }
    if (cur.length) cells.push(cur.map(x => x.text).join(' '));
    rows.push(cells.map(cleanCell));
  }
  return rows;
}
function cleanCell(s) { return String(s || '').replace(/\s+/g, ' ').trim(); }

function detectTitle(lines, idx) {
  const l = (lines || []).find(x => x.text && x.text.length > 4);
  return l ? l.text.slice(0, 90) : `Slide ${idx}`;
}

function rowsToTsv(rows) {
  return (rows || []).map(r => r.map(v => String(v ?? '').replace(/\t/g, ' ').replace(/\r?\n/g, ' ')).join('\t')).join('\n');
}
function tsvToRows(text) {
  return String(text || '').split(/\r?\n/).map(line => line.split('\t').map(cleanCell)).filter(r => r.some(c => c));
}
function collectEditedRows() {
  slideResults.forEach((s, i) => {
    const ta = document.getElementById(`tsv_${i}`);
    if (ta) s.rows = tsvToRows(ta.value);
  });
}

function addPreview(result, idx) {
  const card = document.createElement('div');
  card.className = 'slide-card';
  card.innerHTML = `
    <h4>Slide ${result.slideNo}</h4>
    <img src="${result.dataUrl}" alt="Slide ${result.slideNo}" />
    <div class="slide-meta">
      <span>Words: ${result.words.length}</span>
      <span>Rows: ${result.rows.length}</span>
      <span>Image: ${result.imageWidth}×${result.imageHeight}</span>
    </div>
    <textarea class="tsv" id="tsv_${idx}" spellcheck="false"></textarea>
  `;
  card.querySelector('textarea').value = rowsToTsv(result.rows);
  els.previewBox.appendChild(card);
}

function safeSheetName(name) { return String(name).replace(/[\\/*?:\[\]]/g, ' ').slice(0, 31); }
function normalizeSheetRows(rows) {
  const max = Math.max(1, ...rows.map(r => r.length));
  return rows.map(r => {
    const a = [...r];
    while (a.length < max) a.push('');
    return a;
  });
}

function downloadExcel(baseName) {
  if (!slideResults.length) { alert('Pehle OCR/Convert karo.'); return; }
  collectEditedRows();
  const wb = XLSX.utils.book_new();
  wb.Props = { Title: `${baseName} editable OCR`, Author: 'SRDM PPT Image Converter' };
  slideResults.forEach((s, idx) => {
    const rows = normalizeSheetRows(s.rows && s.rows.length ? s.rows : [[s.text || '']]);
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const maxCols = Math.max(1, ...rows.map(r => r.length));
    ws['!cols'] = Array.from({ length: maxCols }, (_, c) => ({ wch: c === 0 ? 12 : 10 }));
    ws['!rows'] = rows.map(() => ({ hpt: 18 }));
    XLSX.utils.book_append_sheet(wb, ws, safeSheetName(`Slide ${idx + 1}`));
  });
  XLSX.writeFile(wb, `${baseName}_editable_ocr.xlsx`);
}

async function downloadPpt(baseName) {
  if (!slideResults.length) { alert('Pehle OCR/Convert karo.'); return; }
  collectEditedRows();
  const pptMode = document.getElementById('pptMode').value;
  const includeImage = document.getElementById('includeImage').checked;
  const blueTheme = document.getElementById('blueTheme').checked;

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'SRDM PPT Image Converter';
  pptx.subject = 'Image PPT converted to editable text';
  pptx.title = baseName;
  pptx.company = 'SRDM ZP SATNA';
  pptx.lang = 'hi-IN';
  pptx.theme = {
    headFontFace: 'Nirmala UI',
    bodyFontFace: 'Nirmala UI',
    lang: 'hi-IN'
  };

  for (const s of slideResults) {
    const slide = pptx.addSlide();
    slide.background = { color: blueTheme ? 'F4F8FF' : 'FFFFFF' };

    if (includeImage && s.dataUrl) {
      slide.addImage({ data: s.dataUrl, x: 0, y: 0, w: 13.333, h: 7.5 });
      slide.addShape(pptx.ShapeType?.rect || 'rect', { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: 'FFFFFF', transparency: 30 }, line: { color: 'FFFFFF', transparency: 100 } });
    }

    if (blueTheme) {
      slide.addShape(pptx.ShapeType?.rect || 'rect', { x: 0, y: 0, w: 13.333, h: 0.38, fill: { color: '0B3D73' }, line: { color: '0B3D73' } });
    }

    if (pptMode === 'layout') addSameLayoutText(slide, s, blueTheme);
    else if (pptMode === 'table') addPptTable(slide, s, blueTheme);
    else addOneTextBox(slide, s, blueTheme);

    slide.addText('Converted editable output • SRDM ZP SATNA', {
      x: 0.25, y: 7.18, w: 12.8, h: 0.18,
      fontFace: 'Nirmala UI', fontSize: 7, color: includeImage ? '444444' : '666666', align: 'center', margin: 0
    });
  }

  await pptx.writeFile({ fileName: `${baseName}_editable.pptx` });
}

function addSameLayoutText(slide, s, blueTheme) {
  const W = 13.333, H = 7.5;
  const imgW = Math.max(1, s.imageWidth || 1600), imgH = Math.max(1, s.imageHeight || 900);
  const lines = s.lineBoxes && s.lineBoxes.length ? s.lineBoxes : groupWordsIntoLines(s.words || []);
  if (!lines.length) return addOneTextBox(slide, s, blueTheme);
  for (const l of lines) {
    const x = clamp((l.x0 / imgW) * W - 0.015, 0.02, W - 0.2);
    const y = clamp((l.y0 / imgH) * H - 0.01, 0.02, H - 0.15);
    const w = clamp(((l.x1 - l.x0) / imgW) * W + 0.12, 0.2, W - x - 0.03);
    const h = clamp(((l.y1 - l.y0) / imgH) * H + 0.04, 0.10, 0.35);
    const fs = clamp(h * 55, 4.8, 14);
    const top = y < 0.45;
    slide.addText(l.text, {
      x, y, w, h,
      fontFace: 'Nirmala UI',
      fontSize: fs,
      bold: top || isLikelyHeader(l.text),
      color: top && blueTheme ? 'FFFFFF' : '111111',
      margin: 0,
      breakLine: false,
      fit: 'shrink',
      valign: 'mid',
      align: guessAlign(l.text, x, w)
    });
  }
}

function addPptTable(slide, s, blueTheme) {
  const rows = normalizeSheetRows(s.rows || []).slice(0, 48);
  if (!rows.length) return addOneTextBox(slide, s, blueTheme);
  slide.addText(s.title || `Slide ${s.slideNo}`, {
    x: 0.2, y: 0.05, w: 12.9, h: 0.28,
    fontFace: 'Nirmala UI', fontSize: 13, bold: true,
    color: blueTheme ? 'FFFFFF' : '0B3D73', align: 'center', margin: 0
  });
  const maxCols = Math.max(1, ...rows.map(r => r.length));
  const tableRows = rows.map((r, ri) => {
    const arr = [...r];
    while (arr.length < maxCols) arr.push('');
    return arr.map(cell => ({
      text: String(cell || ''),
      options: {
        fontFace: 'Nirmala UI',
        fontSize: maxCols > 24 ? 3.6 : maxCols > 18 ? 4.5 : maxCols > 12 ? 5.5 : 6.5,
        color: ri === 0 && blueTheme ? 'FFFFFF' : '111111',
        bold: ri === 0,
        fill: ri === 0 && blueTheme ? { color: '0B3D73' } : { color: 'FFFFFF' },
        valign: 'mid',
        align: 'center',
        margin: 0.02
      }
    }));
  });
  slide.addTable(tableRows, {
    x: 0.15, y: 0.52, w: 13.03, h: 6.55,
    border: { type: 'solid', color: 'AFC3D7', pt: 0.35 },
    valign: 'mid',
    fit: 'shrink'
  });
}

function addOneTextBox(slide, s, blueTheme) {
  slide.addText(s.title || `Slide ${s.slideNo}`, {
    x: 0.2, y: 0.05, w: 12.9, h: 0.3,
    fontFace: 'Nirmala UI', fontSize: 14, bold: true,
    color: blueTheme ? 'FFFFFF' : '0B3D73', align: 'center', margin: 0
  });
  slide.addText(s.text || rowsToTsv(s.rows || []), {
    x: 0.35, y: 0.65, w: 12.6, h: 6.3,
    fontFace: 'Nirmala UI', fontSize: 9,
    color: '111111', margin: 0.05,
    fit: 'shrink', breakLine: false,
    fill: { color: 'FFFFFF', transparency: 6 },
    line: { color: 'BFD0E2', pt: 0.5 }
  });
}

function clamp(v, min, max) { return Math.max(min, Math.min(max, v)); }
function isLikelyHeader(t) { return /^(rank|score|janpad|engineer|सम|श्रेणी|जल|srdm|#|work type)/i.test(String(t || '').trim()); }
function guessAlign(text, x, w) {
  const t = String(text || '').trim();
  if (/^[\d.,%₹\-+/]+$/.test(t)) return 'center';
  if (x + w / 2 > 4.5 && x + w / 2 < 8.8 && t.length > 14) return 'center';
  return 'left';
}

async function makeTestPpt() {
  try {
    requireLibraries();
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    const slide = pptx.addSlide();
    slide.background = { color: 'F4F8FF' };
    slide.addShape(pptx.ShapeType?.rect || 'rect', { x: 0, y: 0, w: 13.333, h: 0.6, fill: { color: '0B3D73' }, line: { color: '0B3D73' } });
    slide.addText('PPT Export Test - Editable Text', { x: 0.4, y: 0.12, w: 12.5, h: 0.35, color: 'FFFFFF', bold: true, align: 'center', fontSize: 18, fontFace: 'Nirmala UI' });
    slide.addTable([
      [{ text: 'Column A', options: { bold: true, fill: { color: '0B3D73' }, color: 'FFFFFF' } }, { text: 'Column B', options: { bold: true, fill: { color: '0B3D73' }, color: 'FFFFFF' } }],
      ['Editable cell 1', 'Editable cell 2'],
      ['Hindi test', 'संपादन योग्य टेक्स्ट']
    ], { x: 1, y: 1.3, w: 11.3, h: 2.0, border: { type: 'solid', color: 'AFC3D7', pt: 0.75 }, fontFace: 'Nirmala UI', fontSize: 14, align: 'center', valign: 'mid' });
    await pptx.writeFile({ fileName: 'ppt_export_test_editable.pptx' });
    setStatus(100, 'Test PPT downloaded. Agar ye open ho rahi hai, PPT library sahi hai.');
  } catch (err) {
    console.error(err);
    setStatus(100, 'Test PPT error: ' + (err.message || err));
  }
}
