/*
  SRDM PPT Image to Editable Excel/PPT Converter
  Static GitHub Pages compatible version.
  It reads image-only PPTX files in the browser, runs OCR, and exports editable XLSX/PPTX.
*/

const pptFileEl = document.getElementById("pptFile");
const fileNameEl = document.getElementById("fileName");
const convertBtn = document.getElementById("convertBtn");
const resetBtn = document.getElementById("resetBtn");
const statusEl = document.getElementById("status");
const progressWrap = document.getElementById("progressWrap");
const progressBar = document.getElementById("progressBar");
const downloadsEl = document.getElementById("downloads");
const previewEl = document.getElementById("preview");

let slideResults = [];

pptFileEl.addEventListener("change", () => {
  const file = pptFileEl.files?.[0];
  fileNameEl.textContent = file ? file.name : "PPTX file choose करें";
});

convertBtn.addEventListener("click", convertPPT);
resetBtn.addEventListener("click", resetTool);

function resetTool() {
  pptFileEl.value = "";
  fileNameEl.textContent = "PPTX file choose करें";
  downloadsEl.innerHTML = "";
  previewEl.className = "preview-empty";
  previewEl.textContent = "PPT upload करके Convert Now दबाएँ.";
  slideResults = [];
  progressWrap.classList.add("hidden");
  setProgress(0, "Ready");
}

function setProgress(percent, message) {
  progressWrap.classList.remove("hidden");
  progressBar.style.width = `${Math.max(0, Math.min(100, percent))}%`;
  statusEl.textContent = message;
}


function loadExternalScript(url) {
  return new Promise((resolve, reject) => {
    const existing = [...document.scripts].find(s => s.src === url);
    if (existing) {
      if (existing.dataset.loaded === "true") return resolve();
      existing.addEventListener("load", resolve, { once: true });
      existing.addEventListener("error", () => reject(new Error(`Failed loading ${url}`)), { once: true });
      return;
    }
    const script = document.createElement("script");
    script.src = url;
    script.async = true;
    script.crossOrigin = "anonymous";
    script.dataset.dynamicLibrary = "true";
    script.onload = () => { script.dataset.loaded = "true"; resolve(); };
    script.onerror = () => reject(new Error(`Failed loading ${url}`));
    document.head.appendChild(script);
  });
}

async function ensureLibrary(name, checker, urls) {
  if (checker()) return checker();
  let lastError = null;
  for (const url of urls) {
    try {
      await loadExternalScript(url);
      if (checker()) return checker();
    } catch (err) {
      lastError = err;
      console.warn(err);
    }
  }
  throw new Error(`${name} library load नहीं हुई. Internet/CDN/adblock check करें. ${lastError ? lastError.message : ""}`);
}

function resolvePptxGenConstructor() {
  const candidates = [
    window.pptxgen,
    window.pptxgenjs,
    window.PptxGenJS,
    window.PPTXGenJS,
    window.PptxGen,
    window.PPTXGen
  ];
  for (const item of candidates) {
    if (typeof item === "function") return item;
    if (item && typeof item.default === "function") return item.default;
  }
  return null;
}

async function ensureCoreLibraries(outputType) {
  await ensureLibrary("JSZip", () => window.JSZip, [
    "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js",
    "https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js",
    "https://unpkg.com/jszip@3.10.1/dist/jszip.min.js"
  ]);

  await ensureLibrary("Tesseract.js", () => window.Tesseract, [
    "https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js",
    "https://unpkg.com/tesseract.js@5/dist/tesseract.min.js"
  ]);

  if (outputType === "excel" || outputType === "both") {
    await ensureLibrary("XLSX", () => window.XLSX, [
      "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
      "https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js",
      "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"
    ]);
  }

  if (outputType === "ppt" || outputType === "both") {
    await ensureLibrary("PptxGenJS", resolvePptxGenConstructor, [
      "https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js",
      "https://unpkg.com/pptxgenjs@3.12.0/dist/pptxgen.bundle.js",
      "https://cdn.jsdelivr.net/npm/pptxgenjs@3/dist/pptxgen.bundle.js"
    ]);
  }
}


async function convertPPT() {
  const file = pptFileEl.files?.[0];
  if (!file) {
    alert("Please upload a PPTX file first.");
    return;
  }

  convertBtn.disabled = true;
  downloadsEl.innerHTML = "";
  previewEl.className = "";
  previewEl.innerHTML = "";
  slideResults = [];

  try {
    const outputType = document.getElementById("outputType").value;
    setProgress(2, "Loading browser libraries...");
    await ensureCoreLibraries(outputType);

    setProgress(3, "Reading PPTX file...");
    const zip = await JSZip.loadAsync(await file.arrayBuffer());

    setProgress(8, "Finding slide images...");
    const slideImages = await extractSlideImages(zip);

    if (!slideImages.length) {
      throw new Error("No images found inside PPTX. This tool is mainly for image-only PPT slides.");
    }

    const lang = document.getElementById("ocrLang").value;
    const parseMode = document.getElementById("parseMode").value;

    for (let i = 0; i < slideImages.length; i++) {
      const item = slideImages[i];
      const percentBase = 10 + Math.round((i / slideImages.length) * 72);
      setProgress(percentBase, `OCR reading slide ${i + 1} of ${slideImages.length}...`);

      const ocr = await runOCR(item.dataUrl, lang, (m) => {
        if (m.status === "recognizing text" && typeof m.progress === "number") {
          const p = percentBase + Math.round(m.progress * (72 / slideImages.length));
          setProgress(p, `OCR reading slide ${i + 1} of ${slideImages.length}... ${Math.round(m.progress * 100)}%`);
        }
      });

      const rows = parseOCRToRows(ocr, parseMode);
      const title = detectTitle(rows, item.slideNo);
      const imageSize = await getImageSize(item.dataUrl);
      const lineBoxes = groupWordsIntoLineBoxes(ocr.words);

      const result = {
        slideNo: item.slideNo,
        imageName: item.imageName,
        dataUrl: item.dataUrl,
        imageWidth: imageSize.width,
        imageHeight: imageSize.height,
        text: ocr.text,
        words: ocr.words,
        lineBoxes,
        rows,
        title
      };
      slideResults.push(result);
      addPreviewCard(result);
    }

    if (outputType === "excel" || outputType === "both") {
      setProgress(88, "Creating editable Excel...");
      createExcel(slideResults, cleanBaseName(file.name));
    }

    if (outputType === "ppt" || outputType === "both") {
      setProgress(95, "Creating editable PPT...");
      await createEditablePPT(slideResults, cleanBaseName(file.name));
    }

    setProgress(100, "Done. Download buttons are ready.");
  } catch (err) {
    console.error(err);
    setProgress(100, `Error: ${err.message || err}`);
  } finally {
    convertBtn.disabled = false;
  }
}

function cleanBaseName(name) {
  return name.replace(/\.pptx$/i, "").replace(/[^a-zA-Z0-9_-]+/g, "_").slice(0, 80) || "converted";
}

async function extractSlideImages(zip) {
  const slidePaths = Object.keys(zip.files)
    .filter(p => /^ppt\/slides\/slide\d+\.xml$/i.test(p))
    .sort((a, b) => slideNumber(a) - slideNumber(b));

  const results = [];

  for (const slidePath of slidePaths) {
    const no = slideNumber(slidePath);
    const relPath = `ppt/slides/_rels/slide${no}.xml.rels`;
    const slideXml = await zip.file(slidePath)?.async("text");
    const relXml = await zip.file(relPath)?.async("text");

    if (!slideXml || !relXml) continue;

    const relMap = parseRelationshipMap(relXml);
    const imageRids = extractImageRidsWithArea(slideXml);
    const candidates = [];

    for (const img of imageRids) {
      const target = relMap[img.rid];
      if (!target) continue;
      const normalized = normalizeTargetPath("ppt/slides", target);
      if (!/^ppt\/media\//i.test(normalized)) continue;
      candidates.push({ ...img, path: normalized });
    }

    // Image-only PPTs usually have one large image. Select the largest referenced image per slide.
    const unique = dedupeByPath(candidates);
    unique.sort((a, b) => (b.area || 0) - (a.area || 0));

    const chosen = unique[0];
    if (chosen && zip.file(chosen.path)) {
      const blob = await zip.file(chosen.path).async("blob");
      const dataUrl = await blobToDataURL(blob);
      results.push({ slideNo: no, imageName: chosen.path.split("/").pop(), dataUrl });
    }
  }

  // Fallback: if slide relationship parsing fails, read all media images.
  if (!results.length) {
    const mediaPaths = Object.keys(zip.files)
      .filter(p => /^ppt\/media\/.*\.(png|jpg|jpeg)$/i.test(p))
      .sort(naturalSort);

    for (let i = 0; i < mediaPaths.length; i++) {
      const blob = await zip.file(mediaPaths[i]).async("blob");
      results.push({ slideNo: i + 1, imageName: mediaPaths[i].split("/").pop(), dataUrl: await blobToDataURL(blob) });
    }
  }

  return results;
}

function slideNumber(path) {
  const m = path.match(/slide(\d+)\.xml/i);
  return m ? Number(m[1]) : 0;
}

function naturalSort(a, b) {
  return a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
}

function parseRelationshipMap(xml) {
  const map = {};
  const re = /<Relationship\s+[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*>/gi;
  let m;
  while ((m = re.exec(xml)) !== null) {
    map[m[1]] = m[2];
  }
  return map;
}

function extractImageRidsWithArea(xml) {
  const items = [];
  const blipRe = /<a:blip[^>]*(?:r:embed|r:link)="([^"]+)"[^>]*>/gi;
  let m;
  while ((m = blipRe.exec(xml)) !== null) {
    const rid = m[1];
    const around = xml.slice(Math.max(0, m.index - 1500), Math.min(xml.length, m.index + 2500));
    const ext = around.match(/<a:ext[^>]*cx="(\d+)"[^>]*cy="(\d+)"/i);
    const area = ext ? Number(ext[1]) * Number(ext[2]) : 0;
    items.push({ rid, area });
  }
  return items;
}

function normalizeTargetPath(baseDir, target) {
  if (target.startsWith("/")) return target.slice(1);
  const parts = `${baseDir}/${target}`.split("/");
  const stack = [];
  for (const part of parts) {
    if (!part || part === ".") continue;
    if (part === "..") stack.pop();
    else stack.push(part);
  }
  return stack.join("/");
}

function dedupeByPath(items) {
  const map = new Map();
  for (const item of items) {
    if (!map.has(item.path) || (item.area || 0) > (map.get(item.path).area || 0)) {
      map.set(item.path, item);
    }
  }
  return [...map.values()];
}

function blobToDataURL(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}


function getImageSize(dataUrl) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth || img.width || 1600, height: img.naturalHeight || img.height || 900 });
    img.onerror = () => resolve({ width: 1600, height: 900 });
    img.src = dataUrl;
  });
}

async function runOCR(dataUrl, lang, logger) {
  const result = await Tesseract.recognize(dataUrl, lang, { logger });
  return {
    text: result?.data?.text || "",
    words: normalizeWords(result?.data?.words || [])
  };
}

function normalizeWords(words) {
  return words
    .map(w => {
      const bbox = w.bbox || {};
      return {
        text: (w.text || "").trim(),
        confidence: w.confidence || 0,
        x0: bbox.x0 ?? bbox.left ?? 0,
        y0: bbox.y0 ?? bbox.top ?? 0,
        x1: bbox.x1 ?? bbox.right ?? 0,
        y1: bbox.y1 ?? bbox.bottom ?? 0
      };
    })
    .filter(w => w.text);
}



function groupWordsIntoLineBoxes(words) {
  const sorted = [...words]
    .filter(w => w.text && (w.x1 - w.x0) > 1 && (w.y1 - w.y0) > 1)
    .sort((a, b) => (a.y0 - b.y0) || (a.x0 - b.x0));

  const heights = sorted.map(w => Math.max(8, w.y1 - w.y0));
  const medianHeight = median(heights) || 14;
  const yTol = Math.max(9, medianHeight * 0.72);
  const groups = [];

  for (const word of sorted) {
    const midY = (word.y0 + word.y1) / 2;
    let best = null;
    let bestDiff = Infinity;

    for (const group of groups) {
      const diff = Math.abs(group.midY - midY);
      if (diff < yTol && diff < bestDiff) {
        best = group;
        bestDiff = diff;
      }
    }

    if (best) {
      best.words.push(word);
      best.midY = (best.midY * (best.words.length - 1) + midY) / best.words.length;
    } else {
      groups.push({ midY, words: [word] });
    }
  }

  return groups
    .map(group => {
      const rowWords = group.words.sort((a, b) => a.x0 - b.x0);
      return {
        text: rowWords.map(w => w.text).join(" ").replace(/\s+/g, " ").trim(),
        x0: Math.min(...rowWords.map(w => w.x0)),
        y0: Math.min(...rowWords.map(w => w.y0)),
        x1: Math.max(...rowWords.map(w => w.x1)),
        y1: Math.max(...rowWords.map(w => w.y1)),
        confidence: rowWords.reduce((sum, w) => sum + (w.confidence || 0), 0) / Math.max(rowWords.length, 1)
      };
    })
    .filter(line => line.text)
    .sort((a, b) => (a.y0 - b.y0) || (a.x0 - b.x0));
}

function parseOCRToRows(ocr, mode) {
  if (mode === "plain") {
    return [[ocr.text.trim()]];
  }

  if (mode === "lines" || !ocr.words.length) {
    return ocr.text
      .split(/\r?\n/)
      .map(x => x.trim())
      .filter(Boolean)
      .map(line => [line]);
  }

  return groupWordsAsTable(ocr.words);
}

function groupWordsAsTable(words) {
  const sorted = [...words].sort((a, b) => (a.y0 - b.y0) || (a.x0 - b.x0));
  const heights = sorted.map(w => Math.max(8, w.y1 - w.y0));
  const medianHeight = median(heights) || 14;
  const yTol = Math.max(10, medianHeight * 0.65);

  const lineGroups = [];
  for (const word of sorted) {
    const midY = (word.y0 + word.y1) / 2;
    let best = null;
    let bestDiff = Infinity;

    for (const group of lineGroups) {
      const diff = Math.abs(group.midY - midY);
      if (diff < yTol && diff < bestDiff) {
        best = group;
        bestDiff = diff;
      }
    }

    if (best) {
      best.words.push(word);
      best.midY = (best.midY * (best.words.length - 1) + midY) / best.words.length;
    } else {
      lineGroups.push({ midY, words: [word] });
    }
  }

  lineGroups.sort((a, b) => a.midY - b.midY);

  const rows = [];
  for (const group of lineGroups) {
    const rowWords = group.words.sort((a, b) => a.x0 - b.x0);
    const gaps = [];
    for (let i = 1; i < rowWords.length; i++) {
      gaps.push(Math.max(0, rowWords[i].x0 - rowWords[i - 1].x1));
    }
    const positiveGaps = gaps.filter(g => g > 2);
    const gapMedian = median(positiveGaps) || 18;
    const splitGap = Math.max(25, gapMedian * 2.2);

    const cells = [];
    let current = [];

    for (let i = 0; i < rowWords.length; i++) {
      if (i > 0) {
        const gap = Math.max(0, rowWords[i].x0 - rowWords[i - 1].x1);
        if (gap > splitGap) {
          cells.push(current.map(w => w.text).join(" "));
          current = [];
        }
      }
      current.push(rowWords[i]);
    }
    if (current.length) cells.push(current.map(w => w.text).join(" "));

    if (cells.some(c => c.trim())) rows.push(cells.map(cleanCell));
  }

  return rows;
}

function cleanCell(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .replace(/₹\s+/g, "₹")
    .trim();
}

function median(arr) {
  if (!arr.length) return 0;
  const sorted = [...arr].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

function detectTitle(rows, slideNo) {
  const firstGood = rows.flat().find(x => String(x).trim().length > 4);
  return firstGood ? String(firstGood).slice(0, 90) : `Slide ${slideNo}`;
}

function addPreviewCard(result) {
  const card = document.createElement("div");
  card.className = "slide-card";
  card.innerHTML = `
    <h4>Slide ${result.slideNo}</h4>
    <img class="slide-img" src="${result.dataUrl}" alt="Slide ${result.slideNo}" />
    <pre class="ocr-text"></pre>
  `;
  card.querySelector("pre").textContent = result.text || "No text detected";
  previewEl.appendChild(card);
}

function createExcel(results, baseName) {
  const wb = XLSX.utils.book_new();

  results.forEach((slide) => {
    const data = normalizeRows(slide.rows);
    const ws = XLSX.utils.aoa_to_sheet(data.length ? data : [[slide.text || ""]]);
    const maxCols = Math.max(...(data.map(r => r.length)), 1);
    ws["!cols"] = Array.from({ length: maxCols }, () => ({ wch: 20 }));
    XLSX.utils.book_append_sheet(wb, ws, safeSheetName(`Slide ${slide.slideNo}`));
  });

  XLSX.writeFile(wb, `${baseName}_editable.xlsx`);
  addDownloadMessage("Excel file generated by browser download.");
}

function safeSheetName(name) {
  return name.replace(/[\\/*?:\[\]]/g, " ").slice(0, 31);
}

function normalizeRows(rows) {
  const maxCols = Math.max(1, ...rows.map(r => r.length));
  return rows.map(row => {
    const out = [...row];
    while (out.length < maxCols) out.push("");
    return out;
  });
}

async function createEditablePPT(results, baseName) {
  const PptxGen = await ensureLibrary("PptxGenJS", resolvePptxGenConstructor, [
    "https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js",
    "https://unpkg.com/pptxgenjs@3.12.0/dist/pptxgen.bundle.js",
    "https://cdn.jsdelivr.net/npm/pptxgenjs@3/dist/pptxgen.bundle.js"
  ]);

  const pptx = new PptxGen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "SRDM PPT Image Converter";
  pptx.subject = "Editable OCR converted presentation";
  pptx.title = `${baseName} editable converted`;
  pptx.company = "SRDM ZP SATNA";

  const pptStyle = document.getElementById("pptStyle").value;
  const blueTheme = document.getElementById("blueTheme").checked;
  const addReferenceImage = document.getElementById("addReferenceImage")?.checked;
  const shapeTypes = PptxGen.ShapeType || pptx.ShapeType || window.pptxgen?.ShapeType || { rect: "rect" };

  for (const slideData of results) {
    const slide = pptx.addSlide();
    if (blueTheme) {
      slide.background = { color: "F4F8FF" };
      slide.addShape(shapeTypes.rect || "rect", {
        x: 0, y: 0, w: 13.333, h: 0.48,
        fill: { color: "0B3D73" },
        line: { color: "0B3D73" }
      });
    }

    if (addReferenceImage) {
      try {
        slide.addImage({ data: slideData.dataUrl, x: 0, y: 0, w: 13.333, h: 7.5, transparency: 82 });
      } catch (e) {
        console.warn("Reference image could not be added", e);
      }
    }

    if (pptStyle === "layout") {
      addSameLayoutText(slide, slideData, blueTheme);
    } else if (pptStyle === "text") {
      slide.addText(slideData.text || "", {
        x: 0.45, y: 0.75, w: 12.45, h: 6.25,
        fontFace: "Arial",
        fontSize: 9,
        color: "111111",
        valign: "top",
        fit: "shrink",
        margin: 0.08,
        breakLine: false,
        fill: { color: "FFFFFF", transparency: 4 },
        line: { color: "C9D6E8", pt: 1 }
      });
    } else {
      addTableSlide(slide, slideData, blueTheme);
    }

    slide.addText("Converted to editable format • SRDM ZP SATNA", {
      x: 0.25, y: 7.18, w: 12.85, h: 0.18,
      fontSize: 7,
      color: "666666",
      align: "center",
      margin: 0
    });
  }

  await pptx.writeFile({ fileName: `${baseName}_editable.pptx` });
  addDownloadMessage("Editable PPT file generated by browser download.");
}

function addSameLayoutText(slide, slideData, blueTheme) {
  const slideW = 13.333;
  const slideH = 7.5;
  const imgW = Math.max(1, slideData.imageWidth || 1600);
  const imgH = Math.max(1, slideData.imageHeight || 900);
  const lines = slideData.lineBoxes && slideData.lineBoxes.length ? slideData.lineBoxes : groupWordsIntoLineBoxes(slideData.words || []);

  if (!lines.length) {
    slide.addText(slideData.text || "No text detected", {
      x: 0.5, y: 1, w: 12.2, h: 5.8,
      fontSize: 12,
      color: "333333",
      fit: "shrink",
      valign: "top"
    });
    return;
  }

  for (const line of lines) {
    const rawX = (line.x0 / imgW) * slideW;
    const rawY = (line.y0 / imgH) * slideH;
    const rawW = ((line.x1 - line.x0) / imgW) * slideW;
    const rawH = ((line.y1 - line.y0) / imgH) * slideH;

    const x = clamp(rawX - 0.02, 0.02, slideW - 0.1);
    const y = clamp(rawY - 0.01, 0.02, slideH - 0.12);
    const w = clamp(rawW + 0.12, 0.18, slideW - x - 0.03);
    const h = clamp(rawH + 0.05, 0.10, 0.40);
    const estimatedFont = clamp(rawH * 72 * 0.82, 5, 16);

    const isTopLine = y < 0.55;
    slide.addText(line.text, {
      x, y, w, h,
      fontFace: "Arial",
      fontSize: estimatedFont,
      bold: isLikelyHeader(line.text, isTopLine),
      color: isTopLine && blueTheme ? "FFFFFF" : "111111",
      margin: 0,
      breakLine: false,
      fit: "shrink",
      valign: "mid",
      align: guessAlign(line.text, x, w, slideW)
    });
  }
}

function addTableSlide(slide, slideData, blueTheme) {
  slide.addText(slideData.title || `Slide ${slideData.slideNo}`, {
    x: 0.25, y: 0.08, w: 12.85, h: 0.35,
    fontFace: "Arial",
    fontSize: 14,
    bold: true,
    color: blueTheme ? "FFFFFF" : "0B3D73",
    align: "center",
    margin: 0.02
  });

  const rows = normalizeRows(slideData.rows).slice(0, 45);
  const colCount = Math.max(1, ...rows.map(r => r.length));
  const tableRows = rows.map((row, rIdx) => row.map(cell => ({
    text: String(cell || ""),
    options: {
      fontFace: "Arial",
      fontSize: colCount > 8 ? 5.5 : 7,
      bold: rIdx === 0,
      color: rIdx === 0 && blueTheme ? "FFFFFF" : "111111",
      fill: rIdx === 0 && blueTheme ? { color: "0B3D73" } : { color: "FFFFFF" },
      align: "center",
      valign: "mid",
      margin: 0.02,
      breakLine: false
    }
  })));

  if (tableRows.length) {
    slide.addTable(tableRows, {
      x: 0.25, y: 0.78, w: 12.85, h: 6.18,
      border: { type: "solid", color: "8EA9C9", pt: 0.5 },
      margin: 0.02,
      autoFit: false,
      fit: "shrink"
    });
  } else {
    slide.addText("No text detected", { x: 0.5, y: 1, w: 12, h: 1, fontSize: 18, color: "666666" });
  }
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function isLikelyHeader(text, isTopLine) {
  if (isTopLine) return true;
  const s = String(text || "").trim();
  return /^(rank|score|janpad|engineer|समग्र|श्रेणी|जल गंगा|srdm)/i.test(s) || s.length < 22 && /[:：]$/.test(s);
}

function guessAlign(text, x, w, slideW) {
  const t = String(text || "").trim();
  if (/^[\d.,%₹\-+/]+$/.test(t)) return "center";
  const center = x + w / 2;
  if (center > slideW * 0.35 && center < slideW * 0.65 && t.length > 15) return "center";
  return "left";
}

function addDownloadMessage(text) {
  const div = document.createElement("div");
  div.className = "download-link";
  div.textContent = text;
  downloadsEl.appendChild(div);
}
