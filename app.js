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

  // Editable PPT generation is now done with our own OOXML writer using JSZip.
  // This removes the earlier "pptxgen is not defined" dependency problem.
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
  const pptStyle = document.getElementById("pptStyle").value;
  const blueTheme = document.getElementById("blueTheme").checked;
  const addReferenceImage = document.getElementById("addReferenceImage")?.checked;

  const blob = await buildEditablePptxWithJSZip(results, {
    baseName,
    pptStyle,
    blueTheme,
    addReferenceImage
  });

  downloadBlob(blob, `${baseName}_editable.pptx`);
  addDownloadMessage("Editable PPT generated without PptxGenJS dependency.");
}

async function buildEditablePptxWithJSZip(results, options) {
  const zip = new JSZip();
  const slides = results.length ? results : [];
  const imageRels = [];

  zip.file("[Content_Types].xml", buildContentTypesXml(slides.length));
  zip.file("_rels/.rels", rootRelsXml());
  zip.file("docProps/core.xml", corePropsXml(options.baseName || "Editable Converted PPT"));
  zip.file("docProps/app.xml", appPropsXml(slides.length));
  zip.file("ppt/presentation.xml", presentationXml(slides.length));
  zip.file("ppt/_rels/presentation.xml.rels", presentationRelsXml(slides.length));
  zip.file("ppt/theme/theme1.xml", themeXml());
  zip.file("ppt/slideMasters/slideMaster1.xml", slideMasterXml());
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMasterRelsXml());
  zip.file("ppt/slideLayouts/slideLayout1.xml", slideLayoutXml());
  zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayoutRelsXml());

  for (let i = 0; i < slides.length; i++) {
    const slideData = slides[i];
    const slideNo = i + 1;
    let imageInfo = null;

    if (options.addReferenceImage && slideData.dataUrl) {
      imageInfo = dataUrlToBinary(slideData.dataUrl, slideNo);
      zip.file(`ppt/media/${imageInfo.fileName}`, imageInfo.bytes);
    }

    zip.file(`ppt/slides/slide${slideNo}.xml`, slideXml(slideData, options, imageInfo));
    zip.file(`ppt/slides/_rels/slide${slideNo}.xml.rels`, slideRelsXml(imageInfo));
  }

  return await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
}

const EMU_PER_INCH = 914400;
const PPT_W_IN = 13.333333;
const PPT_H_IN = 7.5;
const PPT_W_EMU = 12192000;
const PPT_H_EMU = 6858000;

function inchToEmu(v) {
  return Math.round(v * EMU_PER_INCH);
}

function ptToSz(pt) {
  return Math.round(pt * 100);
}

function xmlEscape(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function buildContentTypesXml(slideCount) {
  const slideOverrides = Array.from({ length: slideCount }, (_, i) =>
    `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
  ).join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  ${slideOverrides}
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
}

function rootRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
}

function corePropsXml(title) {
  const safeTitle = xmlEscape(title || "Editable Converted PPT");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${safeTitle}</dc:title>
  <dc:creator>SRDM PPT Image Converter</dc:creator>
  <cp:lastModifiedBy>SRDM PPT Image Converter</cp:lastModifiedBy>
</cp:coreProperties>`;
}

function appPropsXml(slideCount) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>SRDM PPT Image Converter</Application>
  <PresentationFormat>On-screen Show (16:9)</PresentationFormat>
  <Slides>${slideCount}</Slides>
</Properties>`;
}

function presentationXml(slideCount) {
  const ids = Array.from({ length: slideCount }, (_, i) =>
    `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`
  ).join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>${ids}</p:sldIdLst>
  <p:sldSz cx="${PPT_W_EMU}" cy="${PPT_H_EMU}" type="wide"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></p:defaultTextStyle>
</p:presentation>`;
}

function presentationRelsXml(slideCount) {
  const slideRels = Array.from({ length: slideCount }, (_, i) =>
    `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`
  ).join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  ${slideRels}
</Relationships>`;
}

function slideMasterXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>`;
}

function slideMasterRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;
}

function slideLayoutXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`;
}

function slideLayoutRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;
}

function themeXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="SRDM">
  <a:themeElements>
    <a:clrScheme name="SRDM"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F4E79"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="0B3D73"/></a:accent1><a:accent2><a:srgbClr val="27A7DF"/></a:accent2><a:accent3><a:srgbClr val="70AD47"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="ED7D31"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
    <a:fontScheme name="SRDM"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface="Nirmala UI"/><a:cs typeface="Nirmala UI"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface="Nirmala UI"/><a:cs typeface="Nirmala UI"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="SRDM"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
  </a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/>
</a:theme>`;
}

function slideXml(slideData, options, imageInfo) {
  let shapeId = 2;
  const shapes = [];
  const bgColor = options.blueTheme ? "F4F8FF" : "FFFFFF";

  if (options.addReferenceImage && imageInfo) {
    shapes.push(pictureXml(shapeId++, "rId2", imageInfo.fileName));
  }

  if (options.blueTheme) {
    shapes.push(rectXml(shapeId++, 0, 0, PPT_W_EMU, inchToEmu(0.48), "0B3D73", "0B3D73"));
  }

  if (options.pptStyle === "layout") {
    shapes.push(...sameLayoutTextXml(slideData, options, shapeId));
    shapeId += 10000;
  } else if (options.pptStyle === "text") {
    shapes.push(textBoxXml(shapeId++, slideData.text || "No text detected", inchToEmu(0.45), inchToEmu(0.75), inchToEmu(12.45), inchToEmu(6.25), {
      fontSize: 9,
      color: "111111",
      align: "l",
      fill: "FFFFFF",
      line: "C9D6E8",
      margin: 0.05
    }));
  } else {
    const tableShapes = tableTextXml(slideData, options, shapeId);
    shapes.push(...tableShapes);
    shapeId += tableShapes.length;
  }

  shapes.push(textBoxXml(shapeId++, "Converted to editable format • SRDM ZP SATNA", inchToEmu(0.25), inchToEmu(7.18), inchToEmu(12.85), inchToEmu(0.18), {
    fontSize: 7,
    color: "666666",
    align: "ctr",
    margin: 0
  }));

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="${bgColor}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      ${shapes.join("\n")}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
}

function slideRelsXml(imageInfo) {
  const imageRel = imageInfo
    ? `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${xmlEscape(imageInfo.fileName)}"/>`
    : "";

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  ${imageRel}
</Relationships>`;
}

function sameLayoutTextXml(slideData, options, startId) {
  const shapes = [];
  const imgW = Math.max(1, slideData.imageWidth || 1600);
  const imgH = Math.max(1, slideData.imageHeight || 900);
  const lines = slideData.lineBoxes && slideData.lineBoxes.length ? slideData.lineBoxes : groupWordsIntoLineBoxes(slideData.words || []);

  if (!lines.length) {
    return [textBoxXml(startId, slideData.text || "No text detected", inchToEmu(0.5), inchToEmu(1), inchToEmu(12.2), inchToEmu(5.8), {
      fontSize: 12,
      color: "333333",
      align: "l"
    })];
  }

  let id = startId;
  for (const line of lines) {
    const rawX = (line.x0 / imgW) * PPT_W_IN;
    const rawY = (line.y0 / imgH) * PPT_H_IN;
    const rawW = ((line.x1 - line.x0) / imgW) * PPT_W_IN;
    const rawH = ((line.y1 - line.y0) / imgH) * PPT_H_IN;

    const x = clamp(rawX - 0.02, 0.02, PPT_W_IN - 0.1);
    const y = clamp(rawY - 0.01, 0.02, PPT_H_IN - 0.12);
    const w = clamp(rawW + 0.15, 0.18, PPT_W_IN - x - 0.03);
    const h = clamp(rawH + 0.06, 0.10, 0.45);
    const estimatedFont = clamp(rawH * 72 * 0.82, 5, 16);
    const isTopLine = y < 0.55;

    shapes.push(textBoxXml(id++, line.text, inchToEmu(x), inchToEmu(y), inchToEmu(w), inchToEmu(h), {
      fontSize: estimatedFont,
      bold: isLikelyHeader(line.text, isTopLine),
      color: isTopLine && options.blueTheme ? "FFFFFF" : "111111",
      align: alignToOOXML(guessAlign(line.text, x, w, PPT_W_IN)),
      margin: 0
    }));
  }

  return shapes;
}

function tableTextXml(slideData, options, startId) {
  const shapes = [];
  let id = startId;

  shapes.push(textBoxXml(id++, slideData.title || `Slide ${slideData.slideNo}`, inchToEmu(0.25), inchToEmu(0.08), inchToEmu(12.85), inchToEmu(0.35), {
    fontSize: 14,
    bold: true,
    color: options.blueTheme ? "FFFFFF" : "0B3D73",
    align: "ctr",
    margin: 0.02
  }));

  const rows = normalizeRows(slideData.rows || []).slice(0, 40);
  const colCount = Math.max(1, ...rows.map(r => r.length), 1);

  if (!rows.length) {
    shapes.push(textBoxXml(id++, "No text detected", inchToEmu(0.5), inchToEmu(1), inchToEmu(12), inchToEmu(1), { fontSize: 18, color: "666666" }));
    return shapes;
  }

  const x0 = 0.25;
  const y0 = 0.78;
  const tableW = 12.85;
  const tableH = 6.18;
  const rowH = tableH / rows.length;
  const colW = tableW / colCount;
  const fontSize = colCount > 9 ? 5.2 : colCount > 6 ? 6.2 : 7.2;

  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < colCount; c++) {
      const cellText = String(rows[r][c] || "");
      const fill = r === 0 && options.blueTheme ? "0B3D73" : "FFFFFF";
      const color = r === 0 && options.blueTheme ? "FFFFFF" : "111111";
      shapes.push(rectXml(id++, inchToEmu(x0 + c * colW), inchToEmu(y0 + r * rowH), inchToEmu(colW), inchToEmu(rowH), fill, "8EA9C9"));
      shapes.push(textBoxXml(id++, cellText, inchToEmu(x0 + c * colW + 0.02), inchToEmu(y0 + r * rowH + 0.01), inchToEmu(colW - 0.04), inchToEmu(rowH - 0.02), {
        fontSize,
        bold: r === 0,
        color,
        align: "ctr",
        margin: 0
      }));
    }
  }

  return shapes;
}

function rectXml(id, x, y, w, h, fill, line) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Rectangle ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="${line || fill}"/></a:solidFill></a:ln></p:spPr></p:sp>`;
}

function textBoxXml(id, text, x, y, w, h, opts = {}) {
  const fontSize = ptToSz(opts.fontSize || 10);
  const color = opts.color || "111111";
  const bold = opts.bold ? ' b="1"' : "";
  const align = opts.align || "l";
  const margin = typeof opts.margin === "number" ? inchToEmu(opts.margin) : 0;
  const fillXml = opts.fill ? `<a:solidFill><a:srgbClr val="${opts.fill}"/></a:solidFill>` : `<a:noFill/>`;
  const lineXml = opts.line ? `<a:ln><a:solidFill><a:srgbClr val="${opts.line}"/></a:solidFill></a:ln>` : `<a:ln><a:noFill/></a:ln>`;

  const lines = String(text ?? "").split(/\r?\n/);
  const paragraphs = lines.map(line =>
    `<a:p><a:pPr algn="${align}"/><a:r><a:rPr lang="en-US" altLang="hi-IN" sz="${fontSize}"${bold}><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="Arial"/><a:ea typeface="Nirmala UI"/><a:cs typeface="Nirmala UI"/></a:rPr><a:t>${xmlEscape(line)}</a:t></a:r><a:endParaRPr lang="en-US" sz="${fontSize}"/></a:p>`
  ).join("");

  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fillXml}${lineXml}</p:spPr><p:txBody><a:bodyPr wrap="square" rtlCol="0" lIns="${margin}" tIns="${margin}" rIns="${margin}" bIns="${margin}"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
}

function pictureXml(id, rid, fileName) {
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${xmlEscape(fileName)}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"><a:alphaModFix amt="18000"/></a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${PPT_W_EMU}" cy="${PPT_H_EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
}

function alignToOOXML(value) {
  if (value === "center") return "ctr";
  if (value === "right") return "r";
  return value || "l";
}

function dataUrlToBinary(dataUrl, slideNo) {
  const [header, base64] = String(dataUrl).split(",");
  const mimeMatch = header.match(/data:([^;]+)/i);
  const mime = mimeMatch ? mimeMatch[1].toLowerCase() : "image/png";
  const ext = mime.includes("jpeg") || mime.includes("jpg") ? "jpg" : "png";
  const binary = atob(base64 || "");
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return { bytes, fileName: `reference_slide_${slideNo}.${ext}`, mime, ext };
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
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
