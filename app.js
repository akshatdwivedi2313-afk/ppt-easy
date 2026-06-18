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

      const result = {
        slideNo: item.slideNo,
        imageName: item.imageName,
        dataUrl: item.dataUrl,
        text: ocr.text,
        words: ocr.words,
        rows,
        title
      };
      slideResults.push(result);
      addPreviewCard(result);
    }

    const outputType = document.getElementById("outputType").value;

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
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "SRDM PPT Image Converter";
  pptx.subject = "Editable OCR converted presentation";
  pptx.title = `${baseName} editable converted`;
  pptx.company = "SRDM ZP SATNA";

  const pptStyle = document.getElementById("pptStyle").value;
  const blueTheme = document.getElementById("blueTheme").checked;

  for (const slideData of results) {
    const slide = pptx.addSlide();
    if (blueTheme) {
      slide.background = { color: "F4F8FF" };
      slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.55, fill: { color: "0B3D73" }, line: { color: "0B3D73" } });
    }

    slide.addText(slideData.title || `Slide ${slideData.slideNo}`, {
      x: 0.25, y: 0.08, w: 12.85, h: 0.35,
      fontFace: "Arial",
      fontSize: 14,
      bold: true,
      color: blueTheme ? "FFFFFF" : "0B3D73",
      align: "center",
      margin: 0.02
    });

    slide.addText("Converted to editable format • SRDM ZP SATNA", {
      x: 0.25, y: 7.12, w: 12.85, h: 0.22,
      fontSize: 8,
      color: "666666",
      align: "center",
      margin: 0
    });

    if (pptStyle === "text") {
      slide.addText(slideData.text || "", {
        x: 0.45, y: 0.78, w: 12.45, h: 6.15,
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
  }

  await pptx.writeFile({ fileName: `${baseName}_editable.pptx` });
  addDownloadMessage("Editable PPT file generated by browser download.");
}

function addDownloadMessage(text) {
  const div = document.createElement("div");
  div.className = "download-link";
  div.textContent = text;
  downloadsEl.appendChild(div);
}
