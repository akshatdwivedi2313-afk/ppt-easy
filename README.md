# SRDM PPT Image to Editable Excel/PPT Converter

Static GitHub Pages compatible browser tool.

## What this tool does

- Upload image-only `.pptx`
- Extract slide images from the PPTX in the browser
- Run OCR with Tesseract.js
- Export editable `.xlsx`
- Export editable `.pptx`
- No backend/server required

## Best PPT mode

For PowerPoint output, use:

`PPT Rebuild Style → Same-layout editable text boxes (Recommended)`

This creates editable text boxes near the original OCR positions. The table mode is experimental because browser OCR cannot always detect exact table cell boundaries.

## Important notes

- Internet is required because GitHub Pages loads JSZip, Tesseract.js, XLSX and PptxGenJS from CDN.
- If `pptxgen is not defined` appears, use this fixed version. It now loads PptxGenJS dynamically from fallback CDNs and detects multiple global names.
- Hindi OCR can be slow and imperfect. Use high-resolution PPT images for better output.

## Deploy on GitHub Pages

1. Upload all files to a GitHub repository.
2. Go to `Settings → Pages`.
3. Select `Deploy from branch`.
4. Branch: `main`, Folder: `/root`.
5. Save.

Your site will be available at:

`https://USERNAME.github.io/REPOSITORY_NAME/`

## Files

- `index.html`
- `style.css`
- `app.js`
- `README.md`
