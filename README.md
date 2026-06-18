# SRDM PPT Image to Editable Excel/PPT Converter v4

Ye GitHub Pages-compatible browser tool hai. User image-only PPTX upload karta hai, tool slide images extract karta hai, browser OCR chalata hai, aur editable Excel / editable PPT generate karta hai.

## v4 me kya fix hai

- `pptxgen is not defined` fixed: PptxGenJS local `libs/pptxgen.min.js` se load hota hai, CDN dependency nahi.
- Manual OOXML PPT writer hata diya gaya: PPT export ab official PptxGenJS se banta hai, isliye PowerPoint file corrupt nahi honi chahiye.
- Excel single-column mess ko reduce karne ke liye Visual Grid OCR default hai.
- Har slide ka OCR TSV preview/edit box diya hai. Export se pehle cells manually correct kiye ja sakte hain.
- Test PPT Export button added: deployment ke baad pehle isse verify kar sakte ho.

## GitHub Pages deployment

1. Is ZIP ko extract karo.
2. Saari files repo ke root me upload/commit karo.
3. GitHub → Repository → Settings → Pages.
4. Source: Deploy from branch.
5. Branch: `main`, Folder: `/root`.
6. Save.
7. Site open karke hard refresh karo: `Ctrl + F5`.

## Folder structure

```text
index.html
style.css
app.js
libs/
  jszip.min.js
  xlsx.full.min.js
  pptxgen.min.js
  tesseract.min.js
  worker.min.js
  tesseract-core-simd-lstm.wasm.js
```

## Important limitation

GitHub Pages static site hai, isliye OCR browser me hota hai. Low-resolution / blurry / small Hindi table me 100% perfect rows-columns automatic nahi milenge. Best practical output ke liye:

- Excel Build Mode: `Visual Grid OCR`
- PPT Build Mode: `Same-layout editable text boxes`
- OCR Language: `English + Hindi`

Agar production-level perfect table conversion chahiye, toh frontend ke saath backend OCR/AI API add karna padega.
