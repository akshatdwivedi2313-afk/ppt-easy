# PPT Easy Real Converter v5

Ye version **accurate conversion** ke liye hai. GitHub Pages sirf frontend host karta hai, isliye image-only PPT ko actual editable Excel/PPT me convert karne ke liye backend zaroori hai.

## Repo Structure

```text
index.html
style.css
app.js
backend/
  app.py
  requirements.txt
  Dockerfile
  render.yaml
```

## Local Run

```bash
cd backend
python -m venv .venv
.venv\Scripts\activate   # Windows
pip install -r requirements.txt
uvicorn app:app --reload --host 0.0.0.0 --port 8000
```

System me Tesseract install hona chahiye:

- Windows: Tesseract OCR install karo aur PATH me add karo.
- Hindi OCR ke liye Hindi traineddata install ho.

Frontend ke liye root `index.html` open karo aur Backend URL me `http://localhost:8000` rakho.

## Online Deploy

### 1) Backend Render par deploy

Render par New Web Service banao, same GitHub repo connect karo, root directory `backend` set karo.

Docker deploy best hai. Render Dockerfile use karega.

### 2) GitHub Pages frontend

Settings → Pages → Deploy from branch → `main` → `/root`.

Site open hone ke baad backend URL me Render URL paste karo, jaise:

```text
https://ppt-easy-backend.onrender.com
```

## Important

Browser-only OCR exact table nahi bana sakta. v5 backend OpenCV + Tesseract based table detection use karta hai, isliye Excel output v4 se zyada usable hoga.
