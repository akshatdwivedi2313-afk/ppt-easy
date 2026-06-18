const $ = (id) => document.getElementById(id);
const backendInput = $('backendUrl');
const saved = localStorage.getItem('pptEasyBackend') || 'https://ppt-easy.onrender.com';
backendInput.value = saved;

$('saveBackend').addEventListener('click', async (e) => {
  e.preventDefault();
  const url = backend();
  if (!url) {
    $('health').textContent = 'Backend URL खाली है.';
    return;
  }
  localStorage.setItem('pptEasyBackend', url);
  $('health').textContent = 'Backend URL saved. Checking connection... wait 30-60 sec on Render free plan.';
  await checkHealth();
});

$('pptFile').addEventListener('change', () => {
  const f=$('pptFile').files[0];
  $('fileName').textContent = f ? f.name : 'PPTX choose karo';
});
$('convertBtn').addEventListener('click', convertNow);

setTimeout(checkHealth, 300);

function setStatus(msg, pct=0){
  $('status').textContent=msg;
  $('bar').style.width=Math.max(0,Math.min(100,pct))+'%';
}
function backend(){
  return (backendInput.value || '').trim().replace(/\/+$/,'');
}

async function fetchWithTimeout(url, options={}, timeoutMs=70000){
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try{
    return await fetch(url, { ...options, signal: controller.signal, cache: 'no-store' });
  } finally {
    clearTimeout(timer);
  }
}

async function checkHealth(){
  const url = backend();
  if(!url){ $('health').textContent = 'Backend URL required.'; return false; }
  $('health').textContent = 'Checking backend... Render free plan first request slow ho sakti hai.';
  try{
    const r = await fetchWithTimeout(url + '/health?ts=' + Date.now(), { method:'GET' }, 70000);
    if(!r.ok){ throw new Error('HTTP ' + r.status + ' ' + r.statusText); }
    const j = await r.json();
    if(j && j.ok){
      $('health').textContent = 'Backend connected ✅';
      return true;
    }
    $('health').textContent = 'Backend response mila, par ok true nahi hai.';
    return false;
  }catch(e){
    console.error('Backend check failed', e);
    $('health').textContent = 'Backend not connected: ' + (e.name === 'AbortError' ? 'timeout' : e.message) + '. Pehle /health URL browser me open karke check karo.';
    return false;
  }
}

async function convertNow(){
  const file = $('pptFile').files[0];
  if(!file){ alert('PPTX choose karo.'); return; }
  if(!backend()){ alert('Backend URL required.'); return; }
  $('convertBtn').disabled=true;
  setStatus('Checking backend...', 5);
  const ok = await checkHealth();
  if(!ok){ $('convertBtn').disabled=false; setStatus('Backend connection failed. /health check karo.',100); return; }
  setStatus('Uploading PPT to backend...', 10);
  try{
    const fd = new FormData();
    fd.append('file', file);
    fd.append('output', $('outputType').value);
    fd.append('lang', $('ocrLang').value);
    fd.append('excel_mode', $('excelMode').value);
    fd.append('ppt_mode', $('pptMode').value);
    fd.append('include_image', $('includeImage').checked ? 'true':'false');
    const r = await fetchWithTimeout(backend() + '/api/convert', { method:'POST', body:fd }, 300000);
    if(!r.ok){ const t=await r.text(); throw new Error(t || 'Conversion failed'); }
    setStatus('Downloading output...', 92);
    const blob = await r.blob();
    let filename = 'converted_output';
    const disp = r.headers.get('content-disposition') || '';
    const m = disp.match(/filename\*=UTF-8''([^;]+)|filename="?([^";]+)"?/i);
    if(m) filename = decodeURIComponent(m[1] || m[2]);
    const dlUrl = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = dlUrl; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(dlUrl);
    setStatus('Done. File downloaded.',100);
  }catch(e){
    console.error(e);
    setStatus('Error: ' + (e.name === 'AbortError' ? 'timeout' : e.message),100);
  } finally {
    $('convertBtn').disabled=false;
  }
}
