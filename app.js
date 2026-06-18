const $ = (id) => document.getElementById(id);
const backendInput = $('backendUrl');
const saved = localStorage.getItem('pptEasyBackend') || 'http://localhost:8000';
backendInput.value = saved;

$('saveBackend').addEventListener('click', () => { localStorage.setItem('pptEasyBackend', backendInput.value.trim()); checkHealth(); });
$('pptFile').addEventListener('change', () => { const f=$('pptFile').files[0]; $('fileName').textContent = f ? f.name : 'PPTX choose karo'; });
$('convertBtn').addEventListener('click', convertNow);
checkHealth();

function setStatus(msg, pct=0){ $('status').textContent=msg; $('bar').style.width=Math.max(0,Math.min(100,pct))+'%'; }
function backend(){ return (backendInput.value || '').trim().replace(/\/$/,''); }
async function checkHealth(){
  try{ const r=await fetch(backend()+'/health'); const j=await r.json(); $('health').textContent = j.ok ? 'Backend connected ✅' : 'Backend issue'; }
  catch(e){ $('health').textContent = 'Backend not connected. Local/Render backend start karo.'; }
}
async function convertNow(){
  const file = $('pptFile').files[0];
  if(!file){ alert('PPTX choose karo.'); return; }
  if(!backend()){ alert('Backend URL required.'); return; }
  $('convertBtn').disabled=true; setStatus('Uploading PPT to backend...', 10);
  try{
    const fd = new FormData();
    fd.append('file', file);
    fd.append('output', $('outputType').value);
    fd.append('lang', $('ocrLang').value);
    fd.append('excel_mode', $('excelMode').value);
    fd.append('ppt_mode', $('pptMode').value);
    fd.append('include_image', $('includeImage').checked ? 'true':'false');
    const r = await fetch(backend()+'/api/convert', { method:'POST', body:fd });
    if(!r.ok){ const t=await r.text(); throw new Error(t || 'Conversion failed'); }
    setStatus('Downloading output...', 92);
    const blob = await r.blob();
    let filename = 'converted_output';
    const disp = r.headers.get('content-disposition') || '';
    const m = disp.match(/filename="?([^";]+)"?/i); if(m) filename=m[1];
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href=url; a.download=filename; document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url); setStatus('Done. File downloaded.',100);
  }catch(e){ console.error(e); setStatus('Error: '+e.message,100); }
  finally{ $('convertBtn').disabled=false; }
}
