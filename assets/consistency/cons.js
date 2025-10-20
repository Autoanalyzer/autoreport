// Consistency certificate app (lightweight)
// Implements: preview, PDF + Drive upload. Keeps code isolated under assets/consistency

// --- State ---
const state = {
  certificateNo: '',
  verifyDate: '',
  tagNo: '', tagName: '', location: '', maker: '',
  sensorModel: '', sensorSerial: '', transmitterModel: '', transmitterSerial: '',
  rangeMin: '', rangeMax: '', unit: '%Cs',
  funcL: [
    {desc:'Remove', status:'YES', complete:'YES'},
    {desc:'Install', status:'YES', complete:'YES'},
    {desc:'Supply', status:'YES', complete:'YES'},
    {desc:'Cable connector', status:'YES', complete:'YES'}
  ],
  funcR: [
    {desc:'Support (Sensor&Tx.)', status:'YES', complete:'YES'},
    {desc:'Cover (Sensor&Tx.)', status:'YES', complete:'YES'},
    {desc:'Terminal cable (Sensor&Tx.)', status:'YES', complete:'YES'},
    {desc:'Test Sensor', status:'YES', complete:'YES'}
  ],
  paramTemplate: 'standard',
  paramsL: [],
  paramsR: [],
  technician:'', technicianPosition:'', approver:'', signDate:'', remarks:'',
  images: {logo:null, photo1:null, photo2:null, photo3:null}
};

// --- Config ---
const APP = window.APP_CONFIG || {};
let GOOGLE_CLIENT_ID = APP.googleClientId || (typeof localStorage!=='undefined' && localStorage.getItem('gdrive_client_id')) || '';
const CONS_FOLDER_ID = (APP.consistencyFolderId || '19AuhoVwj3vXhwwFvx2t0JjabXlDmeoFu');
const GAS_URL = APP.gasUrl || '';
const GAS_SECRET = APP.gasSecret || '';
const GAS_FOLDER_ID = (APP.consistencyFolderId || CONS_FOLDER_ID);
let googleTokenClient = null, googleAccessToken = null;

// --- Tag presets (built-in + user) ---
function builtinTagPresets(){
  // Mapping: tag -> preset
  return {
    '421QC205': {tag:'421QC205', tagName:'primary screen feed consistency', location:'Pulp1', maker:'MetsoSP', sensorModel:'', sensorSerial:'', transmitterModel:'A4730024', transmitterSerial:'153374401', rangeMin:2, rangeMax:6, unit:'%Cs'},
    '421QC412': {tag:'421QC412', tagName:'pulp to reactor consistency', location:'Pulp1', maker:'ValmetSP', sensorModel:'T00162', sensorSerial:'161675640', transmitterModel:'A4730024', transmitterSerial:'61675640', rangeMin:4, rangeMax:16, unit:'%Cs'},
    '421QC507': {tag:'421QC507', tagName:'post O2 press E003 pulp consistency', location:'Pulp1', maker:'ValmetSP', sensorModel:'T00162', sensorSerial:'153374401', transmitterModel:'A4730024', transmitterSerial:'104965875', rangeMin:2, rangeMax:7, unit:'%Cs'},
    '411QC001': {tag:'411QC001', tagName:'-', location:'Pulp1', maker:'ValmetSP', sensorModel:'', sensorSerial:'', transmitterModel:'', transmitterSerial:'', rangeMin:3, rangeMax:10, unit:'%Cs'},
    '421QC005': {tag:'421QC005', tagName:'-', location:'Pulp1', maker:'ValmetSP', sensorModel:'', sensorSerial:'', transmitterModel:'A4730024', transmitterSerial:'', rangeMin:2, rangeMax:7, unit:'%Cs'},
    '421QC004': {tag:'421QC004', tagName:'-', location:'Pulp1', maker:'ValmetSP', sensorModel:'', sensorSerial:'', transmitterModel:'', transmitterSerial:'', rangeMin:3, rangeMax:6, unit:'%Cs'},
    '431QC078': {tag:'431QC078', tagName:'dewatpress E004 pulp consistency', location:'Pulp1', maker:'ValmetSP', sensorModel:'', sensorSerial:'', transmitterModel:'K17585', transmitterSerial:'212706533', rangeMin:4, rangeMax:10, unit:'%Cs'},
    '431QC120': {tag:'431QC120', tagName:'wash press E002 pulp consistency', location:'Pulp1', maker:'KPM', sensorModel:'KC3', sensorSerial:'', transmitterModel:'', transmitterSerial:'', rangeMin:2, rangeMax:7, unit:'%Cs'},
    '431QC502': {tag:'431QC502', tagName:'coarse screeen pulp consistency', location:'Pulp1', maker:'ValmetSP', sensorModel:'', sensorSerial:'081961446', transmitterModel:'', transmitterSerial:'', rangeMin:2, rangeMax:6, unit:'%Cs'},
    '431QC149': {tag:'431QC149', tagName:'dewatpress E004 pulp consistency', location:'Pulp1', maker:'MetsoSP', sensorModel:'', sensorSerial:'193781264', transmitterModel:'', transmitterSerial:'', rangeMin:3, rangeMax:5, unit:'%Cs'},
    '431QT6648': {tag:'431QT6648', tagName:'-', location:'Pulp1', maker:'ValmetSP', sensorModel:'', sensorSerial:'193781264', transmitterModel:'', transmitterSerial:'', rangeMin:3, rangeMax:5, unit:'%Cs'},
    '431QC646': {tag:'431QC646', tagName:'-', location:'Pulp1', maker:'ValmetSP', sensorModel:'', sensorSerial:'1624546117', transmitterModel:'', transmitterSerial:'', rangeMin:2, rangeMax:6, unit:'%Cs'},
    '431QC322': {tag:'431QC322', tagName:'-', location:'Pulp1', maker:'ValmetSP', sensorModel:'', sensorSerial:'', transmitterModel:'A4730024', transmitterSerial:'081961487', rangeMin:3.5, rangeMax:6.5, unit:'%Cs'},
    '432QC063': {tag:'432QC063', tagName:'-', location:'Pulp2', maker:'MetsoSP', sensorModel:'', sensorSerial:'155074790', transmitterModel:'', transmitterSerial:'155074790', rangeMin:2, rangeMax:6, unit:'%Cs'},
    '432QC002': {tag:'432QC002', tagName:'-', location:'Pulp2', maker:'KPM', sensorModel:'', sensorSerial:'', transmitterModel:'', transmitterSerial:'16483085', rangeMin:2, rangeMax:6, unit:'%Cs'},
    '432QC033': {tag:'432QC033', tagName:'-', location:'Pulp2', maker:'KPM', sensorModel:'KC/3', sensorSerial:'18233283', transmitterModel:'', transmitterSerial:'18233283', rangeMin:6, rangeMax:14, unit:'%Cs'},
    '432QC127': {tag:'432QC127', tagName:'-', location:'Pulp2', maker:'ValmetSP', sensorModel:'', sensorSerial:'', transmitterModel:'', transmitterSerial:'182179142', rangeMin:2, rangeMax:6.5, unit:'%Cs'},
    '432QC185': {tag:'432QC185', tagName:'-', location:'Pulp2', maker:'KPM', sensorModel:'KC/3', sensorSerial:'18453341', transmitterModel:'', transmitterSerial:'18423404', rangeMin:2, rangeMax:8, unit:'%Cs'},
    '432QC302': {tag:'432QC302', tagName:'-', location:'Pulp2', maker:'KPM', sensorModel:'KC/3', sensorSerial:'20203475', transmitterModel:'', transmitterSerial:'20203475', rangeMin:2, rangeMax:8, unit:'%Cs'},
    '432QC153': {tag:'432QC153', tagName:'-', location:'Pulp2', maker:'ValmetSP', sensorModel:'', sensorSerial:'', transmitterModel:'', transmitterSerial:'', rangeMin:6, rangeMax:14, unit:'%Cs'},
    '432QC004': {tag:'432QC004', tagName:'-', location:'Pulp2', maker:'KPM', sensorModel:'KC/3', sensorSerial:'', transmitterModel:'', transmitterSerial:'23163787', rangeMin:2, rangeMax:8, unit:'%Cs'},
    '422QC046': {tag:'422QC046', tagName:'-', location:'Pulp2', maker:'KPM', sensorModel:'KC/4', sensorSerial:'', transmitterModel:'', transmitterSerial:'15462946', rangeMin:2, rangeMax:6, unit:'%Cs'},
    '422QC062': {tag:'422QC062', tagName:'-', location:'Pulp2', maker:'MetsoSP', sensorModel:'', sensorSerial:'', transmitterModel:'', transmitterSerial:'104765938', rangeMin:'', rangeMax:'', unit:'%Cs'}
  };
}
function localTagKey(){ return 'consistency-tag-presets'; }
function loadLocalTagPresets(){ try{ return JSON.parse(localStorage.getItem(localTagKey())||'{}')||{}; }catch(_){ return {}; } }
function saveLocalTagPresets(map){ try{ localStorage.setItem(localTagKey(), JSON.stringify(map)); return true; }catch(_){ return false; } }
function mergedTagPresets(){ const base=builtinTagPresets(); const add=loadLocalTagPresets(); return Object.assign({}, base, add); }
function ensureTagOptionInSelect(tag){ try{ const s=document.querySelector('select[name="tagNo"]'); if(!s||!tag) return; if(![...s.options].some(o=>o.value===tag)){ const opt=document.createElement('option'); opt.value=tag; opt.textContent=tag; s.appendChild(opt);} }catch(_){ } }
function enhanceTagSelect(){ const s=document.querySelector('select[name="tagNo"]'); if(!s) return; const map=mergedTagPresets(); const keys=Object.keys(map).sort((a,b)=>String(a).localeCompare(String(b))); s.innerHTML = '<option value="">Select…</option><option value="__new__">+ Add new Tag</option>' + keys.map(k=>`<option value="${k}">${k}</option>`).join(''); s.addEventListener('change', ()=>{ const v=s.value; if(v==='__new__'){ const code = prompt('New Tag No.:',''); if(!code){ s.value=''; return; } const tag = String(code).trim(); state.tagNo = tag; // immediately persist this tag preset (from current form)
      saveCurrentTagPreset(); ensureTagOptionInSelect(tag); s.value = tag; applyStateToForm(); updatePreview(); toast('Saved tag preset'); } else if(v){ applyTagPreset(v); } }); if(state.tagNo){ ensureTagOptionInSelect(state.tagNo); s.value = state.tagNo; } }
function applyTagPreset(tag){ const map=mergedTagPresets(); const p = map[tag]; if(!p) return; state.tagNo = tag; state.tagName = p.tagName||''; state.location = p.location||''; state.maker = p.maker||''; state.sensorModel = p.sensorModel||''; state.sensorSerial = p.sensorSerial||''; state.transmitterModel = p.transmitterModel||''; state.transmitterSerial = p.transmitterSerial||''; state.rangeMin = (p.rangeMin==null?'':p.rangeMin); state.rangeMax = (p.rangeMax==null?'':p.rangeMax); state.unit = p.unit||'%Cs'; applyStateToForm(); updatePreview(); }
function saveCurrentTagPreset(){ const tag = (state.tagNo||'').trim(); if(!tag) return false; const add = loadLocalTagPresets(); add[tag] = { tag, tagName: state.tagName||'', location: state.location||'', maker: state.maker||'', sensorModel: state.sensorModel||'', sensorSerial: state.sensorSerial||'', transmitterModel: state.transmitterModel||'', transmitterSerial: state.transmitterSerial||'', rangeMin: state.rangeMin||'', rangeMax: state.rangeMax||'', unit: state.unit||'%Cs' }; const ok = saveLocalTagPresets(add); try{ enhanceTagSelect(); ensureTagOptionInSelect(tag); const sel=document.querySelector('select[name="tagNo"]'); if(sel) sel.value=tag; }catch(_){ } return ok; }
function deleteTagPreset(tag){
  if(!tag) return false;
  const base=builtinTagPresets();
  if(base[tag]) return false; // do not delete built-in presets
  const add=loadLocalTagPresets();
  if(!add[tag]) return false;
  delete add[tag];
  const ok = saveLocalTagPresets(add);
  // Reflect in UI immediately
  try{
    const sel=document.querySelector('select[name="tagNo"]');
    if(sel){
      const opt = Array.from(sel.options).find(o=>o.value===tag);
      if(opt) opt.remove();
      if(sel.value===tag) sel.value='';
    }
  }catch(_){ }
  return ok;
}

// --- Utils ---
const $ = (s, e=document) => e.querySelector(s);
const $$ = (s, e=document) => Array.from(e.querySelectorAll(s));
function toast(m){ try{ let b=document.getElementById('toast'); if(!b){ b=document.createElement('div'); b.id='toast'; b.style.cssText='position:fixed;right:12px;bottom:12px;z-index:9999'; document.body.appendChild(b);} const el=document.createElement('div'); el.textContent=m; el.style.cssText='margin-top:6px;background:#111827;color:#fff;padding:10px 12px;border-radius:8px;box-shadow:0 8px 22px rgba(0,0,0,.2);font-size:14px'; b.appendChild(el); setTimeout(()=>el.remove(),2400);}catch(_){}}
function busy(t){ let o=document.getElementById('busy-ov'); if(!o){ o=document.createElement('div'); o.id='busy-ov'; o.style.cssText='position:fixed;inset:0;background:rgba(17,24,39,.45);display:flex;align-items:center;justify-content:center;z-index:9998'; const m=document.createElement('div'); m.id='busy-ov-msg'; m.style.cssText='background:#fff;padding:12px 16px;border-radius:10px;box-shadow:0 16px 40px rgba(0,0,0,.25);font-weight:700'; m.textContent=t||'Working...'; o.appendChild(m); document.body.appendChild(o);} else { $('#busy-ov-msg').textContent=t||'Working...'; o.style.display='flex'; } }
function idle(){ const o=document.getElementById('busy-ov'); if(o) o.style.display='none'; }
function formatDateDMY(v){ if(!v) return ''; try{ const d=new Date(v); if(isNaN(+d)) return ''; const dd=String(d.getDate()).padStart(2,'0'), mm=String(d.getMonth()+1).padStart(2,'0'), yy=d.getFullYear(); return `${dd}/${mm}/${yy}`; }catch(_){ return ''; } }
function checkIcon(v){ return (String(v).toUpperCase()==='YES') ? '☑' : '☐'; }
function checkBoxHTML(v){
  try{
    const on = String(v).toUpperCase()==='YES';
    return '<span class="cb '+(on?'ok':'off')+'">'+(on?'✓':'')+'</span>';
  }catch(_){ return '<span class="cb off"></span>'; }
}
function setImg(id,u){ const el=document.getElementById(id); if(!el) return; el.src=u||'data:image/gif;base64,R0lGODlhAQABAAAAACw='; el.style.visibility='visible'; }

// --- Certificate number (AC/YYYY/NNNN) like pH & Conductivity ---
function consCertKey(year){ return 'cons-cert-seq-' + year; }
function formatCertNo(year, num){ return `AC/${year}/${String(num).padStart(4,'0')}`; }
function parseCertNo(str){ const m = /^AC\/(\d{4})\/(\d{3,})$/i.exec(String(str||'').trim()); return m ? { year:Number(m[1]), num:Number(m[2]) } : null; }
function getNextSeq(year){ let base = 149; try{ const v = localStorage.getItem(consCertKey(year)); if(v!=null && v!=='' && !isNaN(+v)) base = Number(v); }catch(_){ } return base + 1; }
function markSeqUsed(year, num){ try{ const cur = Number(localStorage.getItem(consCertKey(year)) || '149'); if(num > cur) localStorage.setItem(consCertKey(year), String(num)); }catch(_){ } }
function ensureCertificateNo(){ const y = new Date().getFullYear(); const p = parseCertNo(state.certificateNo); const next = getNextSeq(y); if(p && p.year===y){ const num = p.num < 150 ? next : Math.max(p.num, next); state.certificateNo = formatCertNo(y, num); } else { state.certificateNo = formatCertNo(y, next); } try{ const f=document.getElementById('data-form'); const el=f && f.elements['certificateNo']; if(el) el.value=state.certificateNo; }catch(_){ } }
function bumpCertificateNoAfterSave(){ const y = new Date().getFullYear(); const p = parseCertNo(state.certificateNo); const used = (p && p.year===y) ? p.num : getNextSeq(y); markSeqUsed(y, used); const next = used + 1; state.certificateNo = formatCertNo(y, next); try{ const f=document.getElementById('data-form'); const el=f && f.elements['certificateNo']; if(el) el.value=state.certificateNo; }catch(_){ } updatePreview(); }

// --- Form <-> state ---
function readForm(){
  const f = $('#data-form'); if(!f) return;
  const get = name => (f.elements[name] ? f.elements[name].value : '');
  state.certificateNo = get('certificateNo');
  state.verifyDate = get('verifyDate');
  state.tagNo = get('tagNo');
  state.tagName = get('tagName');
  state.location = get('location');
  state.maker = get('maker');
  state.sensorModel = get('sensorModel');
  state.sensorSerial = get('sensorSerial');
  state.transmitterModel = get('transmitterModel');
  state.transmitterSerial = get('transmitterSerial');
  state.rangeMin = get('rangeMin');
  state.rangeMax = get('rangeMax');
  state.unit = get('unit') || '%Cs';
  state.paramTemplate = get('paramTemplate') || state.paramTemplate || 'standard';
  // function check
  for(let i=0;i<state.funcL.length;i++){
    const row = $('#fcL-'+i);
    if(!row) continue;
    const sEl=row.querySelector('[name="status"]');
    const cEl=row.querySelector('[name="complete"]');
    state.funcL[i].status = (sEl && sEl.type==='checkbox') ? (sEl.checked?'YES':'NO') : (sEl?sEl.value:'');
    state.funcL[i].complete = (cEl && cEl.type==='checkbox') ? (cEl.checked?'YES':'NO') : (cEl?cEl.value:'');
  }
  for(let i=0;i<state.funcR.length;i++){
    const row = $('#fcR-'+i);
    if(!row) continue;
    const sEl=row.querySelector('[name="status"]');
    const cEl=row.querySelector('[name="complete"]');
    state.funcR[i].status = (sEl && sEl.type==='checkbox') ? (sEl.checked?'YES':'NO') : (sEl?sEl.value:'');
    state.funcR[i].complete = (cEl && cEl.type==='checkbox') ? (cEl.checked?'YES':'NO') : (cEl?cEl.value:'');
  }
  // params
  for(let i=0;i<state.paramsL.length;i++){
    const el = $('#pl-'+i); if(el){ state.paramsL[i].v = el.value; }
  }
  for(let i=0;i<state.paramsR.length;i++){
    const el = $('#pr-'+i); if(el){ state.paramsR[i].v = el.value; }
  }
  state.technician = get('technician');
  state.approver = get('approver');
  state.signDate = get('signDate');
  state.remarks = get('remarks');
}

function bindText(){ $$('[data-bind]').forEach(el=>{ const k=el.getAttribute('data-bind'); el.textContent = state[k] || '';}); }
function renderFuncChecks(){
  const mk = (row)=>`<tr><td>${row.desc}</td><td class="tcenter">${checkBoxHTML(row.status)}</td><td class="tcenter">${checkBoxHTML(row.status==='YES'?'NO':'YES')}</td><td class="tcenter">${checkBoxHTML(row.complete)}</td><td class="tcenter">${checkBoxHTML(row.complete==='YES'?'NO':'YES')}</td></tr>`;
  const left = $('#fc-left'); const right = $('#fc-right');
  if(left) left.innerHTML = state.funcL.map(mk).join('');
  if(right) right.innerHTML = state.funcR.map(mk).join('');
  // Equalize row heights across both tables for aligned borders
  try{
    requestAnimationFrame(()=>{
      const L = Array.from(document.querySelectorAll('#fc-left tr'));
      const R = Array.from(document.querySelectorAll('#fc-right tr'));
      const n = Math.max(L.length, R.length);
      for(let i=0;i<n;i++){
        const a=L[i], b=R[i]; if(!a||!b) continue;
        a.style.height='auto'; b.style.height='auto';
        const h=Math.max(a.offsetHeight||0, b.offsetHeight||0);
        a.style.height=h+'px'; b.style.height=h+'px';
      }
    });
  }catch(_){ }
}
function renderParams(){
  const show = (v)=>{ const s = (v==null? '' : String(v)); return s.trim()==='' ? '-' : s; };
  const mkL = (r)=>`<tr><td>${r.k}</td><td class="tright">${show(r.v)}</td><td>${r.u||''}</td></tr>`;
  const mkR = (r)=>`<tr><td>${r.k}</td><td class="tright">${show(r.v)}</td></tr>`;
  const l = $('#params-left'); const r = $('#params-right');
  if(l) l.innerHTML = state.paramsL.map(mkL).join('');
  if(r) r.innerHTML = state.paramsR.map(mkR).join('');
}
function updatePreview(){
  bindText();
  renderFuncChecks();
  renderParams();
  // Apply logo (prefer inlined PNG to avoid SVG render issues in html2canvas)
  try{
    const el = document.querySelector('#logo-box img');
    if(el){ el.src = state.images.logo || 'assets/logo_nps.svg'; }
  }catch(_){ }
  setImg('img1', state.images.photo1);
  setImg('img2', state.images.photo2);
  setImg('img3', state.images.photo3);
}

// Ensure default logo is embedded as PNG data URL for reliable PDF rendering
async function ensureDefaultLogo(){
  try{
    if(state.images && state.images.logo) return;
    const resp = await fetch('assets/logo_nps.svg', { cache:'no-cache' });
    if(!resp.ok) return;
    const svgText = await resp.text();
    const blob = new Blob([svgText], { type:'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const img = new Image();
    await new Promise((res,rej)=>{ img.onload=res; img.onerror=rej; img.src=url; });
    const size = Math.max(img.naturalWidth||600, 600);
    const canvas = document.createElement('canvas');
    canvas.width=size; canvas.height=size;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(img,0,0,canvas.width,canvas.height);
    const dataUrl = canvas.toDataURL('image/png');
    URL.revokeObjectURL(url);
    if(!state.images) state.images={};
    state.images.logo = dataUrl;
  }catch(_){ }
}

// --- Signers (dropdowns like pH/Conductivity) ---
const CALIBRATORS=[
  {name:'Anusit s.',position:'Specialist'},
  {name:'Supakit b.',position:'Engineer'},
  {name:'Nattapol w.',position:'Technician'},
  {name:'Phirakrit a.',position:'Technician'},
  {name:'Phattharaphong b.',position:'Technician'},
  {name:'Amonnat S.',position:'Technician'},
  {name:'Chanin t.',position:'Engineer'},
];
const DEFAULT_APPROVERS=[{name:'Chanin T.',position:'Engineer'},{name:'Supakit B.',position:'Engineer'}];
function loadApprovers(){ try{ const a=JSON.parse(localStorage.getItem('approver-list')||'[]'); return Array.isArray(a)?a:[]; }catch(_){ return []; } }
function saveApprovers(list){ try{ localStorage.setItem('approver-list', JSON.stringify(list)); }catch(_){ } }
function ensureApproverSelect(){ const sel=document.getElementById('approver-select'); if(!sel) return; let list=loadApprovers(); if(!list||list.length===0) list=DEFAULT_APPROVERS.slice(); sel.innerHTML='<option value="">Select…</option><option value="__new__">+ Add approver</option>'+list.map(a=>`<option value="${a.name}" data-pos="${a.position||''}">${a.name}</option>`).join(''); if(state.approver && ![...sel.options].some(o=>o.value===state.approver)){ const opt=document.createElement('option'); opt.value=state.approver; opt.textContent=state.approver; sel.appendChild(opt);} sel.value=state.approver||''; sel.onchange=()=>{ const v=sel.value; if(v==='__new__'){ const name=prompt('Approver name:',''); if(!name){ sel.value=''; return; } const pos=prompt('Position:','Engineer')||'Engineer'; const arr=loadApprovers(); if(!arr.some(x=>x&&x.name===name)){ arr.push({name,position:pos}); saveApprovers(arr);} ensureApproverSelect(); sel.value=name; state.approver=name; updatePreview(); } else { state.approver=v; updatePreview(); } } }
function populateTechnicianSelect(){ const sel=document.getElementById('technician-select'); if(!sel) return; sel.innerHTML='<option value="">Select…</option>'+CALIBRATORS.map(c=>`<option value="${c.name}" data-pos="${c.position||''}">${c.name}</option>`).join(''); if(state.technician && ![...sel.options].some(o=>o.value===state.technician)){ const opt=document.createElement('option'); opt.value=state.technician; opt.textContent=state.technician; opt.setAttribute('data-pos', state.technicianPosition||''); sel.appendChild(opt);} sel.value=state.technician||''; sel.onchange=()=>{ state.technician = sel.value || ''; const pos = sel.selectedOptions[0]?.getAttribute('data-pos') || ''; state.technicianPosition = pos; const hint=document.getElementById('tech-role-inline'); if(hint) hint.textContent = pos ? `(${pos})` : ''; updatePreview(); }; const hint=document.getElementById('tech-role-inline'); if(hint) hint.textContent = state.technicianPosition ? `(${state.technicianPosition})` : ''; }

// --- Preview modal (reuses structure from conductivity) ---
function ensurePreviewModal(){ let modal = document.getElementById('preview-modal'); if(!modal){ modal = document.createElement('div'); modal.id='preview-modal'; modal.className='modal hidden'; modal.innerHTML = `
      <div class="modal-backdrop" data-close></div>
      <div class="modal-dialog">
        <header class="modal-head">
          <div class="mh-title">Preview</div>
          <button id="preview-close" class="modal-close" type="button" aria-label="Close">Close</button>
        </header>
        <div class="modal-body"><div id="preview-mount"></div></div>
        <footer class="modal-foot"><button id="preview-close-2" type="button">Close</button></footer>
      </div>`; document.body.appendChild(modal);} return modal; }
async function openPreviewModal(){ try{ await ensureDefaultLogo(); updatePreview(); ensurePreviewModal(); const m=$('#preview-modal'), mount=$('#preview-mount'), base=$('#certificate'); if(!m||!mount||!base){ alert('Preview container missing'); return; } const c=base.cloneNode(true); c.id='certificate-preview'; c.classList.remove('pdf'); mount.innerHTML=''; mount.appendChild(c); m.classList.remove('hidden'); document.body.style.overflow='hidden'; const close=()=>{ m.classList.add('hidden'); mount.innerHTML=''; document.body.style.overflow=''; }; $('#preview-close')?.addEventListener('click',close,{once:true}); $('#preview-close-2')?.addEventListener('click',close,{once:true}); m.querySelector('[data-close]')?.addEventListener('click',close,{once:true}); const onKey=e=>{ if(e.key==='Escape'){ close(); window.removeEventListener('keydown',onKey);} }; window.addEventListener('keydown',onKey);}catch(err){ console.error('Preview error', err); alert('Cannot open preview: '+(err&&err.message?err.message:err)); }}

// --- PDF + upload ---
function makeBaseFilename(){ const tag=(state.tagNo||'TAG').replace(/[^\w\-]+/g,'_'); const date = state.verifyDate || new Date().toISOString().slice(0,10); return `Consistency_${tag}_${date}`; }
async function waitImages(el){ const imgs=[...el.querySelectorAll('img')]; await Promise.all(imgs.map(i=>i.complete&&i.naturalWidth?null:new Promise(r=>{i.onload=i.onerror=()=>r()}))); await Promise.all(imgs.map(i=>i.decode?.().catch(()=>{}))); }
async function generateAndUploadPDF(){ const base=$('#certificate'); if(!base) return; await ensureDefaultLogo(); updatePreview(); const filename = makeBaseFilename()+'.pdf'; const box=document.createElement('div'); box.id='pdf-sandbox'; box.style.cssText='position:fixed;left:0;top:0;width:210mm;z-index:-1;opacity:0;pointer-events:none;'; const el=base.cloneNode(true); el.id='certificate-print'; document.body.appendChild(box); box.appendChild(el); el.classList.add('pdf'); await waitImages(el); try{ if(document.fonts&&document.fonts.ready){ await document.fonts.ready;} }catch(_){ }
  const opt={ margin:[0,0,0,0], filename, image:{type:'jpeg',quality:0.98}, html2canvas:{scale:2,useCORS:true,backgroundColor:'#ffffff',scrollX:0,scrollY:0}, jsPDF:{unit:'mm',format:'a4',orientation:'portrait'} };
  try{
    busy('Creating PDF...');
    let blob=null; await window.html2pdf().from(el).set(opt).toPdf().get('pdf').then(pdf=>{ try{ const n=typeof pdf.getNumberOfPages==='function'?pdf.getNumberOfPages():1; for(let i=n;i>=2;i--){ pdf.deletePage(i);} }catch(_){ } blob=pdf.output('blob'); });
    busy('Uploading to Google Drive...');
    if(GAS_URL){ const res = await uploadViaAppsScript(blob, filename); const url=(res&&(res.url||res.webViewLink||(res.fileId?('https://drive.google.com/file/d/'+res.fileId+'/view'):'')))||''; try{ saveCurrentTagPreset(); bumpCertificateNoAfterSave(); }catch(_){ } toast('Saved: '+filename+(url?' ✓':'')); }
    else { const meta = await uploadToDriveQuiet(blob, filename); const url=meta&&(meta.webViewLink||(meta.id?('https://drive.google.com/file/d/'+meta.id+'/view'):'')); try{ saveCurrentTagPreset(); bumpCertificateNoAfterSave(); }catch(_){ } toast('Saved: '+filename+(url?' ✓':'')); }
  }catch(err){ console.error(err); toast('Save failed'); }
  finally{ idle(); try{ box.remove(); }catch(_){ } }
}

// --- Google Drive upload helpers ---
function ensureGoogleToken(scopes){ return new Promise((resolve,reject)=>{ try{
  if(googleAccessToken){ resolve(googleAccessToken); return; }
  if(!GOOGLE_CLIENT_ID){ reject(new Error('Missing Google Client ID')) ;return; }
  googleTokenClient = google.accounts.oauth2.initTokenClient({
    client_id: GOOGLE_CLIENT_ID,
    scope: scopes || 'https://www.googleapis.com/auth/drive.file',
    callback: (token)=>{ try{ googleAccessToken = token && token.access_token; resolve(googleAccessToken);}catch(e){ reject(e);} }
  });
  googleTokenClient.requestAccessToken({prompt:''});
 }catch(err){ reject(err); }
}); }

async function uploadToDriveQuiet(blob, filename){ const token = await ensureGoogleToken('https://www.googleapis.com/auth/drive.file'); const meta = { name: filename, mimeType:'application/pdf', parents:[CONS_FOLDER_ID] }; const form = new FormData(); form.append('metadata', new Blob([JSON.stringify(meta)], {type:'application/json'})); form.append('file', blob); const res = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', { method:'POST', headers:{ 'Authorization':'Bearer '+token }, body: form }); if(!res.ok){ const t=await res.text(); throw new Error('Drive upload failed: '+t); } return res.json(); }

async function uploadViaAppsScript(blob, filename){
  // Match the same payload format as conductivity module (JSON with base64 file)
  const b64 = await new Promise((resolve, reject)=>{ const r=new FileReader(); r.onload=()=>resolve(String(r.result).split(',')[1]); r.onerror=reject; r.readAsDataURL(blob); });
  const res = await fetch(GAS_URL, {
    method:'POST',
    headers:{ 'Content-Type':'text/plain;charset=utf-8' },
    body: JSON.stringify({
      secret: GAS_SECRET || undefined,
      folderId: GAS_FOLDER_ID || undefined,
      filename,
      mimeType: 'application/pdf',
      fileBase64: b64
    })
  });
  const text = await res.text();
  let json = {};
  try{ json = JSON.parse(text); }catch(_){ }
  if(!res.ok || (json && json.success===false)){
    const m = (json && (json.error||json.message)) || text || ('HTTP '+res.status);
    throw new Error('Apps Script upload failed: '+m);
  }
  return json;
}

// --- Wiring ---
function wireForm(){ const f=$('#data-form'); if(!f) return; f.addEventListener('input', ()=>{ readForm(); updatePreview(); }); f.addEventListener('change', ()=>{ readForm(); updatePreview(); }); $$('input[type="file"][name^="photo"]').forEach(inp=>{ inp.addEventListener('change', ()=>{ const f=inp.files&&inp.files[0]; if(!f){ const k=inp.name; state.images[k]=null; updatePreview(); return; } const r=new FileReader(); r.onload=()=>{ state.images[inp.name]=String(r.result); updatePreview(); }; r.readAsDataURL(f); }); }); $('#btn-save-json')?.addEventListener('click', ()=>{ const data = JSON.stringify(state, null, 2); const blob = new Blob([data], {type:'application/json'}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download = makeBaseFilename()+'.json'; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href), 1000); }); $('#input-load-json')?.addEventListener('change', (e)=>{ const f=e.target.files&&e.target.files[0]; if(!f) return; const r=new FileReader(); r.onload=()=>{ try{ const o=JSON.parse(String(r.result)); Object.assign(state,o||{}); // push values back to form
      applyStateToForm(); updatePreview(); }catch(err){ alert('Invalid JSON'); } }; r.readAsText(f); }); $('#btn-load-demo')?.addEventListener('click', ()=>{ loadDemo(); applyStateToForm(); updatePreview(); }); $('#btn-preview')?.addEventListener('click', openPreviewModal); $('#btn-generate')?.addEventListener('click', generateAndUploadPDF); $('#cta-preview')?.addEventListener('click', openPreviewModal); $('#cta-save')?.addEventListener('click', generateAndUploadPDF);
  // Tag select wiring
  enhanceTagSelect();
  // Signers dropdowns
  populateTechnicianSelect();
  ensureApproverSelect();
  // Update preset button
  const btn=document.getElementById('btn-update-tag'); if(btn) btn.addEventListener('click', ()=>{ readForm(); if(saveCurrentTagPreset()){ toast('Saved tag preset'); } else { alert('Cannot save preset'); } });
  // Add delete button next to update (only for local presets)
  try{ var wrap=document.querySelector('.table-controls'); if(wrap && !document.getElementById('btn-delete-tag')){ var del=document.createElement('button'); del.id='btn-delete-tag'; del.type='button'; del.textContent='Delete this Tag'; del.style.marginLeft='8px'; wrap.appendChild(del); del.addEventListener('click', function(){ var tag=(state.tagNo||'').trim(); if(!tag){ alert('Please select a Tag first'); return; } var base=builtinTagPresets(); if(base[tag]){ alert('Cannot delete built-in Tag'); return; } if(!confirm('Delete this Tag preset?')) return; if(deleteTagPreset(tag)){ // clear state and refresh UI
        state.tagNo=''; state.tagName=''; state.location=''; state.maker=''; state.sensorModel=''; state.sensorSerial=''; state.transmitterModel=''; state.transmitterSerial=''; state.rangeMin=''; state.rangeMax='';
        applyStateToForm(); enhanceTagSelect(); updatePreview(); toast('Deleted tag '+tag); } }); } }catch(_){ }
  // Parameter template select
  const tplSel = document.getElementById('param-template-select');
  if(tplSel){ tplSel.addEventListener('change', ()=>{ state.paramTemplate = tplSel.value || 'standard'; applyParamTemplate(state.paramTemplate); applyStateToForm(); updatePreview(); }); }
  // Build parameter form once in init
  buildParamForm();
}

function applyStateToForm(){ const f=$('#data-form'); if(!f) return; const set=(n,v)=>{ if(f.elements[n]) f.elements[n].value = (v==null?'':v); };
  set('certificateNo', state.certificateNo); set('verifyDate', state.verifyDate); set('tagNo', state.tagNo); set('tagName', state.tagName); set('location', state.location); set('maker', state.maker); set('sensorModel', state.sensorModel); set('sensorSerial', state.sensorSerial); set('transmitterModel', state.transmitterModel); set('transmitterSerial', state.transmitterSerial); set('rangeMin', state.rangeMin); set('rangeMax', state.rangeMax); set('unit', state.unit); set('paramTemplate', state.paramTemplate);
  for(let i=0;i<state.funcL.length;i++){ const row=$('#fcL-'+i); if(row){ const se=row.querySelector('[name="status"]'); const ce=row.querySelector('[name="complete"]'); if(se){ if(se.type==='checkbox') se.checked = (String(state.funcL[i].status).toUpperCase()==='YES'); else se.value = state.funcL[i].status; } if(ce){ if(ce.type==='checkbox') ce.checked = (String(state.funcL[i].complete).toUpperCase()==='YES'); else ce.value = state.funcL[i].complete; } } }
  for(let i=0;i<state.funcR.length;i++){ const row=$('#fcR-'+i); if(row){ const se=row.querySelector('[name="status"]'); const ce=row.querySelector('[name="complete"]'); if(se){ if(se.type==='checkbox') se.checked = (String(state.funcR[i].status).toUpperCase()==='YES'); else se.value = state.funcR[i].status; } if(ce){ if(ce.type==='checkbox') ce.checked = (String(state.funcR[i].complete).toUpperCase()==='YES'); else ce.value = state.funcR[i].complete; } } }
  buildParamForm();
  for(let i=0;i<state.paramsL.length;i++){ const el=$('#pl-'+i); if(el) el.value = state.paramsL[i].v ?? ''; }
  for(let i=0;i<state.paramsR.length;i++){ const el=$('#pr-'+i); if(el) el.value = state.paramsR[i].v ?? ''; }
  set('technician', state.technician); set('approver', state.approver); set('signDate', state.signDate); set('remarks', state.remarks);
}

function loadDemo(){ const today=new Date().toISOString().slice(0,10); const y=new Date().getFullYear(); const next=getNextSeq(y); state.certificateNo = formatCertNo(y, next); state.verifyDate = today; state.tagNo = '432QC302'; state.tagName='-'; state.location='Pulp2'; state.maker='KPM'; state.sensorModel='KX/3'; state.sensorSerial='20203475'; state.transmitterModel='-'; state.transmitterSerial='20203475'; state.rangeMin='2'; state.rangeMax='8'; state.unit='%Cs'; state.paramsL[0].v='3.8'; state.paramsL[1].v='4'; state.paramsL[2].v='4'; state.paramsL[3].v='27'; state.paramsL[4].v='27'; state.paramsL[5].v='27'; state.paramsL[6].v='-'; state.paramsR[1].v=''; state.paramsR[2].v='HV8'; state.paramsR[3].v='2.2'; state.paramsR[4].v='12'; state.paramsR[5].v='9'; state.paramsR[6].v='-'; state.technician='Technician A'; state.approver='Engineer B'; state.signDate=today; state.remarks=''; }

function onReady(fn){ if(document.readyState==='loading'){ document.addEventListener('DOMContentLoaded', fn, {once:true}); } else { fn(); } }

onReady(async function(){ try{ ensureCertificateNo(); await ensureDefaultLogo(); applyParamTemplate(state.paramTemplate||'standard'); wireForm(); applyStateToForm(); updatePreview(); }catch(err){ console.error(err); } });

// ----- Parameter templates & builders -----
const PARAM_TEMPLATES={
  standard:{
    left:[
      {k:'Consist Displayed',u:'%Cs'},
      {k:'Output1',u:'mA'},
      {k:'Output2',u:'mA'},
      {k:'Torque',u:'Nm'},
      {k:'Speed',u:'rpm'},
      {k:'Proc. T',u:'C'},
      {k:'Elec. T',u:'C'},
      {k:'Cons T',u:'C'},
      {k:'Motor T',u:'C'}
    ],
    right:[
      {k:'Sensor Type',v:'Standard ST',u:''},
      {k:'Mounting',u:''},
      {k:'Pulp Type',u:''},
      {k:'Recipe',u:''},
      {k:'P1 (Gain)',u:''},
      {k:'P2 (Offset)',u:''},
      {k:'Alarm',u:''}
    ]
  },
  rotary:{
    left:[
      {k:'Consist Displayed',u:'%Cs'},
      {k:'Output1',u:'mA'},
      {k:'Output2',u:'mA'},
      {k:'Torque',u:'Nm'},
      {k:'Speed',u:'rpm'},
      {k:'Proc. T',u:'C'},
      {k:'Elec. T',u:'C'},
      {k:'Cons T',u:'C'},
      {k:'Motor T',u:'C'}
    ],
    right:[
      {k:'Sensor Type',v:'Rotary',u:''},
      {k:'Mounting',v:'Hor Side',u:''},
      {k:'Pulp Type',v:'HWB',u:''},
      {k:'Recipe',v:'1',u:''},
      {k:'P1 (Gain)',v:'1.0000',u:''},
      {k:'P2 (Offset)',v:'0.3500',u:''},
      {k:'Alarm',v:'Motor Not Running',u:''}
    ]
  },
  hlss:{
    left:[
      {k:'Consist Displayed',u:'%Cs'},
      {k:'Output1',u:'mA'},
      {k:'Output2',u:'mA'},
      {k:'Shear Force',u:'G'},
      {k:'Speed',u:'rpm'},
      {k:'S Temp',u:'C'},
      {k:'P Temp',u:'C'},
      {k:'H Temp',u:'C'}
    ],
    right:[
      {k:'Sensor Type',v:'HLSS',u:''},
      {k:'Mounting',u:''},
      {k:'Pulp Type',u:''},
      {k:'Recipe',u:''},
      {k:'P1 (Gain)',u:''},
      {k:'P2 (Offset)',u:''},
      {k:'Alarm',u:''}
    ]
  },
  blade:{
    left:[
      {k:'Consist Displayed',u:'%Cs'},
      {k:'Output1',u:'mA'},
      {k:'Output2',u:'mA'},
      {k:'Torque',u:'G'},
      {k:'Speed',u:'rpm'},
      {k:'S Temp',u:'C'},
      {k:'P Temp',u:'C'},
      {k:'H Temp',u:'C'}
    ],
    right:[
      {k:'Sensor Type',v:'Blade RL',u:''},
      {k:'Mounting',u:''},
      {k:'Pulp Type',u:''},
      {k:'Recipe',u:''},
      {k:'P1 (Gain)',u:''},
      {k:'P2 (Offset)',u:''},
      {k:'Alarm',u:''}
    ]
  },
  blade_generic:{
    left:[
      {k:'Consist Displayed',u:'%Cs'},
      {k:'Output1',u:'mA'},
      {k:'Output2',u:'mA'},
      {k:'Torque',u:'G'},
      {k:'Speed',u:'rpm'},
      {k:'Proc. T',u:'C'},
      {k:'Elec. T',u:'C'},
      {k:'Cons T',u:'C'},
      {k:'Motor T',u:'C'}
    ],
    right:[
      {k:'Sensor Type',v:'Blade',u:''},
      {k:'Mounting',v:'Hor Up',u:''},
      {k:'Pulp Type',v:'HWB',u:''},
      {k:'Recipe',v:'1',u:''},
      {k:'P1 (Gain)',v:'1',u:''},
      {k:'P2 (Offset)',v:'155',u:''},
      {k:'Alarm',u:''}
    ]
  },
  blade_st:{
    left:[
      {k:'Consist Displayed',u:'%Cs'},
      {k:'Output1',u:'mA'},
      {k:'Output2',u:'mA'},
      {k:'Torque',u:'G'},
      {k:'Speed',u:'rpm'},
      {k:'Proc. T',u:'C'},
      {k:'Elec. T',u:'C'},
      {k:'Cons T',u:'C'},
      {k:'Motor T',u:'C'}
    ],
    right:[
      {k:'Sensor Type',v:'Blade ST',u:''},
      {k:'Mounting',v:'Ver up',u:''},
      {k:'Pulp Type',v:'SW',u:''},
      {k:'Recipe',u:''},
      {k:'P1 (Gain)',v:'2.1',u:''},
      {k:'P2 (Offset)',v:'-6.8',u:''},
      {k:'Alarm',u:''}
    ]
  },
  blade_jl:{
    left:[
      {k:'Consist Displayed',u:'%Cs'},
      {k:'Output1',u:'mA'},
      {k:'Output2',u:'mA'},
      {k:'Torque',u:'G'},
      {k:'Speed',u:'rpm'},
      {k:'S Temp',u:'C'},
      {k:'P Temp',u:'C'},
      {k:'H Temp',u:'C'}
    ],
    right:[
      {k:'Sensor Type',v:'Blade JL',u:''},
      {k:'Mounting',v:'Ver up',u:''},
      {k:'Pulp Type',v:'HWB',u:''},
      {k:'Recipe',v:'1',u:''},
      {k:'P1 (Gain)',v:'0.808',u:''},
      {k:'P2 (Offset)',v:'0.843',u:''},
      {k:'Alarm',u:''}
    ]
  },
  blade_md:{
    left:[
      {k:'Consist Displayed',u:'%Cs'},
      {k:'Output1',u:'mA'},
      {k:'Output2',u:'mA'},
      {k:'Torque',u:'G'},
      {k:'Speed',u:'rpm'},
      {k:'Proc. T',u:'C'},
      {k:'Elec. T',u:'C'},
      {k:'Cons T',u:'C'},
      {k:'Motor T',u:'C'}
    ],
    right:[
      {k:'Sensor Type',v:'Blade MD',u:''},
      {k:'Mounting',v:'Hor Side',u:''},
      {k:'Pulp Type',v:'HWB',u:''},
      {k:'Recipe',v:'1',u:''},
      {k:'P1 (Gain)',v:'2.2',u:''},
      {k:'P2 (Offset)',v:'-9',u:''},
      {k:'Alarm',u:''}
    ]
  }
};
function applyParamTemplate(id){ state.paramTemplate=id||'standard'; const t=PARAM_TEMPLATES[state.paramTemplate]||PARAM_TEMPLATES.standard; state.paramsL=t.left.map(r=>({k:r.k,u:r.u||'',v:''})); state.paramsR=t.right.map(r=>({k:r.k,u:r.u||'',v:r.v||''})); buildParamForm(); }
function buildParamForm(){
  try{
    const left = document.getElementById('params-left-body');
    const right = document.getElementById('params-right-body');
    if(left){
      left.innerHTML='';
      state.paramsL.forEach((r,i)=>{
        const tr=document.createElement('tr');
        tr.innerHTML = `<td>${r.k}</td><td><input id="pl-${i}" /></td><td>${r.u||''}</td>`;
        left.appendChild(tr);
      });
    }
    if(right){
      right.innerHTML='';
      state.paramsR.forEach((r,i)=>{
        const key = String(r.k).toLowerCase();
        const isMount = key==='mounting';
        const isPulp = key==='pulp type' || key==='pulp-type' || key==='pulptype';
        const control = isMount
          ? `<select id="pr-${i}"><option value=""></option><option value="Ver up">Ver Up</option><option value="Hor Side">Hor Side</option></select>`
          : (isPulp
              ? `<select id="pr-${i}"><option value=""></option><option value="HWB">HWB</option><option value="SW">SW</option></select>`
              : `<input id="pr-${i}" />`);
        const tr=document.createElement('tr');
        tr.innerHTML = `<td>${r.k}</td><td>${control}</td>`;
        right.appendChild(tr);
      });
    }
  }catch(_){ }
}
