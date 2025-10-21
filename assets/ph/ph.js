// main.js
// -------------------------------
// Basic state for the form and preview
const state = {
  certificateNo: "",
  calibrateDate: "",
  tagNo: "",
  tagName: "",
  location: "",
  maker: "",
  sensorModel: "",
  sensorSerial: "",
  transmitterModel: "",
  transmitterSerial: "",
  rangeMin: "",
  rangeMax: "",
  unit: "pH",
  acceptError: "",
  solutionType: "Buffer",
  slope: "",
  offset: "",
  calResult: "Calibrated",
  technician: "",
  technicianPosition: "",
  approver: "",
  signDate: "",
  remarks: "",
  buffers: [],
  points: [],
  // Calibrate readings for fixed buffers
  found4: "", found7: "", found10: "",
  cal4: "",   cal7: "",   cal10: "",
  images: { logo: null, photo1: null, photo2: null, photo3: null },
  photoLabels: { photo1: '', photo2: '', photo3: '' }
};

// -------------------------------
// Google Drive / Apps Script configuration
const APP = window.APP_CONFIG || {};
let GOOGLE_CLIENT_ID = APP.googleClientId || (typeof localStorage!=='undefined' && localStorage.getItem('gdrive_client_id')) || '';
const DRIVE_FOLDER_ID = APP.driveFolderId || '1ZAJn2vQJGfVugVchROfTNHdaxjMC6Zx7';
const UPLOAD_TO_DRIVE = (APP.uploadToDrive !== false);
const GAS_URL = APP.gasUrl || 'https://script.google.com/macros/s/AKfycbw1GKVDrhrl43vj5IiBGrxInBtwIJxh3ftLGh0GKs73YU9sBL0a86_0KgXQFdMhkq1N9w/exec';
const GAS_SECRET = APP.gasSecret || 'aa123456789';
const GAS_FOLDER_ID = APP.gasFolderId || DRIVE_FOLDER_ID || '1ZAJn2vQJGfVugVchROfTNHdaxjMC6Zx7';
let googleTokenClient = null;
let googleAccessToken = null;

// -------------------------------
// Calibrators dropdown
const CALIBRATORS = [
  { name: 'Anusit s.',        position: 'Specialist' },
  { name: 'Supakit b.',       position: 'Engineer' },
  { name: 'Nattapol w.',      position: 'Technician' },
  { name: 'Phirakrit a.',     position: 'Technician' },
  { name: 'Phattharaphong b.',position: 'Technician' },
  { name: 'Amonnat S.',       position: 'Technician' },
  { name: 'Chanin t.',        position: 'Engineer' },
];

// Approvers dropdown (persistable)
const DEFAULT_APPROVERS = [
  { name: 'Chanin T.', position: 'Engineer' },
  { name: 'Supakit B.', position: 'Engineer' },
];
function loadApprovers(){ try{ const a = JSON.parse(localStorage.getItem('approver-list')||'[]'); return Array.isArray(a) ? a : []; }catch(_){ return []; } }
function saveApprovers(list){ localStorage.setItem('approver-list', JSON.stringify(list)); }
function ensureApproverSelect(){
  const form = document.getElementById('data-form'); if (!form) return;
  const sel = form.elements['approver']; if (!sel) return;
  let list = loadApprovers(); if (!list || list.length === 0) { list = DEFAULT_APPROVERS.slice(); }
  // Remove deprecated name if present
  list = list.filter(a => a && String(a.name).trim().toLowerCase() !== 'anurak s.');
  saveApprovers(list);
  if (String(state.approver||'').trim().toLowerCase() === 'anurak s.') { state.approver = ''; }
  sel.innerHTML = '<option value="">เลือกผู้อนุมัติ</option><option value="__new__">+ เพิ่มชื่อใหม่…</option>' + list.map(a => `<option value="${a.name}" data-pos="${a.position||''}">${a.name}</option>`).join('');
  sel.addEventListener('change', () => {
    const v = sel.value;
    if (v === '__new__') {
      const name = prompt('ชื่อผู้อนุมัติใหม่:', ''); if (!name) { sel.value=''; return; }
      const pos = prompt('ตำแหน่ง:', 'Engineer') || 'Engineer';
      const arr = loadApprovers(); if (!arr.some(x => x && x.name === name)) { arr.push({ name, position: pos }); saveApprovers(arr); }
      ensureApproverSelect(); sel.value = name;
      state.approver = name; saveToLocal(); updatePreview();
    } else {
      state.approver = v; saveToLocal(); updatePreview();
    }
  }, { once: true });
  // Keep current value
  if (state.approver) { if (sel.value !== state.approver) { const opt = document.createElement('option'); opt.value = state.approver; opt.textContent = state.approver; sel.appendChild(opt); sel.value = state.approver; } }
}

function syncTechnicianUI(){
  const form = document.getElementById('data-form');
  if (!form) return;
  const sel = form.elements['technician'];
  if (!sel) return;
  // Ensure value matches state and update role text
  try {
    if (sel.tagName === 'SELECT') {
      // If state value isn't in options, keep it by adding
      if (state.technician && !Array.from(sel.options).some(o => o.value === state.technician)) {
        const opt = document.createElement('option'); opt.value = state.technician; opt.textContent = state.technician; opt.setAttribute('data-pos',''); sel.appendChild(opt);
      }
      sel.value = state.technician || '';
      const opt = sel.selectedOptions[0];
      state.technicianPosition = opt ? (opt.getAttribute('data-pos') || '') : '';
    }
  } catch (_) { /* noop */ }
  const inline = document.getElementById('tech-role-inline');
  if (inline) inline.textContent = state.technicianPosition ? `(${state.technicianPosition})` : '';
}

// -------------------------------
// Built-in presets (XMTR INFO)
// Map by Tag No. to equipment info; fields: brand (Maker), model (Transmitter Model),
// calRange ("min to max"), desc (Tag Name), location, sensorModel, sensorSerial, transmitterSerial,
// unit, acceptError.
const BUILTIN_TAG_PRESETS = {
  "431QC041": { desc:'White Water', location:'Pulp1', brand:'ABB', sensorModel:'TB564', sensorSerial:'-', model:'TB82', transmitterSerial:'-', calRange:'0 to 14', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "431QC117": { desc:'DO - Tower pH', location:'Pulp1', brand:'ABB', sensorModel:'TB557', sensorSerial:'-', model:'TB82', transmitterSerial:'-', calRange:'0 to 7', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "431QC316": { desc:'EOP Tower pH', location:'Pulp1', brand:'ABB', sensorModel:'TB564', sensorSerial:'-', model:'TB82', transmitterSerial:'-', calRange:'0 to 14', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "431QC410": { desc:'Chemical plant layer 3', location:'Pulp1', brand:'ABB', sensorModel:'TB564', sensorSerial:'-', model:'TB82', transmitterSerial:'-', calRange:'2.5 to 14', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "431QI042": { desc:'pH LC Tank', location:'Pulp1', brand:'ABB', sensorModel:'TB557', sensorSerial:'-', model:'TB82', transmitterSerial:'-', calRange:'0 to 14', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "431QI414": { desc:'DI - Tower pH', location:'Pulp1', brand:'ABB', sensorModel:'TB557', sensorSerial:'-', model:'TB82', transmitterSerial:'-', calRange:'0 to 8', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "623QT352": { desc:'Chemical plant layer 3', location:'Pulp1', brand:'ABB', sensorModel:'TB557', sensorSerial:'-', model:'TB82', transmitterSerial:'-', calRange:'0 to 14', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "422QE377": { desc:'EXTRACTION PDW.', location:'Pulp2', brand:'ABB', sensorModel:'TB557J3E0T15', sensorSerial:'3K610000250620', model:'TB82', transmitterSerial:'3K610000250585', calRange:'8 to 14', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "432QE049": { desc:'PULP TO DO', location:'Pulp2', brand:'ABB', sensorModel:'TBX57', sensorSerial:'3K430000023314', model:'AWT210', transmitterSerial:'3K220000882469', calRange:'2 to 7', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "432QE074": { desc:'WP. EXTRACTION', location:'Pulp2', brand:'ABB', sensorModel:'TB557', sensorSerial:'-', model:'TB82', transmitterSerial:'', calRange:'0 to 10', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "432QE134": { desc:'FILTRATE TO SCRUBBER', location:'Pulp2', brand:'ABB', sensorModel:'TB557', sensorSerial:'3K43000023313', model:'AWT210', transmitterSerial:'3K220000995563', calRange:'4 to 12', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "432QE139": { desc:'WP. EXTRACTION', location:'Pulp2', brand:'ABB', sensorModel:'-', sensorSerial:'-', model:'TB82', transmitterSerial:'-', calRange:'0 to 14', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "432QE170": { desc:'pH TO D1 TOWER', location:'Pulp2', brand:'ABB', sensorModel:'TB557', sensorSerial:'3K430000031446', model:'TB82', transmitterSerial:'-', calRange:'4 to 10', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "432QE175": { desc:'pH D1-MIXER', location:'Pulp2', brand:'-', sensorModel:'-', sensorSerial:'-', model:'-', transmitterSerial:'-', calRange:'', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "432QE202": { desc:'WP. EXTRACTION', location:'Pulp2', brand:'ABB', sensorModel:'TB557', sensorSerial:'3K430000017859', model:'TB82', transmitterSerial:'3K61000062123', calRange:'3 to 8', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "432QE274": { desc:'WP EXTRACTION', location:'Pulp2', brand:'ABB', sensorModel:'TB557J3EM1T15', sensorSerial:'3K610000260134', model:'TB82', transmitterSerial:'-', calRange:'2 to 7', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "432QE275": { desc:'ALKALINE EFF pH', location:'Pulp2', brand:'METTLER TOLEDO', sensorModel:'-', sensorSerial:'-', model:'M300', transmitterSerial:'-', calRange:'7 to 14', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "432QE337": { desc:'PULP TO PROT SCREEN', location:'Pulp2', brand:'METTLER TOLEDO', sensorModel:'-', sensorSerial:'3K610000244879', model:'M400', transmitterSerial:'-', calRange:'3 to 8', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "432QE347": { desc:'-', location:'Pulp2', brand:'ABB', sensorModel:'-', sensorSerial:'-', model:'TB82', transmitterSerial:'-', calRange:'', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "432QE401": { desc:'ALKALINE EFF pH', location:'Pulp2', brand:'-', sensorModel:'-', sensorSerial:'-', model:'-', transmitterSerial:'-', calRange:'', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "432QE815": { desc:'-', location:'Pulp2', brand:'-', sensorModel:'-', sensorSerial:'-', model:'-', transmitterSerial:'-', calRange:'', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "No Tag":    { desc:'', location:'Pulp2', brand:'ABB', sensorModel:'', sensorSerial:'', model:'AWT210', transmitterSerial:'3K220000989998', calRange:'0 to 14', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "484QT700": { desc:'Sumpit black liger QL', location:'', brand:'ABB', sensorModel:'TBX557', sensorSerial:'3K430000030332', model:'AWT210', transmitterSerial:'3K220000869003', calRange:'6 to 14', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
  "472QT104": { desc:'-', location:'LK2', brand:'ABB', sensorModel:'TB557', sensorSerial:'-', model:'AWT210', transmitterSerial:'-', calRange:'0 to 14', unit:'pH', acceptError:'0.5 pH (Error of Reading)' },
};

function loadLocalTagPresets(){ try{ return JSON.parse(localStorage.getItem('tag-presets')||'{}'); }catch(_){ return {}; } }
function saveLocalTagPresets(obj){ localStorage.setItem('tag-presets', JSON.stringify(obj)); }

function parseRange(str){
  if(!str) return {};
  const m = String(str).match(/(-?\d+(?:\.\d+)?)\s*to\s*(-?\d+(?:\.\d+)?)/i);
  if(!m) return {};
  return { rangeMin: +m[1], rangeMax: +m[2] };
}

async function maybeCreatePresetInteractively(tag){
  if(!tag) return null;
  if(!confirm('พบ Tag ใหม่ ต้องการสร้าง Preset ไหม?')) return null;
  const brand = prompt('Maker/Brand:', 'ABB') || '';
  const model = prompt('Transmitter/Analyzer Model:', 'TB82') || '';
  const calRange = prompt('ช่วงการทำงาน (เช่น 3 to 8):', '3 to 8') || '';
  const desc = prompt('Tag Name/Description:', '') || '';
  const sensorModel = prompt('Sensor Model (optional):', '') || '';
  const sensorSerial = prompt('Sensor Serial No. (optional):', '') || '';
  const transmitterSerial = prompt('Transmitter Serial No. (optional):', '') || '';
  const acceptError = prompt('Accept Error:', '0.5 pH (Absolute Error)') || '';
  const unit = prompt('Unit:', 'pH') || 'pH';
  const preset = { brand, model, calRange, desc, sensorModel, sensorSerial, transmitterSerial, acceptError, unit };
  const db = loadLocalTagPresets(); db[tag] = preset; saveLocalTagPresets(db);
  return preset;
}

async function applyTagPreset(tag){
  if(!tag || tag==='__new__') return;
  const localDB = loadLocalTagPresets();
  let preset = localDB[tag] || BUILTIN_TAG_PRESETS[tag] || null;
  if(!preset){ preset = await maybeCreatePresetInteractively(tag); if(!preset) return; }
  const { brand, model, calRange, desc, sensorModel, sensorSerial, transmitterSerial, acceptError, unit, location } = preset;
  if (brand) state.maker = brand;
  if (model) state.transmitterModel = model;
  if (desc) state.tagName = desc;
  if (location) state.location = location;
  if (sensorModel) state.sensorModel = sensorModel;
  if (sensorSerial) state.sensorSerial = sensorSerial;
  if (transmitterSerial) state.transmitterSerial = transmitterSerial;
  const rng = parseRange(calRange);
  if (rng.rangeMin != null) state.rangeMin = rng.rangeMin;
  if (rng.rangeMax != null) state.rangeMax = rng.rangeMax;
  state.unit = unit || state.unit || 'pH';
  if (acceptError) state.acceptError = acceptError;
  fillFormFromState(); saveToLocal(); updatePreview();
}

// -------------------------------
// DOM helpers
const $ = (sel, el = document) => el.querySelector(sel);
const $$ = (sel, el = document) => Array.from(el.querySelectorAll(sel));

function bindText(){ $$('[data-bind]').forEach(el => { const k = el.getAttribute('data-bind'); el.textContent = state[k] || ''; }); }

function setImg(id, dataUrl){
  const el = document.getElementById(id); if (!el) return;
  if (dataUrl) { el.src = dataUrl; el.style.visibility = 'visible'; }
  else { el.src = 'data:image/gif;base64,R0lGODlhAQABAAAAACw='; el.style.visibility = 'visible'; }
}

function renderBuffers(){
  const tbody = document.getElementById('buffer-rows'); tbody.innerHTML = '';
  for (const row of state.buffers) {
    const isEmpty = (!row) || ['equipment','brandModel','serial','expDate'].every(k => (row[k] == null || String(row[k]).trim() === ''));
    if (isEmpty) continue;
    const tr = document.createElement('tr');
    const fmt = (s) => formatDateDMY(s);
    const serial = (row.serial && String(row.serial).trim() !== '') ? row.serial : '-';
    tr.innerHTML = `<td>${row.equipment || ''}</td><td>${row.brandModel || ''}</td><td>${serial}</td><td>${fmt(row.expDate) || ''}</td>`;
    tbody.appendChild(tr);
  }
}

function renderCalPoints(){
  const tbLeft = document.getElementById('cal-left');
  const tbRight = document.getElementById('cal-right');
  const tbSingle = document.getElementById('cal-rows');

  // Build rows from fixed buffers 4,7,10 using form inputs if present; fallback to state.points
  const buffers = [4,7,10];
  const rows = buffers.map((b)=>{
    const uncal = String(state.calResult||'').toLowerCase().startsWith('un');
    const f = parseFloat(state['found'+b])
    const c = parseFloat(state['cal'+b])
    const hasF = !Number.isNaN(f);
    const hasC = !Number.isNaN(c) && !uncal;
    const fPh = hasF ? (f - b) : '';
    const fPct = hasF ? ((f - b) / b * 100) : '';
    const cPh = hasC ? (c - b) : (uncal ? '-' : '');
    const cPct = hasC ? ((c - b) / b * 100) : (uncal ? '-' : '');
    const fmt = (v) => (v === '' ? '' : (Math.round(v * 100) / 100).toFixed(2));
    return {
      buffer: b,
      foundReading: hasF ? f : '',
      foundPct: fmt(fPct),
      foundPh: fmt(fPh),
      calReading: hasC ? c : (uncal ? '-' : ''),
      calPct: (uncal ? '-' : fmt(cPct)),
      calPh: (uncal ? '-' : fmt(cPh)),
    };
  });

  if (tbLeft && tbRight) {
    tbLeft.innerHTML = ''; tbRight.innerHTML = '';
    for (const r of rows) {
      const trL = document.createElement('tr');
      trL.innerHTML = `<td>${r.buffer}</td><td>${r.foundReading}</td><td>${r.foundPct}</td><td>${r.foundPh}</td>`;
      tbLeft.appendChild(trL);
      const trR = document.createElement('tr');
      trR.innerHTML = `<td>${r.buffer}</td><td>${r.calReading}</td><td>${r.calPct}</td><td>${r.calPh}</td>`;
      tbRight.appendChild(trR);
    }
  } else if (tbSingle) {
    tbSingle.innerHTML = '';
    for (const r of rows) {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${r.buffer}</td><td>${r.foundReading}</td><td>${r.foundPct}</td><td>${r.foundPh}</td><td>${r.calReading}</td><td>${r.calPct}</td><td>${r.calPh}</td>`;
      tbSingle.appendChild(tr);
    }
  }
}

async function ensureDefaultLogo(){
  try{
    if (state.images && state.images.logo) return;
    const svgUrl = 'assets/logo_nps.svg';
    const resp = await fetch(svgUrl, { cache: 'no-cache' });
    if (!resp.ok) return;
    const svgText = await resp.text();
    const blob = new Blob([svgText], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const img = new Image();
    await new Promise((res, rej)=>{ img.onload = res; img.onerror = rej; img.src = url; });
    const size = Math.max(img.naturalWidth || 600, 600);
    const canvas = document.createElement('canvas');
    canvas.width = size; canvas.height = size; // square to keep quality
    const ctx = canvas.getContext('2d');
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    const dataUrl = canvas.toDataURL('image/png');
    URL.revokeObjectURL(url);
    state.images.logo = dataUrl; saveToLocal();
  }catch(_){ /* noop */ }
}

function updatePreview(){
  bindText(); renderBuffers(); renderCalPoints();
  // Ensure photo wrappers exist for overlay badges (like sample)
  (function ensurePhotoWrappers(){
    const wrap = document.querySelector('#certificate .photos');
    if (!wrap) return;
    const imgs = Array.from(wrap.querySelectorAll('img'));
    imgs.forEach(img => {
      if (img.parentElement && img.parentElement.classList && img.parentElement.classList.contains('photo')) return;
      const container = document.createElement('div'); container.className = 'photo';
      const badge = document.createElement('div'); badge.className = 'badge';
      const n = (img.id || '').replace('img',''); if (n) badge.id = 'badge'+n;
      // insert before and move img inside
      wrap.insertBefore(container, img);
      container.appendChild(img);
      container.appendChild(badge);
    });
  })();
  const calibBlock = document.getElementById('calibration-record');
  if (calibBlock) {
    const fromArray = Array.isArray(state.points) && state.points.some(p => p && Object.keys(p).some(k => p[k] !== '' && p[k] != null));
    const fromInputs = [4,7,10].some(b => {
      const f = state['found'+b]; const c = state['cal'+b];
      return (f != null && String(f) !== '') || (c != null && String(c) !== '');
    });
    calibBlock.style.display = (fromArray || fromInputs) ? '' : 'none';
  }
  if (state.images.logo) {
    $('#logo-box').style.background = 'none';
    $('#logo-box').style.border = 'none';
    $('#logo-box').style.padding = '0';
    $('#logo-box').style.width = '32mm';
    $('#logo-box').style.height = '26mm';
    $('#logo-box').innerHTML = `<img src="${state.images.logo}" alt="logo" style="width:100%;height:100%;object-fit:contain;display:block;"/>`;
  } else {
    // Fallback to SVG path if conversion not ready yet
    $('#logo-box').style.background = 'none';
    $('#logo-box').style.border = 'none';
    $('#logo-box').style.padding = '0';
    $('#logo-box').style.width = '32mm';
    $('#logo-box').style.height = '26mm';
    $('#logo-box').innerHTML = `<img src="assets/logo_nps.svg" alt="logo" style="width:100%;height:100%;object-fit:contain;display:block;"/>`;
  }
  setImg('img1', state.images.photo1);
  setImg('img2', state.images.photo2);
  setImg('img3', state.images.photo3);
  // Apply photo labels from filenames
  try{
    if (!state.photoLabels) state.photoLabels = { photo1: '', photo2: '', photo3: '' };
    const applyLabel = (n, key) => {
      const b = document.getElementById('badge'+n);
      if (!b) return;
      const txt = (state.photoLabels && state.photoLabels[key]) ? state.photoLabels[key] : '';
      if (txt) { b.textContent = txt; b.style.display = 'inline-block'; }
      else { b.textContent = ''; b.style.display = 'none'; }
    };
    applyLabel(1, 'photo1');
    applyLabel(2, 'photo2');
    applyLabel(3, 'photo3');
  }catch(_){}
  const photosBlock = document.getElementById('photos-block'); if (photosBlock) photosBlock.style.display = '';
  // Ensure Range separator uses hyphen instead of Thai word
  try{ const cert = document.getElementById('certificate'); if (cert) fixRangeSeparatorIn(cert); }catch(_){ }
}

// Build a preset object from current form/state fields
function buildPresetFromState(){
  const calRange = (state.rangeMin!=='' && state.rangeMin!=null && state.rangeMax!=='' && state.rangeMax!=null)
    ? `${state.rangeMin} to ${state.rangeMax}` : '';
  return {
    desc: state.tagName || '',
    location: state.location || '',
    brand: state.maker || '',
    sensorModel: state.sensorModel || '',
    sensorSerial: state.sensorSerial || '',
    model: state.transmitterModel || '',
    transmitterSerial: state.transmitterSerial || '',
    calRange,
    unit: state.unit || 'pH',
    acceptError: state.acceptError || ''
  };
}

function ensureTagOptionInSelect(tag){
  try{
    const form = document.getElementById('data-form');
    const sel = form && form.elements['tagNo'];
    if (!sel) return;
    if (![...sel.options].some(o => o.value === tag)) {
      const opt = document.createElement('option'); opt.value = tag; opt.textContent = tag; sel.appendChild(opt);
    }
  }catch(_){ }
}

function saveCurrentTagPreset(){
  const tag = (state.tagNo || '').trim();
  if (!tag) return false;
  const preset = buildPresetFromState();
  const db = loadLocalTagPresets(); db[tag] = preset; saveLocalTagPresets(db);
  ensureTagOptionInSelect(tag);
  return true;
}

function enhanceTagSelect(){
  try{
    const sel = document.querySelector('select[name="tagNo"]');
    if (!sel) return;
    // Insert "+ new" option after the first empty option
    if (![...sel.options].some(o => o.value === '__new__')){
      const opt = document.createElement('option'); opt.value='__new__'; opt.textContent='+ เพิ่มแท็กใหม่…';
      const after = sel.querySelector('option[value=""]');
      if (after && after.nextSibling) sel.insertBefore(opt, after.nextSibling); else sel.appendChild(opt);
    }
    if (sel.dataset.enhanced === '1') return; sel.dataset.enhanced = '1';
    sel.addEventListener('change', async (e)=>{
      if (sel.value === '__new__'){
        e.preventDefault(); e.stopImmediatePropagation?.();
        const code = prompt('กรอก Tag No. ใหม่:', '');
        if (!code){ sel.value=''; return; }
        state.tagNo = code.trim(); saveCurrentTagPreset(); ensureTagOptionInSelect(state.tagNo); sel.value = state.tagNo; saveToLocal(); updatePreview();
        alert('เพิ่มแท็กใหม่และบันทึก preset แล้ว');
      }
    });
  }catch(_){ }
}

function fixRangeSeparatorIn(root){
  try{
    const row = Array.from(root.querySelectorAll('.kv tr')).find(tr => tr.querySelector('th')?.textContent.trim() === 'Range');
    if (row){ const td = row.querySelector('td'); if (td) td.innerHTML = '<span data-bind="rangeMin"></span> - <span data-bind="rangeMax"></span> <span data-bind="unit"></span>'; }
  }catch(_){/* noop */}
}

async function openPreviewModal(){
  updatePreview();
  const modal = document.getElementById('preview-modal');
  const mount = document.getElementById('preview-mount');
  const base = document.getElementById('certificate');
  try{ await ensureDefaultLogo(); }catch(_){ }
  if (!modal || !mount || !base) return;
  const clone = base.cloneNode(true);
  clone.id = 'certificate-preview';
  clone.classList.remove('pdf');
  mount.innerHTML = '';
  mount.appendChild(clone);
  // Fix range separator inside preview copy too
  try{ fixRangeSeparatorIn(clone); }catch(_){ }
  modal.classList.remove('hidden');
  document.body.style.overflow = 'hidden';
  const close = () => { modal.classList.add('hidden'); mount.innerHTML = ''; document.body.style.overflow = ''; };
  document.getElementById('preview-close')?.addEventListener('click', close, { once:true });
  document.getElementById('preview-close-2')?.addEventListener('click', close, { once:true });
  modal.querySelector('[data-close]')?.addEventListener('click', close, { once:true });
  const onKey = (e)=>{ if(e.key==='Escape'){ close(); window.removeEventListener('keydown', onKey); } };
  window.addEventListener('keydown', onKey);
}

function saveToLocal(){ localStorage.setItem('ph-cert-state', JSON.stringify(state)); }
function loadFromLocal(){ try{ const s = localStorage.getItem('ph-cert-state'); if (!s) return false; Object.assign(state, JSON.parse(s)); return true; } catch (_){ return false; } }
function toJSON(){ return JSON.stringify(state, null, 2); }
function download(filename, text){ const blob = new Blob([text], {type: 'application/json'}); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = filename; a.click(); URL.revokeObjectURL(url); }

// -------------------------------
// Form helpers
function ensureDefaultBuffers(){
  if (!state.buffers || state.buffers.length === 0) {
    state.buffers = [
      { equipment: 'pH buffer solution (pH4)',  brandModel: 'METTLER TOLEDO', serial: '', expDate: '' },
      { equipment: 'pH buffer solution (pH7)',  brandModel: 'METTLER TOLEDO', serial: '', expDate: '' },
      { equipment: 'pH buffer solution (pH10)', brandModel: 'METTLER TOLEDO', serial: '', expDate: '' },
    ];
  }
}

function normalizeBufferRows(){
  ensureDefaultBuffers();
  const isEmpty = (row) => !row || ['equipment','brandModel','serial','expDate'].every(k => (row[k] == null || String(row[k]).trim() === ''));
  let lastUsed = -1;
  for (let i = state.buffers.length - 1; i >= 0; i--) { const r = state.buffers[i]; if (!(r == null || isEmpty(r))) { lastUsed = i; break; } }
  const minRows = 3;
  const keep = Math.max(minRows - 1, lastUsed);
  const newLen = Math.max(minRows, keep + 1);
  while (state.buffers.length < newLen) state.buffers.push({ equipment: '', brandModel: 'METTLER TOLEDO', serial: '', expDate: '' });
  if (state.buffers.length > newLen) state.buffers = state.buffers.slice(0, newLen);
}
function ensureDefaultPoints(){ if (!Array.isArray(state.points)) state.points = []; }

function rebuildEditableTables(){
  normalizeBufferRows();
  const bTbody = $('#buffer-table tbody'); bTbody.innerHTML = '';
  state.buffers.forEach((row, i) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><input data-k="equipment"  data-i="${i}" value="${row.equipment ?? ''}"></td>
      <td><input data-k="brandModel" data-i="${i}" value="${row.brandModel ?? ''}"></td>
      <td><input data-k="serial"     data-i="${i}" value="${row.serial ?? ''}"></td>
      <td><input type="date" data-k="expDate" data-i="${i}" value="${row.expDate ?? ''}"></td>`;
    bTbody.appendChild(tr);
  });

  const cTbody = $('#cal-table tbody');
  if (cTbody) {
    cTbody.innerHTML = '';
    state.points.forEach((p, i) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><input data-k="buffer"       data-i="${i}" type="number" step="any" value="${p.buffer ?? ''}"></td>
        <td><input data-k="foundReading" data-i="${i}" type="number" step="any" value="${p.foundReading ?? ''}"></td>
        <td><input data-k="foundAbs"     data-i="${i}" type="number" step="any" value="${p.foundAbs ?? ''}"></td>
        <td><input data-k="foundAcc"     data-i="${i}" value="${p.foundAcc ?? ''}"></td>
        <td><input data-k="calReading"   data-i="${i}" type="number" step="any" value="${p.calReading ?? ''}"></td>
        <td><input data-k="calAbs"       data-i="${i}" type="number" step="any" value="${p.calAbs ?? ''}"></td>
        <td><input data-k="calAcc"       data-i="${i}" value="${p.calAcc ?? ''}"></td>`;
      cTbody.appendChild(tr);
    });
  }
}

function wireEditableHandlers(){
  $('#buffer-table tbody').addEventListener('input', (e) => {
    const t = e.target; if (!t.matches('input')) return;
    const i = +t.getAttribute('data-i'); const k = t.getAttribute('data-k');
    state.buffers[i][k] = t.value; normalizeBufferRows(); rebuildEditableTables(); saveToLocal(); updatePreview();
  });
  const calTBodyEl = $('#cal-table tbody');
  if (calTBodyEl) {
    calTBodyEl.addEventListener('input', (e) => {
      const t = e.target; if (!t.matches('input')) return;
      const i = +t.getAttribute('data-i'); const k = t.getAttribute('data-k');
      const val = t.type === 'number' ? (t.value === '' ? '' : +t.value) : t.value;
      state.points[i][k] = val; saveToLocal(); updatePreview();
    });
  }
}

function wireFormTabs(){
  const form = document.getElementById('data-form');
  const tabbar = document.querySelector('.form-tabs'); if (!form || !tabbar) return;
  tabbar.addEventListener('click', (e)=>{ const btn = e.target.closest('[data-jump]'); if (!btn) return; const sel = btn.getAttribute('data-jump'); const target = document.querySelector(sel); if (target) target.scrollIntoView({behavior:'smooth'}); });
  const tabs = Array.from(tabbar.querySelectorAll('[data-jump]'));
  const observer = new IntersectionObserver((entries)=>{ for (const en of entries){ if (en.isIntersecting){ const id = '#' + en.target.id; tabs.forEach(b=>b.classList.toggle('active', b.getAttribute('data-jump')===id)); } } },{root: form, threshold: 0.5});
  ['doc-section','xmtr-section','ref-section','cal-section','sign-section','photos-section'].forEach(id=>{ const t=document.getElementById(id); if(t) observer.observe(t); });
}

function wireFormBasics(){
  const form = document.getElementById('data-form');
  function onFieldChange(e){
    const el = e.target; if (!el.name) return;
    if (['logo','photo1','photo2','photo3'].includes(el.name)) return;
    state[el.name] = el.value;
    if (el.name === 'calibrateDate') {
      try { setBufferExpDatesFromBase(state.calibrateDate, true); rebuildEditableTables(); } catch(_){}
    }
    if (el.name === 'calResult') {
      try { updateCalInputsAvailability(); } catch(_){}
    }
    saveToLocal(); updatePreview();
  }
  form.addEventListener('input', onFieldChange);
  form.addEventListener('change', onFieldChange);

  const tagEl = form.elements['tagNo'];
  try{ const certEl = form.elements['certificateNo']; if (certEl){ certEl.readOnly = false; certEl.disabled = false; } }catch(_){ }
  if (tagEl) {
    // Populate Tag No. dropdown from presets
    const localDB = loadLocalTagPresets();
    const keys = Array.from(new Set(Object.keys(BUILTIN_TAG_PRESETS).concat(Object.keys(localDB)))).sort((a,b)=>String(a).localeCompare(String(b)));
    tagEl.innerHTML = '<option value="">เลือก Tag...</option>' + keys.map(k => `<option value="${k}">${k}</option>`).join('');
    tagEl.addEventListener('change', () => applyTagPreset(tagEl.value));
    if (state.tagNo) { if (![...tagEl.options].some(o=>o.value===state.tagNo)){ const opt=document.createElement('option'); opt.value=state.tagNo; opt.textContent=state.tagNo; tagEl.appendChild(opt); } tagEl.value = state.tagNo; }
  }

  // Populate Calibrated by dropdown
  const techSel = form.elements['technician'];
  if (techSel && techSel.tagName === 'SELECT'){
    techSel.innerHTML = '<option value="">เลือกผู้สอบเทียบ...</option>' + CALIBRATORS.map(function(c){ return '<option value="' + c.name + '" data-pos="' + c.position + '">' + c.name + '</option>'; }).join('');
    techSel.addEventListener('change', ()=>{
      const opt = techSel.selectedOptions[0];
      state.technician = techSel.value;
      state.technicianPosition = opt ? (opt.getAttribute('data-pos') || '') : '';
      const inline = document.getElementById('tech-role-inline');
      if (inline) inline.textContent = state.technicianPosition ? `(${state.technicianPosition})` : '';
      saveToLocal(); updatePreview();
    });
  }

  function handleFileInput(input, key){
    if (!input) return;
    input.addEventListener('change', async () => {
      const f = input.files && input.files[0];
      if (!f) { state.images[key] = null; if (key.startsWith('photo')) { if (!state.photoLabels) state.photoLabels = { photo1:'',photo2:'',photo3:'' }; state.photoLabels[key] = ''; } saveToLocal(); updatePreview(); return; }
      const dataUrl = await fileToDataURL(f); state.images[key] = dataUrl;
      if (key.startsWith('photo')){
        if (!state.photoLabels) state.photoLabels = { photo1:'',photo2:'',photo3:'' };
        const name = (f.name || '').replace(/^.*[\\\/]/,'');
        const base = name.replace(/\.[^.]+$/,'');
        state.photoLabels[key] = base;
      }
      saveToLocal(); updatePreview();
      // update inline form thumbnail
      try{
        const map = { photo1: 'thumb-photo1', photo2: 'thumb-photo2', photo3: 'thumb-photo3' };
        const id = map[key]; if (id){ const img = document.getElementById(id); if (img){ img.src = dataUrl; img.style.display='block'; } }
      }catch(_){ }
    });
  }
  handleFileInput(form.elements['logo'], 'logo');
  handleFileInput(form.elements['photo1'], 'photo1');
  handleFileInput(form.elements['photo2'], 'photo2');
  handleFileInput(form.elements['photo3'], 'photo3');

  const addBufferBtn = $('#btn-add-buffer'); if (addBufferBtn) addBufferBtn.addEventListener('click', () => { state.buffers.push({ equipment: '', brandModel: 'METTLER TOLEDO', serial: '', expDate: '' }); rebuildEditableTables(); saveToLocal(); updatePreview(); });
  const addPointBtn  = $('#btn-add-point');  if (addPointBtn)  addPointBtn.addEventListener('click', () => { state.points.push({ buffer: '', foundReading: '', foundAbs: '', foundAcc: '', calReading: '', calAbs: '', calAcc: '' }); rebuildEditableTables(); saveToLocal(); updatePreview(); });

  $('#btn-save-json').addEventListener('click', () => { download(makeBaseFilename() + '.json', toJSON()); });
  $('#input-load-json').addEventListener('change', async (e) => {
    const f = e.target.files && e.target.files[0]; if (!f) return;
    const text = await f.text();
    try { const obj = JSON.parse(text); Object.assign(state, obj); fillFormFromState(); rebuildEditableTables(); saveToLocal(); updatePreview(); }
    catch { alert('ไฟล์ JSON ไม่ถูกต้อง'); }
  });
  $('#btn-load-demo').addEventListener('click', () => { loadDemo(); fillFormFromState(); rebuildEditableTables(); saveToLocal(); updatePreview(); });

  wireFormTabs();
  // Initial sync of technician UI/role
  syncTechnicianUI();

  // Update current Tag preset with values from fields
  const btnUpdate = document.getElementById('btn-update-tag');
  if (btnUpdate) btnUpdate.addEventListener('click', () => {
    const tag = (state.tagNo || '').trim();
    if (!tag) { alert('โปรดเลือก Tag No. ก่อน'); return; }
    // Build calRange string similar to presets
    const min = (state.rangeMin === undefined || state.rangeMin === null || String(state.rangeMin) === '') ? '' : String(state.rangeMin);
    const max = (state.rangeMax === undefined || state.rangeMax === null || String(state.rangeMax) === '') ? '' : String(state.rangeMax);
    const calRange = (min !== '' && max !== '') ? `${min} to ${max}` : '';
    const preset = {
      desc: state.tagName || '',
      location: state.location || '',
      brand: state.maker || '',
      sensorModel: state.sensorModel || '',
      sensorSerial: state.sensorSerial || '',
      model: state.transmitterModel || '',
      transmitterSerial: state.transmitterSerial || '',
      calRange,
      unit: state.unit || 'pH',
      acceptError: state.acceptError || ''
    };
    const db = loadLocalTagPresets(); db[tag] = preset; saveLocalTagPresets(db);
    // Ensure dropdown includes this tag
    const tagEl = form.elements['tagNo'];
    if (tagEl && ![...tagEl.options].some(o=>o.value===tag)) { const opt=document.createElement('option'); opt.value=tag; opt.textContent=tag; tagEl.appendChild(opt); }
    showToast('อัปเดตข้อมูล Tag สำเร็จ', 'success');
  });
}

function fileToDataURL(file){ return new Promise((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(reader.result); reader.onerror = reject; reader.readAsDataURL(file); }); }
function fillFormFromState(){ const form = document.getElementById('data-form'); $$('input[name], select[name], textarea[name]', form).forEach(el => { if (['logo','photo1','photo2','photo3'].includes(el.name)) return; if (el.tagName === 'SELECT') { const val = state[el.name] ?? ''; el.value = val; if (val && el.value !== val) { const opt = document.createElement('option'); opt.value = val; opt.textContent = val; el.appendChild(opt); el.value = val; } } else if (el.type === 'date' || el.type === 'number' || el.tagName === 'TEXTAREA' || el.type === 'text') { el.value = state[el.name] ?? ''; } }); }

function updateCalInputsAvailability(){
  const uncal = String(state.calResult||'').toLowerCase().startsWith('un');
  const phs = {cal4:'เช่น 4.01', cal7:'เช่น 7.03', cal10:'เช่น 10.01'};
  ['cal4','cal7','cal10'].forEach(name=>{
    const el = document.querySelector(`[name="${name}"]`);
    if (!el) return;
    el.disabled = uncal;
    if (uncal){ el.value=''; el.placeholder='-'; }
    else { el.placeholder = phs[name] || ''; }
  });
}

function makeBaseFilename(){ const tag = (state.tagNo || 'TAG').replace(/[^\w\-]+/g, '_'); const date = state.calibrateDate || new Date().toISOString().slice(0,10); return `Calibration_${tag}_${date}`; }
// รอให้ <img> ทั้งหมดในเอกสาร (หรือในโหนดที่ส่งเข้า) โหลดเสร็จ
async function waitImages(el){
  const imgs = Array.from(el.querySelectorAll('img'));
  await Promise.all(imgs.map(img => {
    if (img.complete && img.naturalWidth) return;
    return new Promise(res => { img.onload = img.onerror = () => res(); });
  }));
  // บราวเซอร์ใหม่ๆ จะเดโค้ดภาพให้คมชัดก่อนแคปเจอร์
  await Promise.all(imgs.map(img => img.decode?.().catch(()=>{})));
}

// -------------------------------
// Certificate number helpers: AC/YYYY/NNNN
function certSeqStorageKey(year){ return 'ph-cert-seq-' + year; }
function formatCertNo(year, num){ return `AC/${year}/${String(num).padStart(4,'0')}`; }
function parseCertNo(str){ const m = /^AC[\/\- ]?(\d{4})[\/\- ]?(\d{3,})$/i.exec(String(str||'').trim()); return m ? { year: Number(m[1]), num: Number(m[2]) } : null; }
function getNextSeq(year){ let base = 149; try{ const v = localStorage.getItem(certSeqStorageKey(year)); if (v!=null && v!=='' && !isNaN(+v)) base = Number(v); }catch(_){ } return base + 1; }
function markSeqUsed(year, num){ try{ const key = certSeqStorageKey(year); const cur = Number(localStorage.getItem(key) || '149'); if (num > cur) localStorage.setItem(key, String(num)); }catch(_){ } }
function ensureCertificateNo(){
  const y = new Date().getFullYear();
  const p = parseCertNo(state.certificateNo);
  const next = getNextSeq(y);
  if (p && p.year === y) {
    // If existing number is below base 150, upgrade to current sequence
    const num = p.num < 150 ? next : Math.max(p.num, next);
    state.certificateNo = formatCertNo(y, num);
    try{ const form=document.getElementById('data-form'); const el=form && form.elements['certificateNo']; if(el) el.value = state.certificateNo; }catch(_){ }
    return;
  }
  state.certificateNo = formatCertNo(y, next);
  try{ const form=document.getElementById('data-form'); const el=form && form.elements['certificateNo']; if(el) el.value = state.certificateNo; }catch(_){ }
}
function bumpCertificateNoAfterSave(){ const y = new Date().getFullYear(); const p = parseCertNo(state.certificateNo); const used = (p && p.year===y) ? p.num : getNextSeq(y); markSeqUsed(y, used); const next = used + 1; state.certificateNo = formatCertNo(y, next); saveToLocal(); try{ const form=document.getElementById('data-form'); const el=form && form.elements['certificateNo']; if(el) el.value = state.certificateNo; }catch(_){ } updatePreview(); }

// -------------------------------
// PDF + Upload
async function generatePDF(){
  const base = document.getElementById('certificate');
  const filename = makeDriveFilename() + '.pdf';
  const sandbox = document.createElement('div');
  // Keep the print clone at the viewport origin and invisible.
  // Using huge negative offsets can make html2canvas include blank space
  // above the element. Opacity keeps it out of sight but renderable.
  sandbox.id = 'pdf-sandbox';
  sandbox.style.cssText = 'position:fixed; left:0; top:0; width:210mm; z-index:-1; opacity:0; pointer-events:none;';
  const el = base.cloneNode(true); el.id = 'certificate-print';
  document.body.appendChild(sandbox); sandbox.appendChild(el);
  el.classList.add('pdf');
  await waitImages(el);
  try{ if (document.fonts && document.fonts.ready) { await document.fonts.ready; } }catch(_){ }
  const root = document.documentElement;
  const pxPerMm = 96/25.4;
  // Target height = page height minus top/bottom paddings used in .paper.pdf
  // Keep these in sync with CSS (padding-top 5mm, padding-bottom 10mm)
  const topPadMm = 5.0, bottomPadMm = 10.0;
  const targetH = (297 - topPadMm - bottomPadMm) * pxPerMm;
 // กันเหนียวไม่ให้ล้นจนเกิดหน้าว่าง

  const minScale = 0.80, maxScale = 1.05, upperBound = 1.20;
  const measureAt = async (s) => { root.style.setProperty('--pdf-scale', String(s)); await new Promise(r => requestAnimationFrame(r)); return el.getBoundingClientRect().height; };
  let lo = minScale, hi = maxScale;
  const current = parseFloat(getComputedStyle(root).getPropertyValue('--pdf-scale')) || 0.955;
  lo = Math.max(lo, Math.min(current, hi));
  let hAtHi = await measureAt(hi);
  while (hAtHi < targetH && hi < upperBound) { const nextHi = Math.min(upperBound, hi * 1.06); if (Math.abs(nextHi - hi) < 0.0008) break; hi = nextHi; hAtHi = await measureAt(hi); }
  if (hAtHi > targetH){ for (let i=0;i<10;i++){ const mid = (lo + hi) / 2; const h = await measureAt(mid); if (h > targetH) hi = mid; else lo = mid; if (Math.abs(hi - lo) < 0.0008) break; } } else { lo = hi; }
  const scale = lo * 0.990;
root.style.setProperty('--pdf-scale', String(scale));
  const opt = {
    margin: [0,0,0,0],
    filename,
    image: { type: 'jpeg', quality: 0.98 },
    // Force capture from the page origin to avoid vertical offset
    html2canvas: { scale: 2, useCORS: true, letterRendering: true, backgroundColor: '#ffffff', scrollX: 0, scrollY: 0 },
    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
    // Allow page to naturally break; avoid-all can clip bottom content
    pagebreak: { mode: ['css','legacy'] }
  };
  await new Promise(r => setTimeout(r, 60));
  try {
    // Save current Tag preset before exporting and show progress
    try{ saveCurrentTagPreset(); }catch(_){ }
    showBusy('กำลังสร้างไฟล์ PDF...');
    let pdfBlob = null;
    await window.html2pdf()
      .from(el)
      .set(opt)
      .toPdf()
      .get('pdf')
      .then(pdf => {
        try {
          // Keep multi-page output to avoid cutting bottom sections (e.g. End of Calibration)
          // If a truly blank trailing page appears, we can enhance later to detect and remove only that page.
          const total = typeof pdf.getNumberOfPages === 'function'
            ? pdf.getNumberOfPages()
            : (pdf.internal && pdf.internal.getNumberOfPages ? pdf.internal.getNumberOfPages() : 1);
          try{ const total = typeof pdf.getNumberOfPages === " function\ ? pdf.getNumberOfPages() : (pdf.internal && pdf.internal.getNumberOfPages ? pdf.internal.getNumberOfPages() : 1); for (let i = total; i >= 2; i--) { pdf.deletePage(i); } }catch(_){}
        } catch (_) { /* noop */ }
        pdfBlob = pdf.output('blob');
      });

    if (GAS_URL) {
      showBusy('กำลังส่งไฟล์ไปยังเซิร์ฟเวอร์...');
      const result = await uploadViaAppsScript(pdfBlob, filename);
      const url = (result && (result.url || result.webViewLink || (result.fileId ? ('https://drive.google.com/file/d/' + result.fileId + '/view') : ''))) || '';
      bumpCertificateNoAfterSave();
      showToast('บันทึกสำเร็จ: ' + filename, 'success', url || undefined);
    } else {
      showBusy('กำลังอัปโหลดขึ้น Google Drive...');
      const res = await uploadToDriveQuiet(pdfBlob, filename);
      const url = res && (res.webViewLink || (res.id ? ('https://drive.google.com/file/d/' + res.id + '/view') : ''));
      bumpCertificateNoAfterSave();
      showToast('อัปโหลดสำเร็จ: ' + filename, 'success', url || undefined);
    }
  } catch (err) {
    console.error('PDF generate/upload error', err);
    const msg = (err && err.message) ? err.message : String(err);
    showToast('บันทึกไม่สำเร็จ: ' + msg, 'error');
  } finally {
    hideBusy();
    el.classList.remove('pdf'); sandbox.remove(); root.style.setProperty('--pdf-scale', '0.955');
  }
}
function makeDriveFilename(){ const tag = (state.tagNo || 'TAG').replace(/[^\w\-]+/g, '_'); const date = state.calibrateDate || new Date().toISOString().slice(0,10); return `${date}_${tag}`; }

// OAuth path (เผื่อวันหน้าจะเปิดใช้)
async function uploadToDriveIfConfigured(blob, filename){
  try{
    if(!blob) return;
    if(!DRIVE_FOLDER_ID || !UPLOAD_TO_DRIVE) return;
    if(!GOOGLE_CLIENT_ID){ alert('ยังไม่ได้ตั้งค่า Google OAuth Client ID'); return; }
    const token = await getGoogleAccessToken();
    await uploadToDrive(blob, filename, token);
    alert('อัปโหลดไปยัง Google Drive แล้ว\n' + filename);
  }catch(err){
    console.error('Drive upload failed', err);
    alert('อัปโหลดไปยัง Google Drive ไม่สำเร็จ\nรายละเอียด: ' + (err && err.message ? err.message : err));
  }
}
function loadGIS(){ return new Promise((resolve)=>{ if(window.google && google.accounts && google.accounts.oauth2){ resolve(); return; } const iv = setInterval(()=>{ if(window.google && google.accounts && google.accounts.oauth2){ clearInterval(iv); resolve(); } }, 100); }); }
async function getGoogleAccessToken(){
  await loadGIS();
  return new Promise((resolve, reject)=>{
    if(googleAccessToken){ resolve(googleAccessToken); return; }
    if(!googleTokenClient){
      googleTokenClient = google.accounts.oauth2.initTokenClient({
        client_id: GOOGLE_CLIENT_ID,
        scope: 'https://www.googleapis.com/auth/drive.file',
        callback: (resp)=>{ if(resp && resp.access_token){ googleAccessToken = resp.access_token; resolve(googleAccessToken); } else reject(new Error('no access token')); }
      });
    }
    const authed = (typeof localStorage!=='undefined' && localStorage.getItem('gdrive_authed')==='1');
    googleTokenClient.requestAccessToken({ prompt: authed ? '' : 'consent' });
    try{ localStorage.setItem('gdrive_authed','1'); }catch(_){ }
  });
}
async function uploadToDrive(blob, filename, token){
  const metadata = { name: filename, parents: [DRIVE_FOLDER_ID], mimeType: 'application/pdf' };
  const boundary = 'foo_bar_' + Math.random().toString(36).slice(2);
  const body = new Blob([
    `--${boundary}\r\n` + 'Content-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(metadata) + `\r\n` +
    `--${boundary}\r\n` + 'Content-Type: application/pdf\r\n\r\n',
    blob,
    `\r\n--${boundary}--`
  ], { type: 'multipart/related; boundary=' + boundary });

  const res = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',{
    method:'POST',
    headers:{ Authorization: 'Bearer ' + token, 'Content-Type': 'multipart/related; boundary=' + boundary },
    body
  });
  if(!res.ok){ let txt = ''; try{ txt = await res.text(); }catch(_){ } throw new Error('Drive upload HTTP ' + res.status + ' ' + txt); }
  return res.json();
}

// Quiet wrapper for Drive upload (no alerts)
async function uploadToDriveQuiet(blob, filename){
  if(!blob) return null;
  if(!DRIVE_FOLDER_ID || !UPLOAD_TO_DRIVE) return null;
  if(!GOOGLE_CLIENT_ID){ throw new Error('ยังไม่ได้ตั้งค่า Google OAuth Client ID'); }
  const token = await getGoogleAccessToken();
  return uploadToDrive(blob, filename, token);
}

// Upload via Apps Script Web App
async function uploadViaAppsScript(blob, filename){
  if(!GAS_URL || !/^https?:\/\//i.test(GAS_URL)){
    throw new Error('ยังไม่ได้ตั้งค่า GAS_URL เป็นลิงก์ Web App (ลงท้ายด้วย /exec). ปัจจุบันค่าเป็น: ' + (GAS_URL||'(ว่าง)'));
  }
  const base64 = await new Promise((res, rej)=>{ const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1]); r.onerror = rej; r.readAsDataURL(blob); });
  const resp = await fetch(GAS_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ secret: GAS_SECRET || undefined, folderId: GAS_FOLDER_ID || undefined, filename, mimeType: 'application/pdf', fileBase64: base64 })
  });
  const text = await resp.text();
  let json = {}; try{ json = JSON.parse(text); }catch(_){}
  if(!resp.ok || (json && json.success===false)){
    const msg = (json && (json.error||json.message)) || text || ('HTTP '+resp.status);
    throw new Error('Apps Script upload failed: ' + msg);
  }
  return json;
}

// -------------------------------
// Demo & init
function loadDemo(){
  const y = new Date().getFullYear();
  const next = getNextSeq(y);
  Object.assign(state, {
    certificateNo: formatCertNo(y, next),
    calibrateDate: new Date().toISOString().slice(0,10),
    tagNo: '4B407T00',
    tagName: 'Sumpit black liger QL',
    location: 'Line 3',
    maker: 'ABB',
    sensorModel: 'TBX557',
    sensorSerial: '3K320000033232',
    transmitterModel: 'AWT210',
    transmitterSerial: '3K220000006903',
    rangeMin: 6, rangeMax: 14, unit: 'pH',
    acceptError: '0.5 pH (Error of Reading)',
    solutionType: 'Buffer',
    slope: 54.06, offset: -37.39,
    calResult: 'Calibrated',
    technician: 'Phutthiphong B.', approver: 'Chanin T.',
    signDate: new Date().toISOString().slice(0,10), remarks: ''
  });
  ensureDefaultBuffers();
  state.buffers[0].brandModel = 'METTLER TOLEDO';
  state.buffers[1].brandModel = 'METTLER TOLEDO';
  state.buffers[2].brandModel = 'METTLER TOLEDO';
  const d = new Date();
  const fmt = (dt) => dt.toISOString().slice(0,10);
  state.buffers[0].expDate = fmt(new Date(d.getFullYear()+1, d.getMonth(), d.getDate()));
  state.buffers[1].expDate = fmt(new Date(d.getFullYear()+1, d.getMonth(), d.getDate()+2));
  state.buffers[2].expDate = fmt(new Date(d.getFullYear()+1, d.getMonth(), d.getDate()+4));
  ensureDefaultPoints();
  state.points = [
    { buffer: 4,  foundReading: 6.27, foundAbs: 2.27, foundAcc: '0.7',  calReading: 4.01,  calAbs: 0.01, calAcc: '0.1' },
    { buffer: 7,  foundReading: 7.57, foundAbs: 0.57, foundAcc: '0.7',  calReading: 7.01,  calAbs: 0.01, calAcc: '0.1' },
    { buffer: 10, foundReading: 9.73, foundAbs: 0.27, foundAcc: '0.7',  calReading: 10.01, calAbs: 0.01, calAcc: '0.1' },
  ];
  // also fill new fixed-buffer inputs
  state.found4 = 6.27; state.found7 = 7.57; state.found10 = 9.73;
  state.cal4 = 4.01;   state.cal7 = 7.01;   state.cal10 = 10.01;
}

async function init(){
  const landing = document.getElementById('landing');
  const appView = document.getElementById('app-view');
  function showLanding(){ landing.classList.remove('hidden'); appView.classList.add('hidden'); }
  function dockActionsIntoForm(){
    try{
      const actions = document.querySelector('.app-header .actions');
      const mount = document.getElementById('form-actions');
      if(actions && mount && actions.parentElement !== mount){
        mount.appendChild(actions);
        actions.classList.add('docked');
      }
    }catch(_){/* noop */}
  }
  function showApp(){ landing.classList.add('hidden'); appView.classList.remove('hidden'); dockActionsIntoForm(); }
  if (location.hash === '#ph') { showApp(); } else { showLanding(); }

  const hadLocal = loadFromLocal();
  ensureDefaultBuffers(); ensureDefaultPoints();
  rebuildEditableTables(); wireEditableHandlers(); wireFormBasics();
  enhanceTagSelect();
  ensureApproverSelect();
  // Ensure certificate no default for current year starting at 0150
  try{ ensureCertificateNo(); }catch(_){ }
  // Always reflect current state into form so defaults (e.g., Solution Type) appear.
  try{ fillFormFromState(); syncTechnicianUI(); }catch(_){ }
  // Ensure Solution Type shows a default if empty
  try{
    const solEl = document.querySelector('[name="solutionType"]');
    if (solEl && (!solEl.value || String(solEl.value).trim() === '')){
      solEl.value = state.solutionType || 'Buffer';
      state.solutionType = solEl.value; saveToLocal();
    }
  }catch(_){ }
  // Ensure Brand&Model defaults for buffer table rows where empty
  try{
    ensureDefaultBuffers();
    state.buffers = (state.buffers || []).map(r => {
      if (!r) return { equipment:'', brandModel:'METTLER TOLEDO', serial:'', expDate:'' };
      const bm = (r.brandModel == null || String(r.brandModel).trim() === '') ? 'METTLER TOLEDO' : r.brandModel;
      return Object.assign({}, r, { brandModel: bm });
    });
    rebuildEditableTables(); saveToLocal();
  }catch(_){ }
  // If buffer Exp. Date missing, set to 1 year after calibrate date (or today)
  try{ setBufferExpDatesFromBase(state.calibrateDate || new Date().toISOString().slice(0,10), false); rebuildEditableTables(); }catch(_){ }
  // Ensure default logo is embedded as PNG to avoid missing in PDF
  try{ await ensureDefaultLogo(); }catch(_){ }
  try{ updateCalInputsAvailability(); }catch(_){ }
  updatePreview();

  $('#btn-generate').addEventListener('click', generatePDF);
  const btnPreview = document.getElementById('btn-preview'); if (btnPreview) btnPreview.addEventListener('click', openPreviewModal);
  const ctaPreview = document.getElementById('cta-preview'); const ctaSave = document.getElementById('cta-save'); const ctaHome = document.getElementById('cta-home');
  if (ctaPreview) ctaPreview.addEventListener('click', openPreviewModal);
  if (ctaSave) ctaSave.addEventListener('click', generatePDF);
  if (ctaHome) ctaHome.addEventListener('click', () => { showLanding(); location.hash = ''; });
  const choosePh = document.getElementById('choose-ph'); if (choosePh) choosePh.addEventListener('click', () => { showApp(); location.hash = '#ph'; });
  const backHome = document.getElementById('back-home'); if (backHome) backHome.addEventListener('click', () => { showLanding(); location.hash = ''; });
}
document.addEventListener('DOMContentLoaded', init);

// Utilities
function formatDateDMY(val){
  if(!val) return '';
  try{
    let d;
    if(val instanceof Date){ d = val; }
    else if(/^\d{4}-\d{2}-\d{2}$/.test(val)){ const [y,m,dd] = val.split('-').map(Number); d = new Date(y, m-1, dd); }
    else if(!isNaN(Date.parse(val))){ d = new Date(val); }
    else return String(val);
    const dd = String(d.getDate()).padStart(2,'0');
    const mm = String(d.getMonth()+1).padStart(2,'0');
    const yyyy = d.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }catch(_){ return String(val); }
}

// Helpers to set buffer Exp. Date about one year after base date
function toISO(d){ return new Date(d.getTime() - d.getTimezoneOffset()*60000).toISOString().slice(0,10); }
function setBufferExpDatesFromBase(baseDateStr, force=false){
  try{
    let base = null;
    if (baseDateStr && /^\d{4}-\d{2}-\d{2}$/.test(baseDateStr)) {
      const [y,m,dd] = baseDateStr.split('-').map(Number); base = new Date(y, m-1, dd);
    } else { base = new Date(); }
    const baseYear = base.getFullYear();
    const y = baseYear + 1, m = base.getMonth(), d = base.getDate();
    const offsets = [0,2,4];
    ensureDefaultBuffers();
    for (let i=0; i<state.buffers.length; i++){
      const r = state.buffers[i] || (state.buffers[i] = { equipment:'', brandModel:'', serial:'', expDate:'' });
      // If force, always update. Otherwise, update when empty or year <= base year
      let shouldUpdate = force;
      if (!shouldUpdate){
        const cur = String(r.expDate||'').trim();
        if (cur === '') shouldUpdate = true;
        else if (/^\d{4}-\d{2}-\d{2}$/.test(cur)){
          const [cy,cm,cd] = cur.split('-').map(Number);
          if (cy <= baseYear) shouldUpdate = true;
        } else {
          // unknown format -> update to safe value
          shouldUpdate = true;
        }
      }
      if (!shouldUpdate) continue;
      const off = offsets[i] || 0;
      const exp = new Date(y, m, d + off);
      r.expDate = toISO(exp);
    }
  }catch(_){ }
}

// Global safety: ensure the Start button always works even if init failed
function __gotoApp(){ try{ const landing=document.getElementById('landing'); const appView=document.getElementById('app-view'); if(landing&&appView){ landing.classList.add('hidden'); appView.classList.remove('hidden'); } location.hash = '#ph'; }catch(_){ } }

document.addEventListener('click', (e) => {
  try{
    const btn = e.target && (e.target.id === 'choose-ph' ? e.target : (e.target.closest ? e.target.closest('#choose-ph') : null));
    if (btn) {
      e.preventDefault();
      __gotoApp();
    }
  }catch(_){ /* noop */ }
});

// Mobile robustness: also handle touch and pointer interactions
['touchend','pointerup'].forEach(evt => {
  document.addEventListener(evt, (e) => {
    try{
      const el = e.target && (e.target.id === 'choose-ph' ? e.target : (e.target.closest ? e.target.closest('#choose-ph') : null));
      if (el) { e.preventDefault(); __gotoApp(); }
    }catch(_){ }
  }, { passive:false });
});

// Respond to hash changes, so navigating to #ph always opens the app view
window.addEventListener('hashchange', () => {
  try{
    if (location.hash === '#ph') __gotoApp();
  }catch(_){ }
});

// -------------------------------
// Minimal UI feedback: busy overlay + toast
let __busyEl = null, __toastEl = null, __toastTimer = null, __inlineCssInjected = false;

function __injectInlineCss(){
  if(__inlineCssInjected) return; __inlineCssInjected = true;
  const style = document.createElement('style');
  style.textContent = `
  #busy-ov{position:fixed;inset:0;background:rgba(0,0,0,.28);display:none;align-items:center;justify-content:center;z-index:10000}
  #busy-ov .card{background:#fff;border-radius:10px;padding:14px 16px;box-shadow:0 10px 30px rgba(0,0,0,.3);font-weight:700}
  #toast{position:fixed;left:50%;transform:translateX(-50%);bottom:18px;background:#111827;color:#fff;padding:10px 14px;border-radius:999px;box-shadow:0 8px 24px rgba(0,0,0,.2);z-index:10001;display:none;gap:10px;align-items:center}
  #toast.success{background:#0f766e} #toast.error{background:#b91c1c}
  #toast a{color:#fff;text-decoration:underline}
  `;
  document.head.appendChild(style);
}

function showBusy(text){
  __injectInlineCss();
  if(!__busyEl){
    __busyEl = document.createElement('div');
    __busyEl.id = 'busy-ov';
    __busyEl.innerHTML = '<div class="card">กำลังดำเนินการ...</div>';
    document.body.appendChild(__busyEl);
  }
  __busyEl.querySelector('.card').textContent = text || 'กำลังดำเนินการ...';
  __busyEl.style.display = 'flex';
}
function hideBusy(){ if(__busyEl){ __busyEl.style.display='none'; } }

function showToast(text, type='info', url){
  __injectInlineCss();
  if(!__toastEl){ __toastEl=document.createElement('div'); __toastEl.id='toast'; document.body.appendChild(__toastEl); }
  __toastEl.className = type ? type.replace(/[^\w-]/g,'') : '';
  __toastEl.innerHTML = url ? `${text} <a href="${url}" target="_blank" rel="noopener">เปิด</a>` : text;
  __toastEl.style.display='flex';
  clearTimeout(__toastTimer); __toastTimer = setTimeout(()=>{ __toastEl.style.display='none'; }, 4000);
}

// Override busy/toast texts with Thai defaults
function showBusy(text){
  __injectInlineCss();
  if(!__busyEl){
    __busyEl = document.createElement('div');
    __busyEl.id = 'busy-ov';
    __busyEl.innerHTML = '<div class="card">กำลังประมวลผล...</div>';
    document.body.appendChild(__busyEl);
  }
  __busyEl.querySelector('.card').textContent = text || 'กำลังประมวลผล...';
  __busyEl.style.display = 'flex';
}
function hideBusy(){ if(__busyEl){ __busyEl.style.display='none'; } }
function showToast(text, type='info', url){
  __injectInlineCss();
  if(!__toastEl){ __toastEl=document.createElement('div'); __toastEl.id='toast'; document.body.appendChild(__toastEl); }
  __toastEl.className = type ? type.replace(/[^\w-]/g,'') : '';
  __toastEl.innerHTML = url ? `${text} <a href="${url}" target="_blank" rel="noopener">เปิด</a>` : text;
  __toastEl.style.display='flex';
  clearTimeout(__toastTimer); __toastTimer = setTimeout(()=>{ __toastEl.style.display='none'; }, 4000);
}

// ---------------------------------
// Thai localization to fix garbled literals in HTML
function setLabelTextByName(name, text){
  try{
    const el = document.querySelector(`[name="${name}"]`);
    if(!el) return; const label = el.closest('label'); if(!label) return;
    const tn = Array.from(label.childNodes).find(n => n.nodeType === Node.TEXT_NODE);
    if (tn) tn.nodeValue = text + '\n'; else label.insertBefore(document.createTextNode(text + '\n'), label.firstChild);
  }catch(_){/* noop */}
}

function localizeThai(){
  try{
    const sub = document.querySelector('.landing-head .sub'); if (sub) sub.textContent = 'Analyzer Team Report';
    const choose = document.getElementById('choose-ph'); if (choose) choose.textContent = 'เริ่มใช้งาน';
    const hiDesc = document.querySelector('.card.highlight .card-desc'); if (hiDesc) hiDesc.textContent = 'กรอกแบบฟอร์มเพื่อสร้างเป็นไฟล์ PDF';
    document.querySelectorAll('.card.disabled .card-desc').forEach(el => el.textContent = 'เร็วๆ นี้');
    document.querySelectorAll('.card.disabled button').forEach(el => el.textContent = 'เร็วๆ นี้');
    const foot = document.querySelector('.landing-foot'); if (foot) foot.textContent = 'พัฒนาเพื่อใช้งานภายใน • สร้างไฟล์ PDF ได้แบบออฟไลน์';

    const backHome = document.getElementById('back-home'); if (backHome){ backHome.textContent = 'หน้าแรก'; backHome.title = 'กลับหน้าแรก'; }
    const loadDemo = document.getElementById('btn-load-demo'); if (loadDemo) loadDemo.textContent = 'โหลดตัวอย่าง';
    const saveJson = document.getElementById('btn-save-json'); if (saveJson) saveJson.textContent = 'บันทึก JSON';
    const importJson = document.querySelector('.import-json .button-like'); if (importJson) importJson.textContent = 'นำเข้า JSON';
    const btnPrev = document.getElementById('btn-preview'); if (btnPrev) btnPrev.textContent = 'ดูตัวอย่าง';
    const btnGen  = document.getElementById('btn-generate'); if (btnGen)  btnGen.textContent  = 'สร้างไฟล์ PDF';

    const tabDoc = document.querySelector('button[data-jump="#doc-section"]'); if (tabDoc) tabDoc.textContent = 'เอกสาร';
    const tabRef = document.querySelector('button[data-jump="#ref-section"]'); if (tabRef) tabRef.textContent = 'อ้างอิง';
    const tabCal = document.querySelector('button[data-jump="#cal-section"]'); if (tabCal) tabCal.textContent = 'บันทึก';
    const tabSign = document.querySelector('button[data-jump="#sign-section"]'); if (tabSign) tabSign.textContent = 'ลงชื่อ';
    const tabPhotos = document.querySelector('button[data-jump="#photos-section"]'); if (tabPhotos) tabPhotos.textContent = 'รูป';

    const smyDoc = document.querySelector('#doc-section > summary'); if (smyDoc) smyDoc.textContent = 'ข้อมูลเอกสาร';
    const smyCal = document.querySelector('#cal-section > summary'); if (smyCal) smyCal.textContent = 'บันทึกการสอบเทียบ';
    const smySign = document.querySelector('#sign-section > summary'); if (smySign) smySign.textContent = 'ลงชื่อผู้เกี่ยวข้อง';
    const smyPhotos = document.querySelector('#photos-section > summary'); if (smyPhotos) smyPhotos.textContent = 'รูปประกอบ (3 รูป)';

    setLabelTextByName('certificateNo', 'เลขที่เอกสาร'); const inpCert = document.querySelector('[name="certificateNo"]'); if (inpCert) inpCert.placeholder = 'เช่น AC/2025/0001';
    setLabelTextByName('calibrateDate', 'วันที่สอบเทียบ');
    const firstOpt = document.querySelector('select[name="tagNo"] option[value=""]'); if (firstOpt) firstOpt.textContent = '— เลือก —';
    const ae = document.querySelector('[name="acceptError"]'); if (ae){ setLabelTextByName('acceptError','Accept Error'); ae.placeholder = 'เช่น 0.5 pH หรือ % ของการอ่าน'; }
    const sol = document.querySelector('[name="solutionType"]'); if (sol) sol.placeholder = 'เช่น Buffer';

    const phs = {found4:'เช่น 3.98', cal4:'เช่น 4.01', found7:'เช่น 6.23', cal7:'เช่น 7.03', found10:'เช่น 9.97', cal10:'เช่น 10.01'}; for (const k in phs){ const el=document.querySelector(`[name="${k}"]`); if (el) el.placeholder = phs[k]; }

    setLabelTextByName('technician', 'ผู้สอบเทียบ');
    setLabelTextByName('approver', 'ผู้อนุมัติ (ผู้เชี่ยวชาญ)');
    setLabelTextByName('signDate', 'วันที่ลงชื่อ');
    const remarksLabel = document.querySelector('textarea[name="remarks"]')?.closest('label'); if (remarksLabel){ const tn = Array.from(remarksLabel.childNodes).find(n=>n.nodeType===Node.TEXT_NODE); if (tn) tn.nodeValue = 'หมายเหตุ\n'; }

    const ph1 = document.querySelector('[name="photo1"]')?.closest('label'); if (ph1){ const tn=Array.from(ph1.childNodes).find(n=>n.nodeType===Node.TEXT_NODE); if (tn) tn.nodeValue = 'รูปที่ 1\n'; }
    const ph2 = document.querySelector('[name="photo2"]')?.closest('label'); if (ph2){ const tn=Array.from(ph2.childNodes).find(n=>n.nodeType===Node.TEXT_NODE); if (tn) tn.nodeValue = 'รูปที่ 2\n'; }
    const ph3 = document.querySelector('[name="photo3"]')?.closest('label'); if (ph3){ const tn=Array.from(ph3.childNodes).find(n=>n.nodeType===Node.TEXT_NODE); if (tn) tn.nodeValue = 'รูปที่ 3\n'; }

    const pvTitle = document.querySelector('.preview-panel > h2'); if (pvTitle) pvTitle.textContent = 'ตัวอย่างเอกสาร';
    const mh = document.querySelector('#preview-modal .mh-title'); if (mh) mh.textContent = 'ตัวอย่างเอกสาร';
    const c1 = document.getElementById('preview-close'); if (c1){ c1.textContent = 'ปิด'; c1.setAttribute('aria-label','ปิด'); }
    const c2 = document.getElementById('preview-close-2'); if (c2){ c2.textContent = 'ปิด'; }

    const footerText = document.querySelector('.app-footer span'); if (footerText) footerText.textContent = 'เครื่องมือ Auto Report (ไฟล์ HTML) • สร้างไฟล์ PDF ได้แบบออฟไลน์';
    const ctaPrev = document.getElementById('cta-preview'); if (ctaPrev && !ctaPrev.classList.contains('icon')) ctaPrev.textContent = 'ดูตัวอย่าง';
    const ctaSave = document.getElementById('cta-save'); if (ctaSave && !ctaSave.classList.contains('icon')) ctaSave.textContent = 'สร้างไฟล์ PDF';

    // Fix Range separator text inside preview table
    const rangeRow = Array.from(document.querySelectorAll('#certificate .kv tr')).find(tr => tr.querySelector('th')?.textContent.trim() === 'Range');
    if (rangeRow){ const td = rangeRow.querySelector('td'); if (td) td.innerHTML = '<span data-bind="rangeMin"></span> ถึง <span data-bind="rangeMax"></span> <span data-bind="unit"></span>'; }
  }catch(_){/* noop */}
}

// Apply localization as soon as possible
if (document.readyState === 'loading') { document.addEventListener('DOMContentLoaded', localizeThai, { once:true }); } else { try{ localizeThai(); }catch(_){/*noop*/} }

// --- Post-fix for Thai text (override any garbled literals) ---
function __fixThaiText(){ try{ /* no-op: localization applied in localizeThai() */ }catch(_){ } }

  if (document.readyState === 'loading') { document.addEventListener('DOMContentLoaded', ()=> setTimeout(__fixThaiText, 0), { once:true }); }
else { setTimeout(__fixThaiText, 0); }




// --- Cross-device certificate no. sync (GAS/Drive counter) ---
async function ph_reserveCertViaAppsScript(){
  try{
    if(!GAS_URL) return null;
    const y=new Date().getFullYear();
    const resp=await fetch(GAS_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({op:'reserveCertNo',secret:GAS_SECRET||undefined,year:y,prefix:'AC',folderId:GAS_FOLDER_ID||undefined,module:'ph'})});
    const text=await resp.text(); let j={}; try{ j=JSON.parse(text);}catch(_){ }
    if(!resp.ok) return null;
    const n=(j&&j.num)?Number(j.num):null; if(!n||!isFinite(n)) return null;
    return {year:y,num:n,certificateNo:`AC/${y}/${String(n).padStart(4,'0')}`};
  }catch(_){ return null; }
}
async function ph_reserveCertViaDriveCounter(){
  try{
    const token=await getGoogleAccessToken();
    const y=new Date().getFullYear();
    const name=`cert-seq-ph-${y}.json`;
    const q=`name='${name.replace(/'/g,"\\'")}' and '${DRIVE_FOLDER_ID}' in parents and trashed=false`;
    let r=await fetch('https://www.googleapis.com/drive/v3/files?q='+encodeURIComponent(q)+'&fields=files(id,name)&pageSize=1&spaces=drive',{headers:{Authorization:'Bearer '+token}});
    let j={files:[]}; try{ j=await r.json(); }catch(_){ }
    let id=j.files&&j.files[0]&&j.files[0].id;
    if(!id){
      const meta={name,parents:[DRIVE_FOLDER_ID],mimeType:'application/json'};
      const boundary='m_'+Math.random().toString(36).slice(2);
      const body=new Blob([
        `--${boundary}\r\n`+'Content-Type: application/json; charset=UTF-8\r\n\r\n'+JSON.stringify(meta)+`\r\n`+
        `--${boundary}\r\n`+'Content-Type: application/json; charset=UTF-8\r\n\r\n'+JSON.stringify({last:149})+`\r\n`+
        `--${boundary}--`
      ],{type:'multipart/related; boundary='+boundary});
      const cr=await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',{method:'POST',headers:{Authorization:'Bearer '+token,'Content-Type':'multipart/related; boundary='+boundary},body});
      const cj=await cr.json(); id=cj&&cj.id; if(!id) return null;
    }
    let last=149; try{ const gr=await fetch(`https://www.googleapis.com/drive/v3/files/${id}?alt=media`,{headers:{Authorization:'Bearer '+token}}); const tx=await gr.text(); const o=JSON.parse(tx||'{}'); if(typeof o.last==='number') last=o.last; }catch(_){ }
    const num=(isFinite(last)?last:149)+1;
    try{ await fetch(`https://www.googleapis.com/upload/drive/v3/files/${id}?uploadType=media`,{method:'PATCH',headers:{Authorization:'Bearer '+token,'Content-Type':'application/json; charset=UTF-8'},body:JSON.stringify({last:num})}); }catch(_){ }
    return {year:y,num,certificateNo:`AC/${y}/${String(num).padStart(4,'0')}`};
  }catch(_){ return null; }
}
let __ph_reserving = false;
async function ph_prepareCertificateNoBeforeSave(){
  try{
    if(__ph_reserving) return null; __ph_reserving = true;
    let res=await ph_reserveCertViaAppsScript();
    if(!res) res=await ph_reserveCertViaDriveCounter();
    if(!res){ const y=new Date().getFullYear(); const key=certSeqStorageKey(y); let cur=Number(localStorage.getItem(key)||'149'); const num=cur+1; try{ localStorage.setItem(key,String(num)); }catch(_){ } res={year:y,num,certificateNo:`AC/${y}/${String(num).padStart(4,'0')}`}; }
    state.certificateNo=res.certificateNo; try{ const f=document.getElementById('data-form'); const el=f&&f.elements&&f.elements['certificateNo']; if(el) el.value=state.certificateNo; }catch(_){ } saveToLocal(); updatePreview(); return res;
  }catch(_){ return null; }
  finally { __ph_reserving = false; }
}
// Override default generate to reserve number first
try{
  const __orig_generatePDF = generatePDF;
  generatePDF = async function(){ try{ await ph_prepareCertificateNoBeforeSave(); }catch(_){ } return __orig_generatePDF.apply(this, arguments); };
}catch(_){ }

// Guard against double-bump within same save
(function(){ try{ let __ph_last_bump_ts = 0; const __orig = bumpCertificateNoAfterSave; bumpCertificateNoAfterSave = function(){ const now=Date.now(); if(now-__ph_last_bump_ts<800) return; __ph_last_bump_ts = now; return __orig.apply(this, arguments); }; }catch(_){ } })();

// Fallback downloader when Drive quota exceeded
(function(){
  function downloadBlob(blob, filename){ try{ const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=filename; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href), 1200); }catch(_){ } }
  try{
    if (typeof uploadToDriveQuiet === 'function'){
      const __orig_uploadToDriveQuiet = uploadToDriveQuiet;
      uploadToDriveQuiet = async function(blob, filename){
        try{ return await __orig_uploadToDriveQuiet(blob, filename); }
        catch(err){ try{ if(blob) downloadBlob(blob, filename); }catch(_){ } return { localDownload:true }; }
      };
    }
  }catch(_){ }
  try{
    if (typeof uploadViaAppsScript === 'function'){
      const __orig_uploadViaAppsScript = uploadViaAppsScript;
      uploadViaAppsScript = async function(blob, filename){
        try{ return await __orig_uploadViaAppsScript(blob, filename); }
        catch(err){ try{ if(blob) downloadBlob(blob, filename); }catch(_){ } return { localDownload:true }; }
      };
    }
  }catch(_){ }
})();




