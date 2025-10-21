// Conductivity certificate app (compact)
// Implements: tag presets (builtin + user), preview, PDF + Drive upload

// --- State ---
const state={buf1:"12880", buf2:"1413", calbuf1:"12880", calbuf2:"12880", certificateNo:"",calibrateDate:"",tagNo:"",tagName:"",location:"",maker:"",sensorModel:"",sensorSerial:"",transmitterModel:"",transmitterSerial:"",rangeMin:"",rangeMax:"",unit:"uS/cm",acceptError:"",solutionType:"Buffer",slope:"",offset:"",calResult:"Calibrated",technician:"",technicianPosition:"",approver:"",signDate:"",remarks:"",buffers:[],found1:"",found2:"",cal1:"",cal2:"",images:{logo:null,photo1:null,photo2:null,photo3:null},photoLabels:{photo1:"",photo2:"",photo3:""}};

// --- Config ---
const APP=window.APP_CONFIG||{};let GOOGLE_CLIENT_ID=APP.googleClientId||(typeof localStorage!=='undefined'&&localStorage.getItem('gdrive_client_id'))||'';const DRIVE_FOLDER_ID=(APP.condFolderId||'17-fq-lvMJYi8u37dtwDNlAr0doRe8Tf6');const UPLOAD_TO_DRIVE=(APP.uploadToDrive!==false);const GAS_URL=APP.gasUrl||'';const GAS_SECRET=APP.gasSecret||'';const GAS_FOLDER_ID=(APP.condFolderId||DRIVE_FOLDER_ID);let googleTokenClient=null,googleAccessToken=null;let LOGO_SVG_TEXT=null;

// --- Utils ---
const $=(s,e=document)=>e.querySelector(s),$$=(s,e=document)=>Array.from(e.querySelectorAll(s));
function download(fn,txt){const b=new Blob([txt],{type:'application/json'}),u=URL.createObjectURL(b),a=document.createElement('a');a.href=u;a.download=fn;a.click();URL.revokeObjectURL(u)}
function toast(m){try{let b=document.getElementById('toast');if(!b){b=document.createElement('div');b.id='toast';b.style.cssText='position:fixed;right:12px;bottom:12px;z-index:9999';document.body.appendChild(b)}const el=document.createElement('div');el.textContent=m;el.style.cssText='margin-top:6px;background:#111827;color:#fff;padding:10px 12px;border-radius:8px;box-shadow:0 8px 22px rgba(0,0,0,.2);font-size:14px';b.appendChild(el);setTimeout(()=>el.remove(),2400)}catch(_){}}
function busy(t){let o=document.getElementById('busy-ov');if(!o){o=document.createElement('div');o.id='busy-ov';o.style.cssText='position:fixed;inset:0;background:rgba(17,24,39,.45);display:flex;align-items:center;justify-content:center;z-index:9998';const m=document.createElement('div');m.id='busy-ov-msg';m.style.cssText='background:#fff;padding:12px 16px;border-radius:10px;box-shadow:0 16px 40px rgba(0,0,0,.25);font-weight:700';m.textContent=t||'Working...';o.appendChild(m);document.body.appendChild(o)}else{$('#busy-ov-msg').textContent=t||'Working...';o.style.display='flex'}}
function idle(){const o=document.getElementById('busy-ov');if(o) o.style.display='none'}
function formatDateDMY(v){if(!v) return'';try{const d=new Date(v);if(isNaN(+d))return'';const dd=String(d.getDate()).padStart(2,'0'),mm=String(d.getMonth()+1).padStart(2,'0'),yy=d.getFullYear();return`${dd}/${mm}/${yy}`}catch(_){return''}}

// --- Render ---
function bindText(){$$('[data-bind]').forEach(el=>{const k=el.getAttribute('data-bind');el.textContent=state[k]||''})}
function setImg(id,u){const el=document.getElementById(id);if(!el)return;el.src=u||'data:image/gif;base64,R0lGODlhAQABAAAAACw=';el.style.visibility='visible'}
function ensureDefaultBuffers(){if(!Array.isArray(state.buffers)||state.buffers.length===0){state.buffers=[{equipment:'Buffer Conduct Solution 500',brandModel:'TAFTEK instruments',serial:'',expDate:''},{equipment:'Buffer Conduct Solution 1413',brandModel:'TAFTEK instruments',serial:'',expDate:''},{equipment:'Buffer Conduct Solution 12880',brandModel:'TAFTEK instruments',serial:'',expDate:''}]}}
function renderBuffers(){const tb=$('#buffer-rows');if(!tb)return;tb.innerHTML='';for(const r of state.buffers){if(!r)continue;const s=(r.serial&&String(r.serial).trim()!=='')?r.serial:'-';const tr=document.createElement('tr');tr.innerHTML=`<td>${r.equipment||''}</td><td>${r.brandModel||''}</td><td>${s}</td><td>${formatDateDMY(r.expDate)||''}</td>`;tb.appendChild(tr)}}

// Build calibration result tables (As Found / As Calibrate)
function renderCalPoints(){
  try{
    const tbL = document.getElementById('cal-left');
    const tbR = document.getElementById('cal-right');
    if(!tbL || !tbR) return;
    const fmt = v => (v===''||v==null||Number.isNaN(v)) ? '' : (Math.round(v*100)/100).toFixed(2);
    const mkRow = (buf, reading, suppress=false)=>{
      const b = Number(buf);
      const r = Number(reading);
      const has = !Number.isNaN(r) && reading!=='' && reading!=null && !suppress;
      const err = has && !Number.isNaN(b) ? (r - b) : '';
      const pct = (has && b) ? ((r - b)/b*100) : '';
      return { buffer: (Number.isNaN(b)?'':b), reading: has? r: (suppress?'-':''), pct: suppress?'-':fmt(pct), abs: suppress?'-':fmt(err) };
    };
    const uncal = String(state.calResult||'').toLowerCase().startsWith('un');
    const foundRows = [ mkRow(state.buf1, state.found1, false), mkRow(state.buf2, state.found2, false) ];
    const calRows   = [ mkRow(state.calbuf1, state.cal1, uncal), mkRow(state.calbuf2, state.cal2, uncal) ];
    tbL.innerHTML=''; tbR.innerHTML='';
    foundRows.forEach(r=>{ const tr=document.createElement('tr'); tr.innerHTML=`<td>${r.buffer}</td><td>${r.reading}</td><td>${r.pct}</td><td>${r.abs}</td>`; tbL.appendChild(tr); });
    calRows.forEach(r=>{ const tr=document.createElement('tr'); tr.innerHTML=`<td>${r.buffer}</td><td>${r.reading}</td><td>${r.pct}</td><td>${r.abs}</td>`; tbR.appendChild(tr); });
    // Always show the calibration block, even when no values yet
    try{
      const block = document.getElementById('calibration-record');
      if(block) block.style.display = '';
    }catch(_){ }
  }catch(_){ }
}

// Ensure default NPS logo exists as embedded PNG (works well with html2canvas)
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
    try{ saveToLocal(); }catch(_){ }
  }catch(_){ }
}

function updatePreview(){
  try{ bindText(); }catch(_){ }
  try{ renderBuffers(); }catch(_){ }
  try{ renderCalPoints(); }catch(_){ }
  try{
    const box=document.getElementById('logo-box');
    if(box){
      box.style.background='none'; box.style.border='none'; box.style.padding='0';
      box.style.width='32mm'; box.style.height='26mm';
      const src=(state.images&&state.images.logo)?state.images.logo:'assets/logo_nps.svg';
      box.innerHTML=`<img src="${src}" alt="logo" style="width:100%;height:100%;object-fit:contain;display:block;"/>`;
    }
  }catch(_){ }
  try{ setImg('img1', state.images.photo1); setImg('img2', state.images.photo2); setImg('img3', state.images.photo3); }catch(_){ }
}
  try{ updateInlineThumbs(); }catch(_){ }
// --- Exp. Date helpers (same behavior as pH)
function toISO(d){ return new Date(d.getTime() - d.getTimezoneOffset()*60000).toISOString().slice(0,10); }
function setBufferExpDatesFromBase(baseDateStr, force=false){
  try{
    let base=null;
    if (baseDateStr && /^\d{4}-\d{2}-\d{2}$/.test(baseDateStr)){
      const [y,m,dd] = baseDateStr.split('-').map(Number); base = new Date(y, m-1, dd);
    } else { base = new Date(); }
    const baseYear = base.getFullYear();
    const y = baseYear + 1, m = base.getMonth(), d = base.getDate();
    const offsets = [0,2,4];
    ensureDefaultBuffers();
    for (let i=0;i<state.buffers.length;i++){
      const r = state.buffers[i] || (state.buffers[i] = { equipment:'', brandModel:'', serial:'', expDate:'' });
      let shouldUpdate = force;
      if (!shouldUpdate){
        const cur = String(r.expDate||'').trim();
        if (cur === '') shouldUpdate = true;
        else if (/^\d{4}-\d{2}-\d{2}$/.test(cur)){
          const [cy] = cur.split('-').map(Number);
          if (cy <= baseYear) shouldUpdate = true;
        } else { shouldUpdate = true; }
      }
      if (!shouldUpdate) continue;
      const off = offsets[i] || 0;
      const exp = new Date(y, m, d + off);
      r.expDate = toISO(exp);
    }
  }catch(_){ }
}


// --- Builtin tag presets ---
const BUILTIN_TAGS={
 '402QT005':{desc:'SEWER TO EFF',location:'Pulp2',brand:'ABB',sensorModel:'TOROTDAL',sensorSerial:'3K610000147263',model:'TB82',transmitterSerial:'3K610000147611',rangeMin:0,rangeMax:5,unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'402QT011':{desc:'COND SUMP PUMP',location:'Pulp2',brand:'ABB',sensorModel:'TOROTDAL',sensorSerial:'3K610000147625',model:'TB82',transmitterSerial:'TB82TC2010112',rangeMin:0,rangeMax:30,unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'402QE118':{desc:'COND SEALING PUMP',location:'Pulp2',brand:'ABB',sensorModel:'FOUR - ELECTRODE',sensorSerial:'3k610000264722',model:'TB82',transmitterSerial:'3K610000118038',rangeMin:0,rangeMax:2.5,unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'412QE117':{desc:'CONDUCTIVITY CONDENSATE',location:'Pulp2',brand:'ROSEMOUNT',sensorModel:'TOROTDAL',sensorSerial:'-',model:'-',transmitterSerial:'-',rangeMin:0,rangeMax:2,unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'412QE118':{desc:'CONDUCTIVITY BL. COOLING',location:'Pulp2',brand:'ROSEMOUNT',sensorModel:'-',sensorSerial:'-',model:'-',transmitterSerial:'-',rangeMin:0,rangeMax:1,unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'402QE824':{desc:'-',location:'Pulp2',brand:'ABB',sensorModel:'TOROTDAL',sensorSerial:'3K610000258010',model:'TB82',transmitterSerial:'3K610000147612',rangeMin:'',rangeMax:'',unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'412QE801':{desc:'CONDUCTIVITY HOT WATER TANK',location:'Pulp2',brand:'ABB',sensorModel:'TOROTDAL',sensorSerial:'3K610000234930',model:'TB82',transmitterSerial:'3K610000234920',rangeMin:0,rangeMax:2,unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'432QE805':{desc:'PREBLEACH WP.',location:'Pulp2',brand:'ABB',sensorModel:'TOROTDAL',sensorSerial:'-',model:'TB82',transmitterSerial:'-',rangeMin:0,rangeMax:199.9,unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'432QE806':{desc:'PULP TO EOP WP.',location:'Pulp2',brand:'ABB',sensorModel:'TOROTDAL',sensorSerial:'3K610000234934',model:'TB82',transmitterSerial:'3K610000234912',rangeMin:0,rangeMax:20,unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'432QE807':{desc:'PULP TO D1  WP.',location:'Pulp2',brand:'ABB',sensorModel:'TOROTDAL',sensorSerial:'3K610000234931',model:'TB82',transmitterSerial:'3K610000234918',rangeMin:0,rangeMax:20,unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'432QE809':{desc:'WHITE WATER FROM WETLAP',location:'Pulp2',brand:'ABB',sensorModel:'TOROTDAL',sensorSerial:'3K610000234925',model:'TB82',transmitterSerial:'TB82TC2010112',rangeMin:0,rangeMax:20,unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'432QE810':{desc:'WHITE WATER FROM PM3',location:'Pulp2',brand:'ABB',sensorModel:'FOUR - ELECTRODE',sensorSerial:'3K610000234927',model:'TB82',transmitterSerial:'3K610000217144',rangeMin:0,rangeMax:20,unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'422QE333':{desc:'-',location:'Pulp2',brand:'ABB',sensorModel:'FOUR - ELECTRODE',sensorSerial:'-',model:'TB82',transmitterSerial:'3K610000155309',rangeMin:0,rangeMax:50,unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'432QT812':{desc:'ALKALINE EOP WP',location:'Pulp2',brand:'ABB',sensorModel:'',sensorSerial:'3K610000234926',model:'TB82',transmitterSerial:'TB82TC2010112',rangeMin:'',rangeMax:'',unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'432QT026':{desc:'',location:'Pulp2',brand:'ABB',sensorModel:'TOROTDAL',sensorSerial:'3K610000234929',model:'TB82',transmitterSerial:'TB82TC2010112',rangeMin:'',rangeMax:'',unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'NO TAG':{desc:'',location:'Pulp2',brand:'ABB',sensorModel:'TOROTDAL',sensorSerial:'',model:'AWT210',transmitterSerial:'3K2100000135327',rangeMin:0,rangeMax:200,unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'no tag.':{desc:'',location:'Pulp2',brand:'ABB',sensorModel:'TOROTDAL',sensorSerial:'3K610000234935',model:'TBA82',transmitterSerial:'TB62TC010112',rangeMin:0,rangeMax:20,unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'432QT804':{desc:'',location:'Pulp2',brand:'ABB',sensorModel:'TOROTDAL',sensorSerial:'3K610000234993',model:'TBA82',transmitterSerial:'TB82TC2010112',rangeMin:'',rangeMax:'',unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'472CT103':{desc:'-',location:'LK2',brand:'ABB',sensorModel:'TOROTDAL',sensorSerial:'',model:'AWT210',transmitterSerial:'3K220000844450',rangeMin:0,rangeMax:20,unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'412QE802':{desc:'PULP TO MCO2 WP',location:'Pulp2',brand:'ABB',sensorModel:'TOROTDAL',sensorSerial:'-',model:'TB82',transmitterSerial:'-',rangeMin:0,rangeMax:200,unit:'mS/cm',acceptError:'5% (Error of Reading)'}
,'484QT701':{desc:'-',location:'QL',brand:'ABB',sensorModel:'TOROTDAL',sensorSerial:'-',model:'TB82TC201012',transmitterSerial:'3K610000234920',rangeMin:0,rangeMax:20,unit:'mS/cm',acceptError:'5% (Error of Reading)'}
};
function loadLocalTagPresets(){try{return JSON.parse(localStorage.getItem('cond-tag-presets')||'{}')}catch(_){return{}}}
function saveLocalTagPresets(o){try{localStorage.setItem('cond-tag-presets',JSON.stringify(o))}catch(_){}}
function ensureTagOptionInSelect(tag){try{const s=document.querySelector('select[name="tagNo"]');if(!s||!tag)return;if(![...s.options].some(o=>o.value===tag)){const opt=document.createElement('option');opt.value=tag;opt.textContent=tag;s.appendChild(opt)}}catch(_){}}
function parseRange(str){if(!str)return{};const m=String(str).match(/(-?\d+(?:\.\d+)?)\s*to\s*(-?\d+(?:\.\d+)?)/i);return m?{min:+m[1],max:+m[2]}:{}}
async function maybeCreatePresetInteractively(tag){if(!tag)return null;if(!confirm('Create preset for this Tag?'))return null;const brand=prompt('Maker/Brand:','ABB')||'';const model=prompt('Transmitter Model:','TB82')||'';const calRange=prompt('Range (e.g. 0 to 20):','0 to 20')||'';const desc=prompt('Tag Name/Description:','')||'';const sensorModel=prompt('Sensor Model (optional):','')||'';const sensorSerial=prompt('Sensor Serial (optional):','')||'';const transmitterSerial=prompt('Transmitter Serial (optional):','')||'';const acceptError=prompt('Accept Error:','5% (Error of Reading)')||'5% (Error of Reading)';const unit=prompt('Unit:',state.unit||'uS/cm')||(state.unit||'uS/cm');const rng=parseRange(calRange);const p={brand,model,desc,location:state.location||'Pulp2',sensorModel,sensorSerial,transmitterSerial,rangeMin:rng.min,rangeMax:rng.max,unit,acceptError};const db=loadLocalTagPresets();db[tag]=p;saveLocalTagPresets(db);return p}
function applyTagPreset(tag){if(!tag)return;const db=loadLocalTagPresets();const p=db[tag]||BUILTIN_TAGS[tag];if(!p)return;state.tagNo=tag;state.tagName=p.desc||'';state.location=p.location||'';state.maker=p.brand||'';state.sensorModel=p.sensorModel||'';state.sensorSerial=p.sensorSerial||'';state.transmitterModel=p.model||'';state.transmitterSerial=p.transmitterSerial||'';if(p.rangeMin!==''&&p.rangeMin!=null)state.rangeMin=p.rangeMin;if(p.rangeMax!==''&&p.rangeMax!=null)state.rangeMax=p.rangeMax;state.unit=p.unit||state.unit;state.acceptError=p.acceptError||state.acceptError;fillFormFromState();updatePreview()}
function saveCurrentTagPreset(){const tag=(state.tagNo||'').trim();if(!tag)return false;const p={brand:state.maker||'',model:state.transmitterModel||'',desc:state.tagName||'',location:state.location||'',sensorModel:state.sensorModel||'',sensorSerial:state.sensorSerial||'',transmitterSerial:state.transmitterSerial||'',rangeMin:state.rangeMin,rangeMax:state.rangeMax,unit:state.unit||'uS/cm',acceptError:state.acceptError||'5% (Error of Reading)'};const db=loadLocalTagPresets();db[tag]=p;saveLocalTagPresets(db);ensureTagOptionInSelect(tag);return true}
function enhanceTagSelect(){const s=document.querySelector('select[name="tagNo"]');if(!s)return;const db=loadLocalTagPresets();const keys=Array.from(new Set(Object.keys(BUILTIN_TAGS).concat(Object.keys(db)))).sort((a,b)=>String(a).localeCompare(String(b)));s.innerHTML='<option value="">Selectโ€ฆ</option><option value="__new__">+ Add new Tag</option>'+keys.map(k=>`<option value="${k}">${k}</option>`).join('');s.addEventListener('change',async()=>{const v=s.value;if(v==='__new__'){const code=prompt('New Tag No.:','');if(!code){s.value='';return}const tag=String(code).trim();ensureTagOptionInSelect(tag);s.value=tag;state.tagNo=tag;const db=loadLocalTagPresets();fillFormFromState();updatePreview();saveToLocal()}else{applyTagPreset(v);saveToLocal()}});if(state.tagNo){ensureTagOptionInSelect(state.tagNo);s.value=state.tagNo}const btn=document.getElementById('btn-update-tag');if(btn)btn.addEventListener('click',()=>{if(saveCurrentTagPreset())toast('Saved tag preset');}); try{ var wrap=document.querySelector('.table-controls'); if(wrap && !document.getElementById('btn-delete-tag')){ var del=document.createElement('button'); del.id='btn-delete-tag'; del.type='button'; del.textContent='Delete this Tag'; del.style.marginLeft='8px'; wrap.appendChild(del); del.addEventListener('click', function(){ var tag=(state.tagNo||'').trim(); if(!tag){ alert('Please select a Tag first'); return; } var local=loadLocalTagPresets(); if(!local[tag]){ alert('Cannot delete built-in Tag'); return; } if(!confirm('Delete this Tag preset?')) return; if(deleteTagPreset(tag)){ toast('Deleted tag '+tag); } }); } }catch(_){ }}

// --- Preview modal ---
function ensurePreviewModal(){
  let modal = document.getElementById('preview-modal');
  if (!modal){
    modal = document.createElement('div'); modal.id='preview-modal'; modal.className='modal hidden';
    modal.innerHTML = `
      <div class="modal-backdrop" data-close></div>
      <div class="modal-dialog">
        <header class="modal-head">
          <div class="mh-title">Preview</div>
          <button id="preview-close" class="modal-close" type="button" aria-label="Close">Close</button>
        </header>
        <div class="modal-body"><div id="preview-mount"></div></div>
        <footer class="modal-foot"><button id="preview-close-2" type="button">Close</button></footer>
      </div>`;
    document.body.appendChild(modal);
  }
  return modal;
}
async function openPreviewModal(){
  try{
    updatePreview();
    ensurePreviewModal();
    const m=$('#preview-modal'), mount=$('#preview-mount'), base=$('#certificate');
    if(!m||!mount||!base){ alert('Preview container missing'); return; }
    const c=base.cloneNode(true); c.id='certificate-preview'; c.classList.remove('pdf');
    mount.innerHTML=''; mount.appendChild(c);
    m.classList.remove('hidden'); document.body.style.overflow='hidden';
    const close=()=>{ m.classList.add('hidden'); mount.innerHTML=''; document.body.style.overflow=''; };
    $('#preview-close')?.addEventListener('click',close,{once:true});
    $('#preview-close-2')?.addEventListener('click',close,{once:true});
    m.querySelector('[data-close]')?.addEventListener('click',close,{once:true});
    const onKey=e=>{ if(e.key==='Escape'){ close(); window.removeEventListener('keydown',onKey); } };
    window.addEventListener('keydown',onKey);
  }catch(err){ console.error('Preview error', err); alert('Cannot open preview: '+(err&&err.message?err.message:err)); }
}

// --- PDF + upload ---
function makeBaseFilename(){const tag=(state.tagNo||'TAG').replace(/[^\w\-]+/g,'_'),date=state.calibrateDate||new Date().toISOString().slice(0,10);return`Conductivity_${tag}_${date}`}
async function waitImages(el){const imgs=[...el.querySelectorAll('img')];await Promise.all(imgs.map(i=>i.complete&&i.naturalWidth?null:new Promise(r=>{i.onload=i.onerror=()=>r()})));await Promise.all(imgs.map(i=>i.decode?.().catch(()=>{})))}
function certSeqKey(){const y=new Date().getFullYear();return'cond-cert-seq-'+y}
function bumpCertificateNoAfterSave(){
  try{
    const y = new Date().getFullYear();
    const k = certSeqKey();
    let cur = Number(localStorage.getItem(k) || '149');
    // "cur" already reflects the currently shown certificate no. from ensureCertificateNo
    const used = cur; // this one was just used for the saved PDF
    // mark used in storage (kept the same value)
    localStorage.setItem(k, String(used));
    // compute next certificate no. for immediate display
    const next = used + 1;
    state.certificateNo = `AC/${y}/${String(next).padStart(4,'0')}`;
    // reflect to form input if present
    try{
      const form = document.getElementById('data-form');
      const el = form && form.elements && form.elements['certificateNo'];
      if (el) el.value = state.certificateNo;
    }catch(_){ }
    // persist and refresh preview so user immediately sees the increment
    saveToLocal();
    updatePreview();
  }catch(_){ }
}
async function generateAndUploadPDF(){const base=$('#certificate');if(!base)return;const filename=makeBaseFilename()+'.pdf';const box=document.createElement('div');box.id='pdf-sandbox';box.style.cssText='position:fixed;left:0;top:0;width:210mm;z-index:-1;opacity:0;pointer-events:none;';const el=base.cloneNode(true);el.id='certificate-print';document.body.appendChild(box);box.appendChild(el);el.classList.add('pdf');await waitImages(el);try{if(document.fonts&&document.fonts.ready){await document.fonts.ready}}catch(_){ }const opt={margin:[0,0,0,0],filename,image:{type:'jpeg',quality:0.98},html2canvas:{scale:2,useCORS:true,backgroundColor:'#ffffff',scrollX:0,scrollY:0},jsPDF:{unit:'mm',format:'a4',orientation:'portrait'}};try{busy('Creating PDF...');let blob=null;await window.html2pdf().from(el).set(opt).toPdf().get('pdf').then(pdf=>{try{const n=typeof pdf.getNumberOfPages==='function'?pdf.getNumberOfPages():1;for(let i=n;i>=2;i--){pdf.deletePage(i)}}catch(_){ }blob=pdf.output('blob')});busy('Uploading to Google Drive...');if(GAS_URL){const res=await uploadViaAppsScript(blob,filename);const url=(res&&(res.url||res.webViewLink||(res.fileId?('https://drive.google.com/file/d/'+res.fileId+'/view'):'')))||'';bumpCertificateNoAfterSave();toast('Saved: '+filename+(url?' โ“':''))}else{const meta=await uploadToDriveQuiet(blob,filename);const url=meta&&(meta.webViewLink||(meta.id?('https://drive.google.com/file/d/'+meta.id+'/view'):''));bumpCertificateNoAfterSave();toast('Saved: '+filename+(url?' โ“':''))}}catch(err){console.error(err);toast('Save failed')}finally{idle();try{box.remove()}catch(_){}}}
function loadGIS(){return new Promise(r=>{if(window.google&&google.accounts&&google.accounts.oauth2){r();return}const iv=setInterval(()=>{if(window.google&&google.accounts&&google.accounts.oauth2){clearInterval(iv);r()}},100)})}
async function getGoogleAccessToken(){await loadGIS();return new Promise((res,rej)=>{try{if(googleAccessToken){res(googleAccessToken);return}if(!googleTokenClient){googleTokenClient=google.accounts.oauth2.initTokenClient({client_id:GOOGLE_CLIENT_ID,scope:'https://www.googleapis.com/auth/drive.file',callback:(rp)=>{if(rp&&rp.access_token){googleAccessToken=rp.access_token;res(googleAccessToken)}else rej(new Error('no access token'))}})}const authed=(typeof localStorage!=='undefined'&&localStorage.getItem('gdrive_authed')==='1');googleTokenClient.requestAccessToken({prompt:authed?'':'consent'});try{localStorage.setItem('gdrive_authed','1')}catch(_){}}catch(e){rej(e)}})}
async function uploadToDrive(blob,fn,token){const meta={name:fn,parents:[DRIVE_FOLDER_ID],mimeType:'application/pdf'};const b='b_'+Math.random().toString(36).slice(2);const body=new Blob([`--${b}\r\n`+'Content-Type: application/json; charset=UTF-8\r\n\r\n'+JSON.stringify(meta)+`\r\n`+`--${b}\r\n`+'Content-Type: application/pdf\r\n\r\n',blob,`\r\n--${b}--`],{type:'multipart/related; boundary='+b});const r=await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',{method:'POST',headers:{Authorization:'Bearer '+token,'Content-Type':'multipart/related; boundary='+b},body});if(!r.ok){let t='';try{t=await r.text()}catch(_){ }throw new Error('Drive upload HTTP '+r.status+' '+t)}return r.json()}
async function uploadToDriveQuiet(blob,fn){if(!blob)return null;if(!DRIVE_FOLDER_ID||!UPLOAD_TO_DRIVE)return null;if(!GOOGLE_CLIENT_ID)throw new Error('Missing Google OAuth Client ID');const t=await getGoogleAccessToken();return uploadToDrive(blob,fn,t)}
async function uploadViaAppsScript(blob,fn){const b64=await new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(String(r.result).split(',')[1]);r.onerror=rej;r.readAsDataURL(blob)});const rp=await fetch(GAS_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({secret:GAS_SECRET||undefined,folderId:GAS_FOLDER_ID||undefined,filename:fn,mimeType:'application/pdf',fileBase64:b64})});const tx=await rp.text();let j={};try{j=JSON.parse(tx)}catch(_){ }if(!rp.ok||(j&&j.success===false)){const m=(j&&(j.error||j.message))||tx||('HTTP '+rp.status);throw new Error('Apps Script upload failed: '+m)}return j}

// --- Form wiring ---
// Populate technician/approver like pH
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
function loadApprovers(){try{const a=JSON.parse(localStorage.getItem('approver-list')||'[]');return Array.isArray(a)?a:[]}catch(_){return[]}}
function saveApprovers(l){try{localStorage.setItem('approver-list',JSON.stringify(l))}catch(_){}}
function ensureApproverSelect(){const sel=document.getElementById('approver-select');if(!sel)return;let list=loadApprovers();if(!list||list.length===0)list=DEFAULT_APPROVERS.slice();sel.innerHTML='<option value="">Select…</option><option value="__new__">+ Add approver</option>'+list.map(a=>`<option value="${a.name}" data-pos="${a.position||''}">${a.name}</option>`).join('');if(state.approver&&![...sel.options].some(o=>o.value===state.approver)){const opt=document.createElement('option');opt.value=state.approver;opt.textContent=state.approver;sel.appendChild(opt);}sel.value=state.approver||'';sel.onchange=()=>{const v=sel.value;if(v==='__new__'){const name=prompt('Approver name:','');if(!name){sel.value='';return}const pos=prompt('Position:','Engineer')||'Engineer';const arr=loadApprovers();if(!arr.some(x=>x&&x.name===name)){arr.push({name,position:pos});saveApprovers(arr);}ensureApproverSelect();sel.value=name;state.approver=name;saveToLocal();updatePreview();}else{state.approver=v;saveToLocal();updatePreview();}}}
function populateTechnicianSelect(){const sel=document.getElementById('technician-select');if(!sel)return;sel.innerHTML='<option value="">Select…</option>'+CALIBRATORS.map(c=>`<option value="${c.name}" data-pos="${c.position||''}">${c.name}</option>`).join('');if(state.technician&&![...sel.options].some(o=>o.value===state.technician)){const opt=document.createElement('option');opt.value=state.technician;opt.textContent=state.technician;opt.setAttribute('data-pos',state.technicianPosition||'');sel.appendChild(opt);}sel.value=state.technician||'';sel.onchange=()=>{state.technician=sel.value||'';const pos=sel.selectedOptions[0]?.getAttribute('data-pos')||'';state.technicianPosition=pos;const hint=document.getElementById('tech-role-inline');if(hint)hint.textContent=pos?`(${pos})`:'';saveToLocal();updatePreview();};const hint=document.getElementById('tech-role-inline');if(hint)hint.textContent=state.technicianPosition?`(${state.technicianPosition})`:'';}
function fillFormFromState(){const f=$('#data-form');if(!f)return;$$('input[name], select[name], textarea[name]',f).forEach(el=>{if(['logo','photo1','photo2','photo3'].includes(el.name))return;if(el.tagName==='SELECT'){const v=state[el.name]??'';el.value=v;if(v&&el.value!==v){const o=document.createElement('option');o.value=v;o.textContent=v;el.appendChild(o);el.value=v}}else{el.value=state[el.name]??''}})}
function rebuildEditableTables(){ensureDefaultBuffers();const tb=$('#buffer-table tbody');if(!tb)return;tb.innerHTML='';state.buffers.forEach((r,i)=>{const tr=document.createElement('tr');tr.innerHTML=`<td><input data-k=\"equipment\" data-i=\"${i}\" value=\"${r.equipment??''}\"></td><td><input data-k=\"brandModel\" data-i=\"${i}\" value=\"${r.brandModel??''}\"></td><td><input data-k=\"serial\" data-i=\"${i}\" value=\"${r.serial??''}\"></td><td><input type=\"date\" data-k=\"expDate\" data-i=\"${i}\" value=\"${r.expDate??''}\"></td>`;tb.appendChild(tr)})}
function wireEditableHandlers(){const b=$('#buffer-table tbody');if(b)b.addEventListener('input',e=>{const t=e.target;if(!(t instanceof HTMLInputElement))return;const i=Number(t.getAttribute('data-i')),k=t.getAttribute('data-k');if(!Number.isFinite(i)||!k)return;if(!state.buffers[i])state.buffers[i]={equipment:'',brandModel:'',serial:'',expDate:''};state.buffers[i][k]=t.value;renderBuffers()})}
function wireFormBasics(){
  const f=$('#data-form'); if(!f) return;
  const onCh=e=>{
    const t=e.target; if(!t||!t.name) return;
    if(['logo','photo1','photo2','photo3'].includes(t.name)) return;
    state[t.name]=t.value;
    if(t.name==='calibrateDate'){
      try{ setBufferExpDatesFromBase(state.calibrateDate, true); rebuildEditableTables(); }catch(_){ }
    }
    if(t.name==='calResult'){
      try{ updateCalInputsAvailability(); }catch(_){ }
    }
    updatePreview();
  };
  f.addEventListener('input',onCh);
  f.addEventListener('change',onCh);
  ['photo1','photo2','photo3'].forEach(k=>{
    const i=f.elements[k]; if(!i) return;
    i.addEventListener('change', async()=>{
      const file=i.files && i.files[0];
      if(!file){ state.images[k]=null; updatePreview(); return; }
      const data=await new Promise((res,rej)=>{ const r=new FileReader(); r.onload=()=>res(String(r.result)); r.onerror=rej; r.readAsDataURL(file); });
      state.images[k]=data; updatePreview();
    });
  });
}

// Disable Calibrate inputs when Uncalibrated and show '-'
function updateCalInputsAvailability(){
  try{
    const uncal = String(state.calResult||'').toLowerCase().startsWith('un');
    ['cal1','cal2','calbuf1','calbuf2'].forEach(name=>{
      const el = document.querySelector(`[name="${name}"]`);
      if(!el) return;
      el.disabled = uncal;
      if(uncal){ el.value=''; el.placeholder='-'; } else { el.placeholder=''; }
    });
  }catch(_){ }
}

// --- Demo + init ---
function toJSON(){return JSON.stringify(state,null,2)}
function saveToLocal(){try{localStorage.setItem('cond-cert-state',JSON.stringify(state))}catch(_){}}
function loadFromLocal(){try{const s=localStorage.getItem('cond-cert-state');if(!s)return false;Object.assign(state,JSON.parse(s));return true}catch(_){return false}}
function ensureCertificateNo(){const y=new Date().getFullYear(),k='cond-cert-seq-'+y;let c=Number(localStorage.getItem(k)||'149');const n=c+1;state.certificateNo=`AC/${y}/${String(n).padStart(4,'0')}`;try{localStorage.setItem(k,String(n))}catch(_){}}
function loadDemo(){
  ensureCertificateNo();
  const d=new Date().toISOString().slice(0,10);
  Object.assign(state,{calibrateDate:d,tagNo:'4B407T01',tagName:'ALKALINE EFF cond',location:'Pulp2',maker:'ABB',sensorModel:'TB557',transmitterModel:'TB82',rangeMin:0,rangeMax:20000,unit:'uS/cm',acceptError:'5% (Error of Reading)',solutionType:'Buffer',slope:'1.22',offset:'-64.1',calResult:'Calibrated',technician:'Phutthiphong B.',approver:'Anusit s.',signDate:d});
  state.found1=1541; state.found2=11763;
  state.cal1=1540;   state.cal2=12880;
  ensureDefaultBuffers(); fillFormFromState(); updatePreview(); updateCalInputsAvailability();
}

async function init(){try{$('#app-view')?.classList.remove('hidden')}catch(_){ }const had=loadFromLocal();if(!had)ensureCertificateNo();ensureDefaultBuffers();rebuildEditableTables();wireEditableHandlers();wireFormBasics();fillFormFromState();try{await ensureDefaultLogo();}catch(_){ }updatePreview(); setBufferExpDatesFromBase(state.calibrateDate || new Date().toISOString().slice(0,10), false); rebuildEditableTables(); enhanceTagSelect();populateTechnicianSelect();ensureApproverSelect(); updateCalInputsAvailability(); $('#btn-load-demo')?.addEventListener('click',()=>{loadDemo();saveToLocal();updatePreview()});$('#btn-save-json')?.addEventListener('click',()=>download(makeBaseFilename()+'.json',toJSON()));$('#input-load-json')?.addEventListener('change',async e=>{try{const f=e.target.files[0];const t=await f.text();Object.assign(state,JSON.parse(t));fillFormFromState();rebuildEditableTables();try{await ensureDefaultLogo();}catch(_){ }updatePreview();saveToLocal();updateCalInputsAvailability()}catch(_){}});$('#btn-preview')?.addEventListener('click',async()=>{try{await ensureDefaultLogo();}catch(_){ }openPreviewModal()});$('#btn-generate')?.addEventListener('click',async()=>{try{await ensureDefaultLogo();}catch(_){ }generateAndUploadPDF()});$('#cta-preview')?.addEventListener('click',async()=>{try{await ensureDefaultLogo();}catch(_){ }openPreviewModal()});$('#cta-save')?.addEventListener('click',async()=>{try{await ensureDefaultLogo();}catch(_){ }generateAndUploadPDF()})}

document.addEventListener('DOMContentLoaded',init);

// Fallback: ensure preview buttons work even if init didn't bind
document.addEventListener('click', (e)=>{
  try{
    const t = e.target;
    const isBtn = (id)=> t && (t.id===id || (t.closest && t.closest('#'+id)));
    if (isBtn('btn-preview') || isBtn('cta-preview')){ e.preventDefault(); openPreviewModal(); }
  }catch(_){ }
});




function deleteTagPreset(tag){
  if(!tag) return false;
  const db = loadLocalTagPresets();
  if(!db[tag]) return false;
  delete db[tag];
  saveLocalTagPresets(db);
  try{
    const sel = document.querySelector('select[name="tagNo"]');
    if(sel){
      const opt = Array.from(sel.options).find(o => o.value === tag);
      if (opt) opt.remove();
      if (sel.value === tag) sel.value = '';
    }
  }catch(_){ }
  return true;
}













// Inline thumbnails update
function updateInlineThumbs(){
  try{
    ['photo1','photo2','photo3'].forEach(function(k){
      var el = document.getElementById('thumb-'+k);
      if(!el) return;
      var src = (state.images||{})[k];
      if(src){ el.src = src; el.style.display = 'block'; }
      else { try{ el.removeAttribute('src'); }catch(_){} el.style.display = 'none'; }
    });
  }catch(_){ }
}
