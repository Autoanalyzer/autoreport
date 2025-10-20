// PWA bootstrap: register service worker and expose Install action
(function(){
  // Register Service Worker
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', function(){
      navigator.serviceWorker.register('./service-worker.js').catch(()=>{});
    });
  }

  // Create install buttons on demand
  let deferredPrompt = null;
  function isInstalled(){
    try{
      if (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches) return true;
      if (window.navigator.standalone) return true; // iOS Safari
    }catch(_){ }
    return false;
  }
  function showInstallHelp(){
    const toast = (msg)=>{
      try{ if (typeof window.showToast === 'function'){ window.showToast(msg); return; } }catch(_){ }
      try{
        let t = document.getElementById('install-hint-toast');
        if (!t){
          t = document.createElement('div');
          t.id = 'install-hint-toast';
          t.style.cssText = 'position:fixed;left:50%;bottom:14px;transform:translateX(-50%);background:rgba(30,41,59,.92);color:#fff;padding:10px 12px;border-radius:10px;font-size:13px;z-index:9999;box-shadow:0 6px 18px rgba(0,0,0,.25)';
          document.body.appendChild(t);
        }
        t.textContent = msg; t.style.display = 'block';
        clearTimeout(t._timer); t._timer = setTimeout(()=>{ t.style.display='none'; }, 4200);
      }catch(_){ /* no-op */ }
    };
    try{
      const ua = navigator.userAgent || '';
      const isIOS = /iPhone|iPad|iPod/i.test(ua);
      const isAndroid = /Android/i.test(ua);
      if (isIOS) {
        toast('บน iPhone/Safari: แตะปุ่มแชร์ แล้วเลือก Add to Home Screen');
      } else if (isAndroid) {
        toast('Android: เปิดเมนู ⋮ แล้วเลือก Add to Home screen / Install app');
      } else {
        toast('แนะนำติดตั้งบนมือถือ (Android/Chrome หรือ iPhone/Safari)');
      }
    }catch(_){ /* ignore */ }
  }
  function ensureButtons(){
    // Where to place the install button
    const targets = [
      document.querySelector('#form-actions'),
      document.querySelector('.app-header .actions')
    ].filter(Boolean);
    targets.forEach((container)=>{
      if (!document.getElementById('btn-install')){
        const btn = document.createElement('button');
        btn.id = 'btn-install';
        btn.type = 'button';
        btn.textContent = 'ติดตั้งแอป';
        btn.style.display = 'none';
        container.appendChild(btn);
        btn.addEventListener('click', async () => {
          if (!deferredPrompt) return;
          btn.disabled = true;
          deferredPrompt.prompt();
          await deferredPrompt.userChoice.catch(()=>{});
          deferredPrompt = null; // reset
          btn.style.display = 'none';
          btn.disabled = false;
        });
      }
    });
    // Mobile CTA button
    const mobileCta = document.getElementById('mobile-cta');
    if (mobileCta && !mobileCta.querySelector('#cta-install')){
      const btn = document.createElement('button');
      btn.id = 'cta-install';
      btn.type = 'button';
      btn.className = 'ghost';
      btn.textContent = 'ติดตั้ง';
      btn.style.display = 'none';
      mobileCta.appendChild(btn);
      btn.addEventListener('click', async () => {
        if (!deferredPrompt) return;
        btn.disabled = true;
        deferredPrompt.prompt();
        await deferredPrompt.userChoice.catch(()=>{});
        deferredPrompt = null;
        btn.style.display = 'none';
        btn.disabled = false;
      });
    }

    // Landing page button (homepage)
    const landingHead = document.querySelector('.landing-head');
    if (landingHead && !document.getElementById('landing-install')){
      const b = document.createElement('button');
      b.id = 'landing-install';
      b.type = 'button';
      b.className = 'ghost';
      b.innerHTML = '<i class="fa-solid fa-download" aria-hidden="true"></i> <span>ติดตั้งแอป</span>';
      b.style.display = isInstalled() ? 'none' : '';
      b.style.marginTop = '8px';
      landingHead.appendChild(b);
      b.addEventListener('click', async ()=>{
        if (!deferredPrompt) return;
        b.disabled = true;
        deferredPrompt.prompt();
        await deferredPrompt.userChoice.catch(()=>{});
        deferredPrompt = null;
        b.style.display = 'none';
        b.disabled = false;
      });
      // Fallback instructions when prompt is not available
      b.addEventListener('click', ()=>{ if (!deferredPrompt) showInstallHelp(); }, { capture:false });
    }
  }

  ensureButtons();

  window.addEventListener('beforeinstallprompt', (e) => {
    // Prevent the mini-info bar from appearing
    e.preventDefault();
    deferredPrompt = e;
    ensureButtons();
    // Show buttons
    const b1 = document.getElementById('btn-install');
    const b2 = document.getElementById('cta-install');
    const b3 = document.getElementById('landing-install');
    if (b1) b1.style.display = '';
    if (b2) b2.style.display = '';
    if (b3 && !isInstalled()) b3.style.display = '';
  });

  window.addEventListener('appinstalled', () => {
    const b1 = document.getElementById('btn-install');
    const b2 = document.getElementById('cta-install');
    const b3 = document.getElementById('landing-install');
    if (b1) b1.style.display = 'none';
    if (b2) b2.style.display = 'none';
    if (b3) b3.style.display = 'none';
    deferredPrompt = null;
  });
})();
