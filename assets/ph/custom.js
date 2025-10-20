// pH certificate tweaks and safe defaults live here
(function(){
  function onReady(fn){ if(document.readyState==='loading'){ document.addEventListener('DOMContentLoaded', fn, {once:true}); } else { fn(); } }

  onReady(function(){
    try{
      // Ensure Solution Type default if empty
      const sol = document.querySelector('[name="solutionType"]');
      if (sol && (!sol.value || String(sol.value).trim()==='')){ sol.value='Buffer'; }
    }catch(_){ }

    try{
      // Fill Brand&Model defaults in editable buffer table if empty
      const rows = document.querySelectorAll('#buffer-table tbody tr');
      rows.forEach(tr => {
        const input = tr.querySelector('input[data-k="brandModel"]');
        if (input && (!input.value || String(input.value).trim()==='')){ input.value='METTLER TOLEDO'; }
      });
    }catch(_){ }

    try{
      // Keep title centered regardless of layout
      const title = document.querySelector('.cert-header .cert-title');
      if (title){ title.style.textAlign='center'; }
    }catch(_){ }
  });
})();

