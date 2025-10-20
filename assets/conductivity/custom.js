(function(){
  function onReady(fn){ if(document.readyState==='loading'){ document.addEventListener('DOMContentLoaded', fn, {once:true}); } else { fn(); } }

  onReady(function(){
    try{
      const sol = document.querySelector('[name="solutionType"]');
      if (sol && (!sol.value || String(sol.value).trim()==='')){ sol.value='Buffer'; }
    }catch(_){ }
  });
})();

// Show inline photo previews when attaching images
(function(){
  function ready(fn){ if(document.readyState==='loading'){ document.addEventListener('DOMContentLoaded', fn, {once:true}); } else { fn(); } }
  ready(function(){
    try{
      ['photo1','photo2','photo3'].forEach(function(name){
        var input = document.querySelector('[name="'+name+'"]');
        var img = document.getElementById('thumb-'+name);
        if(!input || !img) return;
        input.addEventListener('change', function(){
          var f = input.files && input.files[0];
          if(!f){ img.removeAttribute('src'); img.style.display='none'; return; }
          var r = new FileReader();
          r.onload = function(){ img.src = String(r.result); img.style.display='block'; };
          r.readAsDataURL(f);
        });
      });
    }catch(_){ }
  });
})();
