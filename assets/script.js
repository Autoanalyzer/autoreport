// Compatibility loader: the pH report code moved to assets/ph/ph.js
// This file forwards to the new location to avoid breaking old references.
(function(){
  var s = document.createElement('script');
  s.src = 'assets/ph/ph.js';
  s.defer = true;
  document.currentScript && document.currentScript.parentNode
    ? document.currentScript.parentNode.insertBefore(s, document.currentScript)
    : document.head.appendChild(s);
})();

