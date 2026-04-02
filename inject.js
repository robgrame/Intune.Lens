// ============================================================
// Intune Lens — Page-context script
// Monkey-patches fetch/XHR to capture the Bearer token that
// the Intune SPA sends to graph.microsoft.com, then relays it
// to the content script via postMessage.
// ============================================================
(function () {
  'use strict';

  const GRAPH = 'graph.microsoft.com';
  const MSG_TYPE = '__INTUNE_LENS_TOKEN__';
  const origin = window.location.origin;

  console.log('%c[IL-inject]', 'color:#0078d4;font-weight:bold', '🔑 Token interceptor loaded in MAIN world');

  function relay(token) {
    if (token) {
      window.postMessage({ type: MSG_TYPE, token }, origin);
      console.log('%c[IL-inject]', 'color:#0078d4;font-weight:bold', '🔑 Token captured and relayed');
    }
  }

  // --- fetch ---
  const _fetch = window.fetch;
  window.fetch = function (resource, init) {
    try {
      const url = typeof resource === 'string' ? resource : resource?.url || '';
      if (url.includes(GRAPH) && init?.headers) {
        let auth = null;
        const h = init.headers;
        if (h instanceof Headers)          auth = h.get('Authorization');
        else if (Array.isArray(h))         auth = (h.find(p => p[0]?.toLowerCase() === 'authorization') || [])[1];
        else if (typeof h === 'object')    auth = h['Authorization'] || h['authorization'];
        relay(auth);
      }
    } catch { /* never break the host page */ }
    return _fetch.apply(this, arguments);
  };

  // --- XMLHttpRequest ---
  const _open = XMLHttpRequest.prototype.open;
  const _setHeader = XMLHttpRequest.prototype.setRequestHeader;

  XMLHttpRequest.prototype.open = function (method, url) {
    this.__ilUrl = url;
    return _open.apply(this, arguments);
  };

  XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
    try {
      if (name.toLowerCase() === 'authorization' && this.__ilUrl?.includes(GRAPH)) {
        relay(value);
      }
    } catch { /* */ }
    return _setHeader.apply(this, arguments);
  };
})();
