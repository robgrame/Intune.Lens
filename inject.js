// ============================================================
// Intune Lens — Page-context script (MAIN world)
// 1. Captures Bearer tokens from Graph API requests
// 2. Intercepts Graph API RESPONSES to extract object data
//    (device names, user names, IDs) for DOM matching
// ============================================================
(function () {
  'use strict';

  const GRAPH = 'graph.microsoft.com';
  const TAG = '%c[IL-inject]';
  const CSS = 'color:#0078d4;font-weight:bold';

  console.log(TAG, CSS, '🔑 Token + data interceptor loaded in MAIN world');

  // --- Token relay ---
  let relayCount = 0;
  function relayToken(token) {
    if (!token) return;
    window.postMessage({ type: '__INTUNE_LENS_TOKEN__', token }, '*');
    if (++relayCount <= 3) console.log(TAG, CSS, `🔑 Token relayed (#${relayCount})`);
  }

  // --- Object data extraction from Graph responses ---
  function extractObjects(items) {
    if (!Array.isArray(items) || items.length === 0) return;
    const objects = [];
    for (const item of items) {
      if (!item.id) continue;
      // Detect type by characteristic fields
      if (item.deviceName !== undefined)         objects.push({ ...item, _t: 'device' });
      else if (item.userPrincipalName !== undefined && item.displayName !== undefined)
                                                  objects.push({ ...item, _t: 'user' });
      else if (item.publisher !== undefined)      objects.push({ ...item, _t: 'app' });
    }
    if (objects.length > 0) {
      console.log(TAG, CSS, `📦 Captured ${objects.length} objects from Graph response`);
      window.postMessage({ type: '__INTUNE_LENS_DATA__', objects }, '*');
    }
  }

  // --- fetch intercept ---
  const _fetch = window.fetch;
  window.fetch = function (resource, init) {
    let url = '';
    try {
      url = typeof resource === 'string' ? resource : resource?.url || '';
      if (url.includes(GRAPH) && init?.headers) {
        let auth = null;
        const h = init.headers;
        if (h instanceof Headers)          auth = h.get('Authorization');
        else if (Array.isArray(h))         auth = (h.find(p => p[0]?.toLowerCase() === 'authorization') || [])[1];
        else if (typeof h === 'object')    auth = h['Authorization'] || h['authorization'];
        relayToken(auth);
      }
    } catch { /* never break the host page */ }

    const result = _fetch.apply(this, arguments);

    // Tap into Graph responses (non-blocking)
    if (url.includes(GRAPH)) {
      result.then(r => {
        const ct = r.headers?.get('content-type') || '';
        if (!ct.includes('json')) return;
        return r.clone().json();
      }).then(data => {
        if (data?.value) extractObjects(data.value);
      }).catch(() => {});
    }

    return result;
  };

  // --- XMLHttpRequest intercept ---
  const _open = XMLHttpRequest.prototype.open;
  const _setHeader = XMLHttpRequest.prototype.setRequestHeader;

  XMLHttpRequest.prototype.open = function (method, url) {
    this.__ilUrl = url;
    return _open.apply(this, arguments);
  };

  XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
    try {
      if (name.toLowerCase() === 'authorization' && this.__ilUrl?.includes(GRAPH)) {
        relayToken(value);
      }
    } catch { /* */ }
    return _setHeader.apply(this, arguments);
  };

  // Also intercept XHR responses
  const _send = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.send = function () {
    if (this.__ilUrl?.includes(GRAPH)) {
      this.addEventListener('load', function () {
        try {
          const ct = this.getResponseHeader('content-type') || '';
          if (!ct.includes('json')) return;
          const data = JSON.parse(this.responseText);
          if (data?.value) extractObjects(data.value);
        } catch { /* */ }
      });
    }
    return _send.apply(this, arguments);
  };
})();
