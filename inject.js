// ============================================================
// Intune Lens v1.3.0 — Page-context script (MAIN world)
// 1. Captures Bearer tokens from Graph API requests
// 2. Intercepts Graph API RESPONSES via Response.prototype.json
//    to extract object data (immune to fetch wrapper conflicts)
// ============================================================
(function () {
  'use strict';

  const GRAPH = 'graph.microsoft.com';
  const TAG = '%c[IL-inject]';
  const CSS = 'color:#0078d4;font-weight:bold';

  console.log(TAG, CSS, '🔑 v1.3.0 interceptor loaded (MAIN world)');

  // --- Token relay ---
  let relayCount = 0;
  function relayToken(token) {
    if (!token) return;
    window.postMessage({ type: '__INTUNE_LENS_TOKEN__', token }, '*');
    if (++relayCount <= 3) console.log(TAG, CSS, `🔑 Token relayed (#${relayCount})`);
  }

  // --- Object data extraction ---
  let dataRelayCount = 0;
  function extractObjects(items) {
    if (!Array.isArray(items) || items.length === 0) return;
    const objects = [];
    for (const item of items) {
      if (!item.id) continue;
      if (item.deviceName !== undefined)
        objects.push({ ...item, _t: 'device' });
      else if (item.userPrincipalName !== undefined && item.displayName !== undefined)
        objects.push({ ...item, _t: 'user' });
      else if (item.publisher !== undefined)
        objects.push({ ...item, _t: 'app' });
    }
    if (objects.length > 0) {
      dataRelayCount++;
      console.log(TAG, CSS, `📦 Captured ${objects.length} objects (batch #${dataRelayCount})`);
      window.postMessage({ type: '__INTUNE_LENS_DATA__', objects }, '*');
    }
  }

  // ==========================================================
  // PRIMARY: Intercept Response.prototype.json()
  // This works regardless of how many wrappers exist on fetch.
  // ==========================================================
  const _json = Response.prototype.json;
  Response.prototype.json = function () {
    const url = this.url || '';
    const promise = _json.call(this);

    if (url.includes(GRAPH)) {
      promise.then(data => {
        // Extract token from the fact that this response succeeded
        // (token was already captured via webRequest, this is supplementary)
        if (data?.value) extractObjects(data.value);
      }).catch(() => {});
    }

    return promise;
  };

  // Also intercept .text() in case code does JSON.parse(await res.text())
  const _text = Response.prototype.text;
  Response.prototype.text = function () {
    const url = this.url || '';
    const promise = _text.call(this);

    if (url.includes(GRAPH)) {
      promise.then(text => {
        try {
          const data = JSON.parse(text);
          if (data?.value) extractObjects(data.value);
        } catch { /* not JSON */ }
      }).catch(() => {});
    }

    return promise;
  };

  console.log(TAG, CSS, '📡 Response.prototype.json/text interceptors active');

  // ==========================================================
  // SECONDARY: fetch wrapper (token capture + backup data capture)
  // May be overridden by Azure portal — that's OK, primary above works
  // ==========================================================
  const _fetch = window.fetch;
  window.fetch = function (resource, init) {
    try {
      let url = '';
      let headers = null;

      if (typeof resource === 'string') {
        url = resource;
        headers = init?.headers;
      } else if (resource instanceof Request) {
        url = resource.url;
        headers = resource.headers;
      }

      if (url.includes(GRAPH) && headers) {
        let auth = null;
        if (headers instanceof Headers)       auth = headers.get('Authorization');
        else if (Array.isArray(headers))      auth = (headers.find(p => p[0]?.toLowerCase() === 'authorization') || [])[1];
        else if (typeof headers === 'object') auth = headers['Authorization'] || headers['authorization'];
        relayToken(auth);
      }
    } catch { /* never break the host page */ }

    return _fetch.apply(this, arguments);
  };

  // ==========================================================
  // XHR interceptors (token + response data)
  // ==========================================================
  const _open = XMLHttpRequest.prototype.open;
  const _setHeader = XMLHttpRequest.prototype.setRequestHeader;
  const _send = XMLHttpRequest.prototype.send;

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
