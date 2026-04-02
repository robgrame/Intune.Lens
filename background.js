// ============================================================
// Intune Lens — Service Worker (Background)
// Token management, Graph API proxy, response caching
// ============================================================

const CACHE_TTL = 5 * 60 * 1000;       // 5 min cache per object
const TOKEN_MAX_AGE = 45 * 60 * 1000;  // consider token stale after 45 min
const MAX_CACHE_ENTRIES = 300;

const cache = new Map();

// ----------------------------------------------------------
// Token capture via webRequest (reads Authorization header
// from Intune portal's own Graph calls)
// ----------------------------------------------------------
chrome.webRequest.onSendHeaders.addListener(
  (details) => {
    if (!details.requestHeaders) return;
    const auth = details.requestHeaders.find(
      h => h.name.toLowerCase() === 'authorization'
    );
    if (auth?.value?.startsWith('Bearer ')) {
      const token = auth.value.substring(7);
      chrome.storage.session.set({
        authToken: token,
        tokenTimestamp: Date.now()
      });
    }
  },
  { urls: ['https://graph.microsoft.com/*'] },
  ['requestHeaders', 'extraHeaders']
);

// ----------------------------------------------------------
// Message handler — content script ↔ service worker
// ----------------------------------------------------------
chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  switch (msg.type) {

    // Token pushed from inject.js → content.js → here
    case 'setToken': {
      const token = (msg.token || '').replace(/^Bearer\s+/i, '');
      if (token) {
        chrome.storage.session.set({ authToken: token, tokenTimestamp: Date.now() });
      }
      sendResponse({ ok: true });
      return false;
    }

    // Graph API query
    case 'graphQuery': {
      handleGraphQuery(msg.endpoint, msg.cacheKey)
        .then(data  => sendResponse({ ok: true, data }))
        .catch(err  => sendResponse({ ok: false, error: err.message }));
      return true; // keep channel open for async response
    }

    // Extension status (used by popup)
    case 'getStatus': {
      chrome.storage.session.get(['authToken', 'tokenTimestamp'], (r) => {
        const hasToken = !!r.authToken && (Date.now() - (r.tokenTimestamp || 0) < TOKEN_MAX_AGE);
        sendResponse({
          hasToken,
          cacheSize: cache.size,
          tokenAgeSec: r.tokenTimestamp ? Math.round((Date.now() - r.tokenTimestamp) / 1000) : null
        });
      });
      return true;
    }

    // Flush cache
    case 'clearCache': {
      cache.clear();
      sendResponse({ ok: true });
      return false;
    }
  }
});

// ----------------------------------------------------------
// Graph API helper
// ----------------------------------------------------------
async function handleGraphQuery(endpoint, cacheKey) {
  const key = cacheKey || endpoint;

  // 1. Cache hit?
  const cached = cache.get(key);
  if (cached && Date.now() - cached.ts < CACHE_TTL) return cached.data;

  // 2. Retrieve token from session storage
  const { authToken, tokenTimestamp } = await chrome.storage.session.get([
    'authToken', 'tokenTimestamp'
  ]);
  if (!authToken || Date.now() - (tokenTimestamp || 0) > TOKEN_MAX_AGE) {
    throw new Error('No valid token. Browse the Intune portal so the extension can capture one.');
  }

  // 3. Call Graph
  const res = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
    headers: { Authorization: `Bearer ${authToken}` }
  });

  if (!res.ok) {
    if (res.status === 401) {
      await chrome.storage.session.remove(['authToken', 'tokenTimestamp']);
      throw new Error('Token expired — refresh the Intune portal page.');
    }
    throw new Error(`Graph ${res.status}: ${res.statusText}`);
  }

  const data = await res.json();

  // 4. Store in cache (and prune if oversized)
  cache.set(key, { data, ts: Date.now() });
  if (cache.size > MAX_CACHE_ENTRIES) pruneCache();

  return data;
}

function pruneCache() {
  const now = Date.now();
  for (const [k, v] of cache) {
    if (now - v.ts > CACHE_TTL) cache.delete(k);
  }
  // If still too big, drop oldest half
  if (cache.size > MAX_CACHE_ENTRIES) {
    const entries = [...cache.entries()].sort((a, b) => a[1].ts - b[1].ts);
    const toDrop = Math.floor(entries.length / 2);
    for (let i = 0; i < toDrop; i++) cache.delete(entries[i][0]);
  }
}
