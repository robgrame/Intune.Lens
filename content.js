// ============================================================
// Intune Lens — Content Script
// Detects Intune objects in the DOM, shows rich hover cards
// with data fetched from Microsoft Graph.
// ============================================================
(function () {
  'use strict';

  // ==========================================================
  // Constants & Config
  // ==========================================================
  const PROCESSED = 'data-il';
  const ID_ATTR   = 'data-il-id';
  const SCAN_DEBOUNCE = 500;
  const GUID = '[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}';
  const GUID_RE = new RegExp(GUID, 'i');

  // Debug logging — check browser console filtered by "[IL]"
  const DEBUG = true;
  function log(...args) { if (DEBUG) console.log('%c[IL]', 'color:#0078d4;font-weight:bold', ...args); }
  function warn(...args) { if (DEBUG) console.warn('%c[IL]', 'color:#d83b01;font-weight:bold', ...args); }

  // Intune hash-route patterns → object type
  const ROUTE_PATTERNS = [
    { type: 'device', re: new RegExp(`mdmDeviceId/(${GUID})`, 'i') },
    { type: 'user',   re: new RegExp(`userId/(${GUID})`, 'i') },
    { type: 'app',    re: new RegExp(`appId/(${GUID})`, 'i') },
    // Compliance & configuration policies
    { type: 'policy', re: new RegExp(`policyId/(${GUID})`, 'i') },
    { type: 'policy', re: new RegExp(`configurationId/(${GUID})`, 'i') },
    // Broad fallbacks using blade names
    { type: 'device', re: new RegExp(`DeviceSettingsMenuBlade[^)]*?(${GUID})`, 'i') },
    { type: 'user',   re: new RegExp(`UserProfileMenuBlade[^)]*?(${GUID})`, 'i') },
    { type: 'app',    re: new RegExp(`MobileAppMenuBlade[^)]*?(${GUID})`, 'i') },
  ];

  let settings = {
    enabled: true,
    showDeviceCards: true,
    showUserCards: true,
    showAppCards: true,
    showPolicyCards: true,
    hoverDelay: 400
  };

  let currentCard = null;
  let hoverTimer  = null;
  let hideTimer   = null;

  // ==========================================================
  // Token bridge — inject.js runs in MAIN world via manifest,
  // we just listen for its postMessage here
  // ==========================================================
  function setupTokenBridge() {
    window.addEventListener('message', (e) => {
      if (e.data?.type === '__INTUNE_LENS_TOKEN__') {
        log('Token captured from page ✓');
        chrome.runtime.sendMessage({ type: 'setToken', token: e.data.token });
      }
    });
    log('Token bridge listener ready');
  }

  // ==========================================================
  // Graph query (delegates to service worker)
  // ==========================================================
  function graphQuery(endpoint, cacheKey) {
    return new Promise((resolve, reject) => {
      chrome.runtime.sendMessage(
        { type: 'graphQuery', endpoint, cacheKey },
        (r) => {
          if (chrome.runtime.lastError) return reject(new Error(chrome.runtime.lastError.message));
          r?.ok ? resolve(r.data) : reject(new Error(r?.error || 'Graph query failed'));
        }
      );
    });
  }

  // ==========================================================
  // Data fetchers
  // ==========================================================
  const DEVICE_SELECT = [
    'deviceName','operatingSystem','osVersion','model','manufacturer',
    'serialNumber','lastSyncDateTime','complianceState','enrolledDateTime',
    'userPrincipalName','managementAgent','deviceEnrollmentType',
    'isEncrypted','isSupervised','totalStorageSpaceInBytes','freeStorageSpaceInBytes'
  ].join(',');

  async function fetchDevice(id) {
    return graphQuery(`/deviceManagement/managedDevices/${id}?$select=${DEVICE_SELECT}`, `dev:${id}`);
  }

  async function fetchUser(id) {
    const user = await graphQuery(
      `/users/${id}?$select=displayName,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,mail,accountEnabled,createdDateTime`,
      `usr:${id}`
    );
    try {
      const devs = await graphQuery(
        `/users/${id}/managedDevices?$select=deviceName,operatingSystem,complianceState&$top=10`,
        `usr-dev:${id}`
      );
      user._devices = devs.value || [];
    } catch { user._devices = []; }
    return user;
  }

  async function fetchApp(id) {
    return graphQuery(
      `/deviceAppManagement/mobileApps/${id}?$select=displayName,publisher,description,createdDateTime,lastModifiedDateTime`,
      `app:${id}`
    );
  }

  async function fetchPolicy(id) {
    try {
      return await graphQuery(
        `/deviceManagement/deviceCompliancePolicies/${id}?$select=displayName,description,createdDateTime,lastModifiedDateTime,version`,
        `pol:${id}`
      );
    } catch {
      return graphQuery(
        `/deviceManagement/deviceConfigurations/${id}?$select=displayName,description,createdDateTime,lastModifiedDateTime,version`,
        `pol:${id}`
      );
    }
  }

  // ==========================================================
  // Utility helpers
  // ==========================================================
  function esc(s) { const d = document.createElement('div'); d.textContent = s ?? ''; return d.innerHTML; }

  function ago(iso) {
    if (!iso) return 'N/A';
    const ms = Date.now() - new Date(iso).getTime();
    if (ms < 60000)       return 'Just now';
    if (ms < 3600000)     return `${Math.floor(ms / 60000)}m ago`;
    if (ms < 86400000)    return `${Math.floor(ms / 3600000)}h ago`;
    if (ms < 2592000000)  return `${Math.floor(ms / 86400000)}d ago`;
    return new Date(iso).toLocaleDateString();
  }

  function bytes(b) { return b ? `${(b / 1073741824).toFixed(1)} GB` : 'N/A'; }

  function row(label, value) {
    return `<div class="il-row"><span class="il-lbl">${esc(label)}</span><span class="il-val">${esc(String(value ?? '—'))}</span></div>`;
  }

  const COMPLIANCE = {
    compliant:      { cls: 'ok',   txt: '✓ Compliant' },
    noncompliant:   { cls: 'bad',  txt: '✗ Non-Compliant' },
    unknown:        { cls: 'unk',  txt: '? Unknown' },
    notApplicable:  { cls: 'unk',  txt: '— N/A' },
    inGracePeriod:  { cls: 'warn', txt: '⏳ Grace Period' },
    configManager:  { cls: 'unk',  txt: 'ConfigMgr' },
  };
  function badge(state) { return COMPLIANCE[state] || COMPLIANCE.unknown; }

  // ==========================================================
  // Card templates
  // ==========================================================
  function deviceCard(d) {
    const c = badge(d.complianceState);
    const storage = d.totalStorageSpaceInBytes
      ? `${bytes(d.totalStorageSpaceInBytes - (d.freeStorageSpaceInBytes || 0))} / ${bytes(d.totalStorageSpaceInBytes)}`
      : null;
    return `
      <div class="il-hdr">
        <span class="il-ico">💻</span>
        <div class="il-ttl-grp">
          <div class="il-ttl">${esc(d.deviceName)}</div>
          <div class="il-sub">${esc(d.model || '')}${d.manufacturer ? ' · ' + esc(d.manufacturer) : ''}</div>
        </div>
        <span class="il-badge ${c.cls}">${c.txt}</span>
      </div>
      <div class="il-body">
        <div class="il-sec">
          <div class="il-sec-ttl">Device Info</div>
          ${row('OS', `${d.operatingSystem || '—'} ${d.osVersion || ''}`)}
          ${row('Serial', d.serialNumber)}
          ${row('Encryption', d.isEncrypted ? '🔒 Encrypted' : '🔓 Not encrypted')}
          ${row('Management', d.managementAgent)}
          ${storage ? row('Storage', storage) : ''}
        </div>
        <hr class="il-div">
        <div class="il-sec">
          <div class="il-sec-ttl">User & Activity</div>
          ${row('Primary User', d.userPrincipalName)}
          ${row('Last Check-in', ago(d.lastSyncDateTime))}
          ${row('Enrolled', ago(d.enrolledDateTime))}
          ${row('Enrollment Type', d.deviceEnrollmentType)}
        </div>
      </div>
      <div class="il-foot"><span class="il-tag">DEVICE</span><span class="il-brand">Intune Lens</span></div>`;
  }

  function userCard(u) {
    const devItems = (u._devices || []).slice(0, 5).map(d => {
      const c = badge(d.complianceState);
      return `<div class="il-dev-item"><span class="il-dot ${c.cls}"></span>${esc(d.deviceName)} <span class="il-dev-os">${d.operatingSystem || ''}</span></div>`;
    }).join('');
    return `
      <div class="il-hdr">
        <span class="il-ico">👤</span>
        <div class="il-ttl-grp">
          <div class="il-ttl">${esc(u.displayName)}</div>
          <div class="il-sub">${esc(u.userPrincipalName)}</div>
        </div>
        <span class="il-badge ${u.accountEnabled ? 'ok' : 'bad'}">${u.accountEnabled ? '✓ Enabled' : '✗ Disabled'}</span>
      </div>
      <div class="il-body">
        <div class="il-sec">
          <div class="il-sec-ttl">Profile</div>
          ${row('Job Title', u.jobTitle)}
          ${row('Department', u.department)}
          ${row('Office', u.officeLocation)}
          ${row('Phone', u.mobilePhone)}
          ${row('Email', u.mail)}
        </div>
        ${devItems ? `<hr class="il-div"><div class="il-sec"><div class="il-sec-ttl">Managed Devices (${u._devices.length})</div><div class="il-dev-list">${devItems}</div></div>` : ''}
      </div>
      <div class="il-foot"><span class="il-tag">USER</span><span class="il-brand">Intune Lens</span></div>`;
  }

  function appCard(a) {
    const desc = a.description
      ? `<div class="il-desc">${esc(a.description.substring(0, 200))}${a.description.length > 200 ? '…' : ''}</div>` : '';
    return `
      <div class="il-hdr">
        <span class="il-ico">📱</span>
        <div class="il-ttl-grp">
          <div class="il-ttl">${esc(a.displayName)}</div>
          <div class="il-sub">${esc(a.publisher || '')}</div>
        </div>
      </div>
      <div class="il-body">
        <div class="il-sec">
          <div class="il-sec-ttl">App Info</div>
          ${row('Publisher', a.publisher)}
          ${row('Created', ago(a.createdDateTime))}
          ${row('Modified', ago(a.lastModifiedDateTime))}
          ${desc}
        </div>
      </div>
      <div class="il-foot"><span class="il-tag">APP</span><span class="il-brand">Intune Lens</span></div>`;
  }

  function policyCard(p) {
    const desc = p.description
      ? `<div class="il-desc">${esc(p.description.substring(0, 300))}${p.description.length > 300 ? '…' : ''}</div>` : '';
    return `
      <div class="il-hdr">
        <span class="il-ico">📋</span>
        <div class="il-ttl-grp">
          <div class="il-ttl">${esc(p.displayName)}</div>
          <div class="il-sub">Version ${p.version ?? '—'}</div>
        </div>
      </div>
      <div class="il-body">
        <div class="il-sec">
          <div class="il-sec-ttl">Policy Info</div>
          ${row('Created', ago(p.createdDateTime))}
          ${row('Modified', ago(p.lastModifiedDateTime))}
          ${desc}
        </div>
      </div>
      <div class="il-foot"><span class="il-tag">POLICY</span><span class="il-brand">Intune Lens</span></div>`;
  }

  function loadingCard(type) {
    return `
      <div class="il-hdr">
        <span class="il-ico il-spin">⟳</span>
        <div class="il-ttl-grp">
          <div class="il-ttl">Loading ${esc(type)}…</div>
          <div class="il-sub">Fetching from Microsoft Graph</div>
        </div>
      </div>
      <div class="il-body"><div class="il-bar"></div></div>`;
  }

  function errorCard(msg) {
    return `
      <div class="il-hdr">
        <span class="il-ico">⚠️</span>
        <div class="il-ttl-grp">
          <div class="il-ttl">Error</div>
          <div class="il-sub">${esc(msg)}</div>
        </div>
      </div>`;
  }

  // ==========================================================
  // Card lifecycle (show / hide / position)
  // ==========================================================
  function ensureContainer() {
    let c = document.getElementById('il-root');
    if (!c) { c = document.createElement('div'); c.id = 'il-root'; document.body.appendChild(c); }
    return c;
  }

  function showCard(anchor, type, id) {
    clearTimeout(hideTimer);

    const container = ensureContainer();
    const card = document.createElement('div');
    card.className = 'il-card il-enter';
    card.innerHTML = loadingCard(type);

    // Position relative to anchor
    const rect = anchor.getBoundingClientRect();
    const W = 380, MARGIN = 8;
    let left = rect.left + window.scrollX;
    let top  = rect.bottom + window.scrollY + MARGIN;

    if (left + W > window.innerWidth)  left = window.innerWidth - W - MARGIN;
    if (left < MARGIN)                 left = MARGIN;
    if (rect.bottom + 320 > window.innerHeight) {
      top = rect.top + window.scrollY - 320 - MARGIN;
      if (top < 0) top = rect.bottom + window.scrollY + MARGIN;
    }

    card.style.left = `${left}px`;
    card.style.top  = `${top}px`;

    hideImmediate();
    container.appendChild(card);
    currentCard = card;

    card.addEventListener('mouseenter', () => clearTimeout(hideTimer));
    card.addEventListener('mouseleave', scheduleHide);

    requestAnimationFrame(() => card.classList.remove('il-enter'));

    // Fetch & render
    (async () => {
      try {
        const fetchers = { device: fetchDevice, user: fetchUser, app: fetchApp, policy: fetchPolicy };
        const renderers = { device: deviceCard, user: userCard, app: appCard, policy: policyCard };
        const data = await fetchers[type](id);
        if (card === currentCard) card.innerHTML = renderers[type](data);
      } catch (err) {
        if (card === currentCard) card.innerHTML = errorCard(err.message);
      }
    })();
  }

  function scheduleHide() { clearTimeout(hideTimer); hideTimer = setTimeout(hideImmediate, 300); }

  function hideImmediate() {
    if (!currentCard) return;
    currentCard.classList.add('il-exit');
    const ref = currentCard;
    setTimeout(() => ref.remove(), 200);
    currentCard = null;
  }

  // ==========================================================
  // DOM detection — find Intune object links
  // ==========================================================
  function detectFromText(text) {
    if (!text) return null;
    for (const p of ROUTE_PATTERNS) {
      const m = text.match(p.re);
      if (m) return { type: p.type, id: m[1] };
    }
    return null;
  }

  function detect(el) {
    // 1. href
    const href = el.getAttribute('href') || '';
    const r = detectFromText(href);
    if (r) return r;

    // 2. data-* and aria attributes
    for (const attr of el.attributes) {
      if (['class', 'style', 'id', PROCESSED, ID_ATTR].includes(attr.name)) continue;
      const r2 = detectFromText(attr.value);
      if (r2) return r2;
    }

    return null;
  }

  function scan() {
    if (!settings.enabled) return;

    // Very broad selector: links, clickable roles, table rows, anything with data-* GUID
    const candidates = document.querySelectorAll(
      `a[href]:not([${PROCESSED}]),
       [role="link"]:not([${PROCESSED}]),
       [role="row"]:not([${PROCESSED}]),
       [role="gridcell"]:not([${PROCESSED}]),
       [data-href]:not([${PROCESSED}]),
       [data-automationid]:not([${PROCESSED}])`
    );

    let found = 0;
    for (const el of candidates) {
      const obj = detect(el);
      if (!obj) continue;
      const key = `show${obj.type[0].toUpperCase()}${obj.type.slice(1)}Cards`;
      if (settings[key] === false) continue;

      el.setAttribute(PROCESSED, obj.type);
      el.setAttribute(ID_ATTR, obj.id);
      el.classList.add('il-link');
      el.addEventListener('mouseenter', onEnter);
      el.addEventListener('mouseleave', onLeave);
      found++;
    }

    if (found > 0) log(`Scan: tagged ${found} new object elements`);

    // Debug: dump DOM info on first scan
    if (!scan._dbg) {
      scan._dbg = true;
      const allA = document.querySelectorAll('a[href]');
      const sampleHrefs = [...allA].slice(0, 30).map(a => a.getAttribute('href'));
      log(`DOM has ${allA.length} <a> tags. Sample hrefs:`, sampleHrefs);
      log('Hash:', location.hash);
      log('Pathname:', location.pathname);

      // Dump unique role values to help find clickable elements
      const roles = new Set();
      document.querySelectorAll('[role]').forEach(el => roles.add(el.getAttribute('role')));
      log('ARIA roles found on page:', [...roles]);
    }
  }

  // ==========================================================
  // FAB — Floating Action Button for current-page detection
  // Parses the URL hash to detect objects on detail pages.
  // This works regardless of DOM structure.
  // ==========================================================
  let currentFab = null;
  let lastFabHash = '';

  const FAB_PATTERNS = [
    { type: 'device', re: /mdmDeviceId\/([0-9a-f-]{36})/i,           icon: '💻', label: 'Device' },
    { type: 'user',   re: /userId\/([0-9a-f-]{36})/i,                 icon: '👤', label: 'User' },
    { type: 'app',    re: /appId\/([0-9a-f-]{36})/i,                  icon: '📱', label: 'App' },
    { type: 'policy', re: /policyId\/([0-9a-f-]{36})/i,               icon: '📋', label: 'Policy' },
    { type: 'policy', re: /configurationId\/([0-9a-f-]{36})/i,        icon: '📋', label: 'Policy' },
    { type: 'device', re: /DeviceSettingsMenuBlade[^]*?([0-9a-f-]{36})/i, icon: '💻', label: 'Device' },
    { type: 'user',   re: /UserProfileMenuBlade[^]*?([0-9a-f-]{36})/i,    icon: '👤', label: 'User' },
    { type: 'app',    re: /MobileAppMenuBlade[^]*?([0-9a-f-]{36})/i,      icon: '📱', label: 'App' },
  ];

  function updateFab() {
    const hash = location.hash || '';
    if (hash === lastFabHash) return;
    lastFabHash = hash;

    // Remove old FAB
    if (currentFab) { currentFab.remove(); currentFab = null; }

    if (!settings.enabled) return;

    // Try to detect object from URL
    for (const p of FAB_PATTERNS) {
      const m = hash.match(p.re);
      if (!m) continue;

      const id = m[1];
      log(`FAB: detected ${p.type} ${id} from URL hash`);

      const fab = document.createElement('div');
      fab.id = 'il-fab';
      fab.innerHTML = `<span class="il-fab-icon">${p.icon}</span><span class="il-fab-text">🔍 ${p.label} Info</span>`;
      fab.title = `Show ${p.label} details (Intune Lens)`;
      fab.addEventListener('click', () => {
        log(`FAB clicked → showing ${p.type} card for ${id}`);
        showCardFixed(p.type, id);
      });
      document.body.appendChild(fab);
      currentFab = fab;
      return;
    }

    log('FAB: no object detected in URL hash:', hash.substring(0, 120));
  }

  // Show card at a fixed position (for FAB clicks)
  function showCardFixed(type, id) {
    hideImmediate();
    const container = ensureContainer();
    const card = document.createElement('div');
    card.className = 'il-card il-enter il-card-fixed';
    card.innerHTML = loadingCard(type);

    // Fixed position: right side, vertically centered
    card.style.position = 'fixed';
    card.style.right = '20px';
    card.style.top = '80px';
    card.style.left = 'auto';

    container.appendChild(card);
    currentCard = card;

    // Close button
    const close = document.createElement('button');
    close.className = 'il-close';
    close.innerHTML = '✕';
    close.title = 'Close';
    close.addEventListener('click', (e) => { e.stopPropagation(); hideImmediate(); });
    card.prepend(close);

    card.addEventListener('mouseenter', () => clearTimeout(hideTimer));
    card.addEventListener('mouseleave', scheduleHide);
    requestAnimationFrame(() => card.classList.remove('il-enter'));

    // Fetch & render
    (async () => {
      try {
        const fetchers = { device: fetchDevice, user: fetchUser, app: fetchApp, policy: fetchPolicy };
        const renderers = { device: deviceCard, user: userCard, app: appCard, policy: policyCard };
        const data = await fetchers[type](id);
        if (card === currentCard) {
          card.innerHTML = renderers[type](data);
          card.prepend(close);  // re-add close button after innerHTML replace
        }
      } catch (err) {
        if (card === currentCard) {
          card.innerHTML = errorCard(err.message);
          card.prepend(close);
        }
      }
    })();
  }

  function onEnter(e) {
    const el = e.currentTarget;
    const type = el.getAttribute(PROCESSED);
    const id = el.getAttribute(ID_ATTR);
    log(`Hover → ${type} ${id}`);
    clearTimeout(hoverTimer);
    clearTimeout(hideTimer);
    hoverTimer = setTimeout(
      () => showCard(el, type, id),
      settings.hoverDelay || 400
    );
  }

  function onLeave() {
    clearTimeout(hoverTimer);
    scheduleHide();
  }

  // ==========================================================
  // MutationObserver — react to SPA navigation
  // ==========================================================
  let scanTimer = null;

  function setupObserver() {
    new MutationObserver(() => {
      clearTimeout(scanTimer);
      scanTimer = setTimeout(scan, SCAN_DEBOUNCE);
    }).observe(document.body, { childList: true, subtree: true });

    window.addEventListener('hashchange', () => {
      log('Hash changed →', location.hash.substring(0, 100));
      hideImmediate();
      updateFab();
      clearTimeout(scanTimer);
      scan._dbg = false; // re-dump debug on new page
      scanTimer = setTimeout(scan, SCAN_DEBOUNCE);
    });

    // Also poll for URL changes (some SPAs don't fire hashchange)
    let lastUrl = location.href;
    setInterval(() => {
      if (location.href !== lastUrl) {
        lastUrl = location.href;
        log('URL changed (poll) →', location.href.substring(0, 120));
        updateFab();
        clearTimeout(scanTimer);
        scan._dbg = false;
        scanTimer = setTimeout(scan, SCAN_DEBOUNCE);
      }
    }, 1000);
  }

  // ==========================================================
  // Settings (synced via chrome.storage)
  // ==========================================================
  function loadSettings() {
    chrome.storage?.sync?.get?.(['intuneLensSettings'], (r) => {
      if (r?.intuneLensSettings) Object.assign(settings, r.intuneLensSettings);
    });
    chrome.storage?.onChanged?.addListener?.((ch) => {
      if (ch.intuneLensSettings) {
        Object.assign(settings, ch.intuneLensSettings.newValue);
        if (!settings.enabled) hideImmediate();
      }
    });
  }

  // ==========================================================
  // Bootstrap
  // ==========================================================
  function init() {
    if (!location.hostname.includes('intune.microsoft.com')) {
      console.log('[Intune Lens] Not on Intune portal — inactive.');
      return;
    }
    log('🚀 Initializing on', location.href);
    log('document.readyState =', document.readyState);
    loadSettings();
    setupTokenBridge();
    ensureContainer();
    setupObserver();

    // Check token status after a short delay
    setTimeout(() => {
      chrome.runtime.sendMessage({ type: 'getStatus' }, (r) => {
        if (r?.hasToken) {
          log('✅ Token available, cache size:', r.cacheSize);
        } else {
          warn('⚠ No token yet — navigate around Intune to trigger Graph calls');
        }
      });
    }, 3000);

    // Initial scan + FAB after page settles
    setTimeout(() => { log('Running initial scan…'); scan(); updateFab(); }, 2000);

    log('✅ Ready — hover over Intune object links to see details.');
    log('💡 Filter console by "[IL]" to see only Intune Lens messages.');
  }

  document.readyState === 'loading'
    ? document.addEventListener('DOMContentLoaded', init)
    : init();
})();
