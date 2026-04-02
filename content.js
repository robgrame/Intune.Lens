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
  // Token bridge — inject page-level script
  // ==========================================================
  function setupTokenBridge() {
    const s = document.createElement('script');
    s.src = chrome.runtime.getURL('inject.js');
    s.onload = () => s.remove();
    (document.head || document.documentElement).appendChild(s);

    window.addEventListener('message', (e) => {
      if (e.origin !== location.origin) return;
      if (e.data?.type === '__INTUNE_LENS_TOKEN__') {
        chrome.runtime.sendMessage({ type: 'setToken', token: e.data.token });
      }
    });
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
  function detect(el) {
    const href = el.getAttribute('href') || '';
    for (const p of ROUTE_PATTERNS) {
      const m = href.match(p.re);
      if (m) return { type: p.type, id: m[1] };
    }
    return null;
  }

  function scan() {
    if (!settings.enabled) return;
    const links = document.querySelectorAll(`a[href*="#"]:not([${PROCESSED}])`);
    for (const a of links) {
      const obj = detect(a);
      if (!obj) continue;
      const key = `show${obj.type[0].toUpperCase()}${obj.type.slice(1)}Cards`;
      if (settings[key] === false) continue;

      a.setAttribute(PROCESSED, obj.type);
      a.setAttribute(ID_ATTR, obj.id);
      a.classList.add('il-link');
      a.addEventListener('mouseenter', onEnter);
      a.addEventListener('mouseleave', onLeave);
    }
  }

  function onEnter(e) {
    const el = e.currentTarget;
    clearTimeout(hoverTimer);
    clearTimeout(hideTimer);
    hoverTimer = setTimeout(
      () => showCard(el, el.getAttribute(PROCESSED), el.getAttribute(ID_ATTR)),
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
      hideImmediate();
      clearTimeout(scanTimer);
      scanTimer = setTimeout(scan, SCAN_DEBOUNCE);
    });
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
    if (!location.hostname.includes('intune.microsoft.com')) return;
    console.log('[Intune Lens] Initializing…');
    loadSettings();
    setupTokenBridge();
    ensureContainer();
    setupObserver();
    setTimeout(scan, 2000);
    console.log('[Intune Lens] Ready — hover over Intune objects to see details.');
  }

  document.readyState === 'loading'
    ? document.addEventListener('DOMContentLoaded', init)
    : init();
})();
