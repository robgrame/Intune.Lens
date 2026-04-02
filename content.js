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

  // Object data captured from Graph responses (inject.js → here)
  const objectCache = new Map();  // id → full object
  const nameToObj   = new Map();  // lowercase name/upn → { id, type }

  // ==========================================================
  // Bridge: tokens + intercepted Graph data from inject.js
  // ==========================================================
  function setupBridge() {
    window.addEventListener('message', (e) => {
      if (e.data?.type === '__INTUNE_LENS_TOKEN__') {
        log('Token captured from page ✓');
        chrome.runtime.sendMessage({ type: 'setToken', token: e.data.token });
      }
      if (e.data?.type === '__INTUNE_LENS_DATA__') {
        const objs = e.data.objects || [];
        log(`📦 Received ${objs.length} objects from page intercept`);
        for (const obj of objs) {
          objectCache.set(obj.id, obj);
          const type = obj._t;
          if (obj.deviceName)         nameToObj.set(obj.deviceName.toLowerCase().trim(), { id: obj.id, type });
          if (obj.displayName)        nameToObj.set(obj.displayName.toLowerCase().trim(), { id: obj.id, type });
          if (obj.userPrincipalName)  nameToObj.set(obj.userPrincipalName.toLowerCase().trim(), { id: obj.id, type });
        }
        log(`📇 Lookup table now has ${nameToObj.size} entries`);
        // Trigger a grid scan now that we have data to match
        scanGridCells();
      }
    });
    log('Bridge listener ready (tokens + data)');
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
    'id','deviceName','operatingSystem','osVersion','model','manufacturer',
    'serialNumber','lastSyncDateTime','complianceState','enrolledDateTime',
    'userPrincipalName','managementAgent','deviceEnrollmentType',
    'isEncrypted','isSupervised','totalStorageSpaceInBytes','freeStorageSpaceInBytes',
    'azureADDeviceId','deviceCategoryDisplayName','jailBroken',
    'wiFiMacAddress','ethernetMacAddress','phoneNumber'
  ].join(',');

  async function fetchDevice(id) {
    let dev = objectCache.get(id);
    if (!dev?.deviceName) {
      dev = await graphQuery(`/deviceManagement/managedDevices/${id}?$select=${DEVICE_SELECT}`, `dev:${id}`);
    }

    // Compliance policy states
    try {
      const cp = await graphQuery(
        `/deviceManagement/managedDevices/${id}/deviceCompliancePolicyStates`,
        `dev-compliance:${id}`
      );
      dev._compliancePolicies = cp.value || [];
    } catch { dev._compliancePolicies = []; }

    // Configuration profile states
    try {
      const cf = await graphQuery(
        `/deviceManagement/managedDevices/${id}/deviceConfigurationStates`,
        `dev-config:${id}`
      );
      dev._configProfiles = cf.value || [];
    } catch { dev._configProfiles = []; }

    // Group memberships (via Azure AD device)
    if (dev.azureADDeviceId) {
      try {
        const aadDevices = await graphQuery(
          `/devices?$filter=deviceId eq '${dev.azureADDeviceId}'&$select=id`,
          `dev-aad:${dev.azureADDeviceId}`
        );
        if (aadDevices.value?.[0]?.id) {
          const groups = await graphQuery(
            `/devices/${aadDevices.value[0].id}/transitiveMemberOf/microsoft.graph.group?$select=displayName,groupTypes&$top=999`,
            `dev-groups:${dev.azureADDeviceId}`
          );
          dev._groups = (groups.value || []).filter(g => g.displayName);
        }
      } catch { /* groups are optional */ }
    }
    if (!dev._groups) dev._groups = [];

    // Managed app install states — try multiple approaches
    dev._apps = [];
    // Approach 1: beta managedAppDiagnostics (v1.0 detectedApps for installed apps)
    try {
      const detected = await graphQuery(
        `/deviceManagement/managedDevices/${id}/detectedApps?$select=displayName,version,sizeInByte&$top=100`,
        `dev-detected:${id}`
      );
      const apps = (detected.value || []).map(a => ({
        displayName: a.displayName,
        displayVersion: a.version,
        installState: 'installed'
      }));
      // Now get app statuses from beta to find failures
      try {
        const statuses = await graphQuery(
          `/beta/deviceManagement/managedDevices/${id}/managedDeviceAppConfigurationStates`,
          `dev-appstatus:${id}`
        );
        for (const s of (statuses.value || [])) {
          if (s.state === 'error' || s.state === 'failed') {
            apps.push({ displayName: s.displayName, installState: 'failed', displayVersion: s.version || '' });
          }
        }
      } catch { /* optional */ }
      dev._apps = apps;
      log(`📱 Device apps: ${apps.length} (${apps.filter(a => a.installState === 'installed').length} installed, ${apps.filter(a => a.installState === 'failed').length} failed)`);
    } catch (err) {
      log(`📱 Device apps: ${err.message}`);
    }

    return dev;
  }

  async function fetchUser(id) {
    let user = objectCache.get(id);
    if (!user?.userPrincipalName) {
      user = await graphQuery(
        `/users/${id}?$select=displayName,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,mail,accountEnabled,createdDateTime`,
        `usr:${id}`
      );
    }
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
    let app = objectCache.get(id);
    if (!app?.publisher) {
      app = await graphQuery(
        `/deviceAppManagement/mobileApps/${id}?$select=id,displayName,publisher,description,createdDateTime,lastModifiedDateTime`,
        `app:${id}`
      );
    }
    // Fetch install summary (installed/failed/notInstalled counts)
    try {
      const summary = await graphQuery(
        `/deviceAppManagement/mobileApps/${id}/installSummary`,
        `app-summary:${id}`
      );
      app._summary = summary;
    } catch { app._summary = null; }

    // Fetch assignments (which groups it's assigned to)
    try {
      const assignments = await graphQuery(
        `/deviceAppManagement/mobileApps/${id}/assignments`,
        `app-assign:${id}`
      );
      app._assignments = assignments.value || [];
    } catch { app._assignments = []; }

    return app;
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

  const APP_STATE_MAP = {
    installed:       { cls: 'ok',   label: 'Installed' },
    failed:          { cls: 'bad',  label: 'Failed' },
    notInstalled:    { cls: 'unk',  label: 'Not installed' },
    uninstallFailed: { cls: 'bad',  label: 'Uninstall failed' },
    pendingInstall:  { cls: 'warn', label: 'Pending' },
    unknown:         { cls: 'unk',  label: 'Unknown' },
    notApplicable:   { cls: 'unk',  label: 'N/A' },
  };

  function deviceAppsHtml(apps) {
    if (!apps || apps.length === 0) return '';

    const byState = {};
    for (const app of apps) {
      const st = app.installState || 'unknown';
      if (!byState[st]) byState[st] = [];
      byState[st].push(app);
    }

    const failed = byState.failed || [];
    const uninstallFailed = byState.uninstallFailed || [];
    const pending = byState.pendingInstall || [];
    const installed = byState.installed || [];
    const other = apps.length - failed.length - uninstallFailed.length - pending.length - installed.length;

    // Stats row
    const statsHtml = `
      <div class="il-stats">
        <div class="il-stat ok-bg"><span class="il-stat-n">${installed.length}</span><span class="il-stat-l">Installed</span></div>
        <div class="il-stat bad-bg"><span class="il-stat-n">${failed.length + uninstallFailed.length}</span><span class="il-stat-l">Failed</span></div>
        <div class="il-stat warn-bg"><span class="il-stat-n">${pending.length}</span><span class="il-stat-l">Pending</span></div>
        ${other > 0 ? `<div class="il-stat unk-bg"><span class="il-stat-n">${other}</span><span class="il-stat-l">Other</span></div>` : ''}
      </div>`;

    // Failed apps detail list
    const allFailed = [...failed, ...uninstallFailed];
    const failedHtml = allFailed.length > 0 ? allFailed.map(a =>
      `<div class="il-dev-item"><span class="il-dot bad"></span>${esc(a.displayName || a.applicationId)} ${a.displayVersion ? `<span class="il-dev-os">${esc(a.displayVersion)}</span>` : ''}</div>`
    ).join('') : '';

    // Pending apps
    const pendingHtml = pending.length > 0 ? pending.map(a =>
      `<div class="il-dev-item"><span class="il-dot warn"></span>${esc(a.displayName || a.applicationId)} ${a.displayVersion ? `<span class="il-dev-os">${esc(a.displayVersion)}</span>` : ''}</div>`
    ).join('') : '';

    return `
        <hr class="il-div">
        <div class="il-sec">
          <div class="il-sec-ttl">Managed Apps (${apps.length})</div>
          ${statsHtml}
          ${failedHtml ? '<div class="il-sec-sub">Failed</div>' + failedHtml : ''}
          ${pendingHtml ? '<div class="il-sec-sub">Pending</div>' + pendingHtml : ''}
        </div>`;
  }

  // ==========================================================
  // Card templates
  // ==========================================================
  function deviceCard(d) {
    const c = badge(d.complianceState);
    const storage = d.totalStorageSpaceInBytes
      ? `${bytes(d.totalStorageSpaceInBytes - (d.freeStorageSpaceInBytes || 0))} / ${bytes(d.totalStorageSpaceInBytes)}`
      : null;

    // Compliance policy summary
    const cpols = d._compliancePolicies || [];
    const cpOk = cpols.filter(p => p.state === 'compliant').length;
    const cpBad = cpols.filter(p => p.state === 'nonCompliant').length;
    const cpUnk = cpols.length - cpOk - cpBad;
    const cpHtml = cpols.length > 0 ? `
        <hr class="il-div">
        <div class="il-sec">
          <div class="il-sec-ttl">Compliance Policies (${cpols.length})</div>
          <div class="il-stats">
            <div class="il-stat ok-bg"><span class="il-stat-n">${cpOk}</span><span class="il-stat-l">Compliant</span></div>
            <div class="il-stat bad-bg"><span class="il-stat-n">${cpBad}</span><span class="il-stat-l">Non-compl.</span></div>
            <div class="il-stat unk-bg"><span class="il-stat-n">${cpUnk}</span><span class="il-stat-l">Other</span></div>
          </div>
          ${cpBad > 0 ? cpols.filter(p => p.state === 'nonCompliant').map(p =>
            `<div class="il-dev-item"><span class="il-dot bad"></span>${esc(p.displayName)}</div>`
          ).join('') : ''}
        </div>` : '';

    // Config profiles summary
    const cfgs = d._configProfiles || [];
    const cfOk = cfgs.filter(p => p.state === 'compliant' || p.state === 'notApplicable').length;
    const cfBad = cfgs.filter(p => p.state === 'error' || p.state === 'conflict' || p.state === 'nonCompliant').length;
    const cfHtml = cfgs.length > 0 ? `
        <hr class="il-div">
        <div class="il-sec">
          <div class="il-sec-ttl">Config Profiles (${cfgs.length})</div>
          <div class="il-stats">
            <div class="il-stat ok-bg"><span class="il-stat-n">${cfOk}</span><span class="il-stat-l">OK</span></div>
            <div class="il-stat bad-bg"><span class="il-stat-n">${cfBad}</span><span class="il-stat-l">Error</span></div>
            <div class="il-stat unk-bg"><span class="il-stat-n">${cfgs.length - cfOk - cfBad}</span><span class="il-stat-l">Other</span></div>
          </div>
          ${cfBad > 0 ? cfgs.filter(p => ['error','conflict','nonCompliant'].includes(p.state)).map(p =>
            `<div class="il-dev-item"><span class="il-dot bad"></span>${esc(p.displayName)}</div>`
          ).join('') : ''}
        </div>` : '';

    // Groups
    const groups = d._groups || [];
    const grpHtml = groups.length > 0 ? `
        <hr class="il-div">
        <div class="il-sec">
          <div class="il-sec-ttl">Groups (${groups.length})</div>
          ${groups.map(g => {
            const isDynamic = g.groupTypes?.includes('DynamicMembership');
            return `<div class="il-dev-item"><span class="il-dot unk"></span>${esc(g.displayName)} ${isDynamic ? '<span class="il-dev-os">dynamic</span>' : ''}</div>`;
          }).join('')}
        </div>` : '';

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
          ${row('Category', d.deviceCategoryDisplayName)}
          ${row('Encryption', d.isEncrypted ? '🔒 Encrypted' : '🔓 Not encrypted')}
          ${row('Management', d.managementAgent)}
          ${storage ? row('Storage', storage) : ''}
          ${d.wiFiMacAddress ? row('Wi-Fi MAC', d.wiFiMacAddress) : ''}
          ${d.ethernetMacAddress ? row('Ethernet MAC', d.ethernetMacAddress) : ''}
        </div>
        <hr class="il-div">
        <div class="il-sec">
          <div class="il-sec-ttl">User & Activity</div>
          ${row('Primary User', d.userPrincipalName)}
          ${row('Last Check-in', ago(d.lastSyncDateTime))}
          ${row('Enrolled', ago(d.enrolledDateTime))}
          ${row('Enrollment Type', d.deviceEnrollmentType)}
        </div>
        ${cpHtml}
        ${cfHtml}
        ${grpHtml}
        ${deviceAppsHtml(d._apps)}
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
      ? `<div class="il-desc">${esc(a.description.substring(0, 150))}${a.description.length > 150 ? '…' : ''}</div>` : '';

    // Install summary stats
    const s = a._summary;
    const statsHtml = s ? `
        <div class="il-sec">
          <div class="il-sec-ttl">Install Status</div>
          <div class="il-stats">
            <div class="il-stat ok-bg"><span class="il-stat-n">${s.installedDeviceCount ?? 0}</span><span class="il-stat-l">Installed</span></div>
            <div class="il-stat bad-bg"><span class="il-stat-n">${s.failedDeviceCount ?? 0}</span><span class="il-stat-l">Failed</span></div>
            <div class="il-stat unk-bg"><span class="il-stat-n">${s.notInstalledDeviceCount ?? 0}</span><span class="il-stat-l">Not installed</span></div>
            <div class="il-stat warn-bg"><span class="il-stat-n">${s.pendingInstallDeviceCount ?? 0}</span><span class="il-stat-l">Pending</span></div>
          </div>
        </div>
        <hr class="il-div">` : '';

    // Assignments
    const assigns = (a._assignments || []);
    const assignHtml = assigns.length > 0 ? `
        <div class="il-sec">
          <div class="il-sec-ttl">Assignments (${a._assignments.length})</div>
          <div class="il-assign-list">${assigns.map(asg => {
            const intent = asg.intent || 'unknown';
            const intentMap = {
              required: { cls: 'bad', icon: '🔴' },
              available: { cls: 'ok', icon: '🟢' },
              uninstall: { cls: 'unk', icon: '⚪' },
              availableWithoutEnrollment: { cls: 'ok', icon: '🟡' }
            };
            const ic = intentMap[intent] || { cls: 'unk', icon: '⚫' };
            const target = asg.target;
            let groupName = 'All devices';
            if (target?.groupId) groupName = target.groupId.substring(0, 8) + '…';
            if (target?.['@odata.type']?.includes('allDevices')) groupName = 'All devices';
            if (target?.['@odata.type']?.includes('allUsers')) groupName = 'All users';
            if (target?.['@odata.type']?.includes('allLicensedUsers')) groupName = 'All licensed users';
            return `<div class="il-assign-item"><span class="il-dot ${ic.cls}"></span><span>${esc(intent)}</span><span class="il-assign-target">${esc(groupName)}</span></div>`;
          }).join('')}</div>
        </div>` : '';

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
        <hr class="il-div">
        ${statsHtml}
        ${assignHtml}
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

    // Position fixed in viewport (won't move on page scroll)
    const rect = anchor.getBoundingClientRect();
    const W = 400, MARGIN = 8;
    let left = rect.left;
    let top  = rect.bottom + MARGIN;

    if (left + W > window.innerWidth)  left = window.innerWidth - W - MARGIN;
    if (left < MARGIN)                 left = MARGIN;
    if (rect.bottom + 320 > window.innerHeight) {
      top = rect.top - 320 - MARGIN;
      if (top < 0) top = rect.bottom + MARGIN;
    }

    card.style.position = 'fixed';
    card.style.left = `${left}px`;
    card.style.top  = `${top}px`;

    hideImmediate();
    container.appendChild(card);
    currentCard = card;
    card._pinned = false;

    card.addEventListener('mouseenter', () => {
      clearTimeout(hideTimer);
      // Auto-pin when mouse enters card
      if (!card._pinned) {
        card._pinned = true;
        card.classList.add('il-pinned');
        const pin = card.querySelector('.il-pin');
        if (pin) { pin.textContent = '📍'; pin.title = 'Unpin card'; }
      }
    });
    card.addEventListener('mouseleave', () => { if (!card._pinned) scheduleHide(); });
    // Auto-pin on scroll (user is reading content)
    card.addEventListener('scroll', () => {
      if (!card._pinned) {
        card._pinned = true;
        card.classList.add('il-pinned');
        const pin = card.querySelector('.il-pin');
        if (pin) { pin.textContent = '📍'; pin.title = 'Unpin card'; }
      }
    }, { passive: true });

    requestAnimationFrame(() => card.classList.remove('il-enter'));

    // Fetch & render, then add controls
    (async () => {
      try {
        const fetchers = { device: fetchDevice, user: fetchUser, app: fetchApp, policy: fetchPolicy };
        const renderers = { device: deviceCard, user: userCard, app: appCard, policy: policyCard };
        const data = await fetchers[type](id);
        if (card === currentCard) {
          card.innerHTML = renderers[type](data);
          addCardControls(card);
        }
      } catch (err) {
        if (card === currentCard) {
          card.innerHTML = errorCard(err.message);
          addCardControls(card);
        }
      }
    })();
  }

  // Add close button + drag + pin to a rendered card
  function addCardControls(card) {
    const bar = document.createElement('div');
    bar.className = 'il-toolbar';
    bar.innerHTML = `
      <span class="il-pin" title="Pin card (keep open)">📌</span>
      <span class="il-drag-hint">⠿</span>
      <span class="il-close" title="Close card">✕</span>
    `;
    card.prepend(bar);

    // Close
    bar.querySelector('.il-close').addEventListener('click', (e) => {
      e.stopPropagation();
      hideImmediate(true);
    });

    // Pin toggle
    const pinBtn = bar.querySelector('.il-pin');
    pinBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      card._pinned = !card._pinned;
      card.classList.toggle('il-pinned', card._pinned);
      pinBtn.textContent = card._pinned ? '📍' : '📌';
      pinBtn.title = card._pinned ? 'Unpin card' : 'Pin card (keep open)';
    });

    // Drag via toolbar
    let dragX = 0, dragY = 0, startLeft = 0, startTop = 0, dragging = false;

    bar.addEventListener('mousedown', (e) => {
      if (e.target.closest('.il-close') || e.target.closest('.il-pin')) return;
      e.preventDefault();
      dragging = true;
      card._pinned = true;
      card.classList.add('il-pinned');
      pinBtn.textContent = '📍';
      dragX = e.clientX;
      dragY = e.clientY;
      startLeft = parseInt(card.style.left) || 0;
      startTop  = parseInt(card.style.top)  || 0;
      card.classList.add('il-dragging');
    });

    document.addEventListener('mousemove', (e) => {
      if (!dragging) return;
      card.style.left = `${startLeft + e.clientX - dragX}px`;
      card.style.top  = `${startTop  + e.clientY - dragY}px`;
    });

    document.addEventListener('mouseup', () => {
      if (dragging) { dragging = false; card.classList.remove('il-dragging'); }
    });
  }

  function scheduleHide() {
    if (currentCard?._pinned) return; // never auto-hide pinned cards
    clearTimeout(hideTimer);
    hideTimer = setTimeout(hideImmediate, 600);
  }

  function hideImmediate(force) {
    if (!currentCard) return;
    if (currentCard._pinned && !force) return; // pinned → only close via ✕
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

    // Also run grid-cell scan (name-based matching)
    scanGridCells();

    // Debug: dump DOM info on first scan
    if (!scan._dbg) {
      scan._dbg = true;
      const allA = document.querySelectorAll('a[href]');
      const sampleHrefs = [...allA].slice(0, 30).map(a => a.getAttribute('href'));
      log(`DOM has ${allA.length} <a> tags. Sample hrefs:`, sampleHrefs);
      log('Hash:', location.hash);
      log('Pathname:', location.pathname);

      const roles = new Set();
      document.querySelectorAll('[role]').forEach(el => roles.add(el.getAttribute('role')));
      log('ARIA roles found on page:', [...roles]);
    }
  }

  // ==========================================================
  // Grid cell scan — match visible text against known objects.
  // Uses exact match first, then partial match for long names
  // (app names are often truncated in the portal).
  // ==========================================================
  function findMatch(text) {
    const lower = text.toLowerCase().trim();
    if (!lower || lower.length < 2) return null;

    // 1. Exact match (fast)
    const exact = nameToObj.get(lower);
    if (exact) return exact;

    // 2. Partial: known name starts with this text, or this text starts with a known name
    if (lower.length >= 4) {
      for (const [name, obj] of nameToObj) {
        if (name.startsWith(lower) || lower.startsWith(name)) return obj;
      }
    }
    return null;
  }

  function scanGridCells() {
    if (nameToObj.size === 0 || !settings.enabled) return;

    const all = document.body.getElementsByTagName('*');
    let matched = 0;

    for (let i = 0; i < all.length; i++) {
      const el = all[i];
      if (el.hasAttribute(PROCESSED)) continue;
      const tag = el.tagName;
      if (tag === 'SCRIPT' || tag === 'STYLE' || tag === 'HEAD' || tag === 'HTML' || tag === 'BODY' || tag === 'NOSCRIPT') continue;

      // Check title attribute on ANY element (works for deep containers too)
      let text = el.getAttribute('title')?.trim();
      let matchObj = text ? findMatch(text) : null;

      // For leaf-ish elements, also check aria-label and textContent
      if (!matchObj) {
        const ariaLabel = el.getAttribute('aria-label')?.trim();
        if (ariaLabel) matchObj = findMatch(ariaLabel);
      }

      if (!matchObj && el.children.length <= 3) {
        text = el.textContent?.trim();
        if (text && text.length >= 2 && text.length < 300) {
          matchObj = findMatch(text);
        }
      }

      if (!matchObj) continue;

      el.setAttribute(PROCESSED, matchObj.type);
      el.setAttribute(ID_ATTR, matchObj.id);
      el.classList.add('il-link');
      el.style.cursor = 'pointer';
      el.addEventListener('mouseenter', onEnter);
      el.addEventListener('mouseleave', onLeave);
      matched++;
    }

    if (matched > 0) {
      log(`GridScan: matched ${matched} cells ✓`);
    } else if (IS_BLADE && !scanGridCells._dumped) {
      scanGridCells._dumped = true;
      const samples = [];
      const titles = [];
      for (let i = 0; i < all.length; i++) {
        const el = all[i];
        if (['SCRIPT','STYLE','HEAD','HTML','BODY','NOSCRIPT','BR','HR','SVG','PATH','LINK','META'].includes(el.tagName)) continue;
        const t = el.getAttribute('title')?.trim();
        if (t && t.length >= 3 && !titles.includes(t)) titles.push(t);
        if (el.children.length > 3) continue;
        const tc = el.textContent?.trim();
        if (tc && tc.length >= 3 && tc.length < 100 && !samples.includes(tc)) samples.push(tc);
        if (samples.length >= 20 && titles.length >= 10) break;
      }
      const names = [...nameToObj.keys()].slice(0, 5);
      log('GridScan: 0 matches. Looking for:', names);
      log('🔍 DOM titles found:', titles.slice(0, 15));
      log('🔍 DOM textContent (leaf):', samples.slice(0, 20));
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
    hideImmediate(true);
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
    close.addEventListener('click', (e) => { e.stopPropagation(); hideImmediate(true); });
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
  // Proactive data fetch — when on a list page, we make our
  // own Graph API call to get object data for name matching.
  // This bypasses the Web Worker issue entirely.
  // ==========================================================
  const LIST_PAGES = [
    { re: /Devices(Windows)?Menu.*\/(allDevices|windowsDevices)/i,
      endpoint: `/deviceManagement/managedDevices?$select=${DEVICE_SELECT}&$top=50&$orderby=deviceName`,
      type: 'device', nameField: 'deviceName' },
    { re: /DevicesMenu.*\/overview/i,
      endpoint: `/deviceManagement/managedDevices?$select=${DEVICE_SELECT}&$top=50&$orderby=deviceName`,
      type: 'device', nameField: 'deviceName' },
    { re: /UserManagementMenuBlade/i,
      endpoint: `/users?$select=id,displayName,userPrincipalName,jobTitle,department,accountEnabled&$top=50`,
      type: 'user', nameField: 'displayName' },
    { re: /Apps(Windows)?Menu/i,
      endpoint: `/deviceAppManagement/mobileApps?$select=id,displayName,publisher,description,createdDateTime,lastModifiedDateTime&$top=50`,
      type: 'app', nameField: 'displayName' },
  ];

  let lastListFetchHash = '';

  async function fetchListData(retryCount = 0) {
    try {
      const hash = location.hash || '';
      if (hash === lastListFetchHash && retryCount === 0) return;

      for (const lp of LIST_PAGES) {
        if (!lp.re.test(hash)) continue;
        log(`📋 MATCH → ${lp.type} list. Calling Graph…${retryCount ? ` (retry #${retryCount})` : ''}`);

        try {
          const data = await graphQuery(lp.endpoint, `list:${lp.type}:${hash}`);
          lastListFetchHash = hash;
          const items = data.value || [];
          log(`📋 Got ${items.length} ${lp.type}(s) from Graph ✓`);

          for (const item of items) {
            objectCache.set(item.id, { ...item, _t: lp.type });
            const name = item[lp.nameField] || item.displayName;
            if (name) nameToObj.set(name.toLowerCase().trim(), { id: item.id, type: lp.type });
            if (item.userPrincipalName)
              nameToObj.set(item.userPrincipalName.toLowerCase().trim(), { id: item.id, type: lp.type });
          }

          log(`📇 Lookup table: ${nameToObj.size} entries. Sharing with iframes…`);
          shareDataWithBlades();
          setTimeout(scanGridCells, 300);
          setTimeout(scanGridCells, 1500);
          setTimeout(scanGridCells, 4000);
        } catch (err) {
          if (retryCount < 5 && /token|Token|expired|401/i.test(err.message)) {
            const delay = (retryCount + 1) * 2000;
            log(`⏳ Token not ready — retry #${retryCount + 1} in ${delay / 1000}s`);
            setTimeout(() => fetchListData(retryCount + 1), delay);
          } else {
            warn(`Graph call failed for ${lp.type}: ${err.message}`);
          }
        }
        return;
      }
    } catch (outerErr) {
      warn('fetchListData() crashed:', outerErr.message);
    }
  }

  // ==========================================================
  // Share data: main frame → blade iframes via chrome.storage
  // ==========================================================
  function shareDataWithBlades() {
    const lookup = Object.fromEntries(nameToObj);
    chrome.runtime.sendMessage({ type: 'setLookup', data: lookup }, () => {
      log(`📤 Shared ${nameToObj.size} entries via service worker`);
    });
  }

  // ==========================================================
  // MutationObserver — react to SPA navigation
  // ==========================================================
  let scanTimer = null;

  function setupObserver() {
    new MutationObserver(() => {
      clearTimeout(scanTimer);
      scanTimer = setTimeout(() => { scan(); scanGridCells(); }, SCAN_DEBOUNCE);
    }).observe(document.body, { childList: true, subtree: true });

    window.addEventListener('hashchange', () => {
      log('Hash changed →', location.hash.substring(0, 100));
      hideImmediate();
      updateFab();
      lastListFetchHash = ''; // allow re-fetch on new page
      nameToObj.clear();      // clear stale data
      fetchListData();
      clearTimeout(scanTimer);
      scan._dbg = false;
      scanTimer = setTimeout(scan, SCAN_DEBOUNCE);
    });

    // Also poll for URL changes (some SPAs don't fire hashchange)
    let lastUrl = location.href;
    setInterval(() => {
      if (location.href !== lastUrl) {
        lastUrl = location.href;
        log('URL changed (poll) →', location.href.substring(0, 120));
        updateFab();
        fetchListData();
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
  const IS_MAIN = location.hostname.includes('intune.microsoft.com');
  const IS_BLADE = location.hostname.includes('hosting.portal.azure.net')
                || location.hostname.includes('portal.azure.net')
                || location.hostname.includes('portal.azure.com');

  function init() {
    // Log even on unrecognized hosts for debugging
    if (!IS_MAIN && !IS_BLADE) {
      console.log('%c[IL]', 'color:#0078d4;font-weight:bold',
        `👀 Loaded on unrecognized host: ${location.hostname} — ${location.href.substring(0, 100)}`);
      return;
    }

    const mode = IS_MAIN ? 'Main frame' : 'Blade iframe';
    log(`🚀 Intune Lens v2.5.0 — ${mode} on`, location.href.substring(0, 100));
    loadSettings();
    ensureContainer();

    if (IS_MAIN) {
      setupBridge();
      setupObserver();
      setTimeout(() => {
        chrome.runtime.sendMessage({ type: 'getStatus' }, (r) => {
          if (r?.hasToken) log('✅ Token available, cache size:', r.cacheSize);
          else warn('⚠ No token yet');
        });
      }, 3000);
      setTimeout(() => { scan(); updateFab(); fetchListData(); }, 2000);

      // Diagnostic: dump all iframe URLs
      setTimeout(() => {
        const iframes = document.querySelectorAll('iframe');
        log(`🖼️ Found ${iframes.length} iframes:`);
        iframes.forEach((f, i) => {
          log(`  iframe[${i}]: src="${f.src?.substring(0, 120) || '(empty)'}" name="${f.name || ''}"`)
        });
      }, 4000);
    }

    if (IS_BLADE) {
      setupBladeObserver();
      // Poll service worker for shared lookup data every 2s
      let pollCount = 0;
      const poll = () => {
        pollCount++;
        chrome.runtime.sendMessage({ type: 'getLookup' }, (r) => {
          if (chrome.runtime.lastError) {
            log('🔲 Blade poll error:', chrome.runtime.lastError.message);
            return;
          }
          if (r?.data && Object.keys(r.data).length > 0) {
            const entries = Object.entries(r.data);
            if (entries.length !== nameToObj.size) {
              nameToObj.clear();
              for (const [name, obj] of entries) nameToObj.set(name, obj);
              log(`🔲 Blade: got ${entries.length} entries from service worker (poll #${pollCount})`);
            }
            scanGridCells();
          } else {
            log(`🔲 Blade: no data yet (poll #${pollCount})`);
          }
        });
      };
      // Poll immediately, then every 2s for up to 30s
      setTimeout(poll, 500);
      const interval = setInterval(() => {
        poll();
        if (pollCount > 15) clearInterval(interval);
      }, 2000);
    }

    log('✅ Ready.');
  }

  function setupBladeObserver() {
    let bladeTimer = null;
    new MutationObserver(() => {
      clearTimeout(bladeTimer);
      bladeTimer = setTimeout(scanGridCells, 600);
    }).observe(document.body, { childList: true, subtree: true });
  }

  document.readyState === 'loading'
    ? document.addEventListener('DOMContentLoaded', init)
    : init();
})();
