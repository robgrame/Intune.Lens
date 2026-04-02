// ============================================================
// Intune Lens — Popup (Settings)
// ============================================================
const KEYS = ['enabled', 'showDeviceCards', 'showUserCards', 'showAppCards', 'showPolicyCards'];
const DEFAULTS = {
  enabled: true,
  showDeviceCards: true,
  showUserCards: true,
  showAppCards: true,
  showPolicyCards: true,
  hoverDelay: 400
};

// --- Load saved settings ---
chrome.storage.sync.get(['intuneLensSettings'], (r) => {
  const s = { ...DEFAULTS, ...(r.intuneLensSettings || {}) };
  KEYS.forEach(k => { document.getElementById(k).checked = s[k]; });
  const slider = document.getElementById('hoverDelay');
  slider.value = s.hoverDelay;
  document.getElementById('hoverDelayVal').textContent = `${s.hoverDelay} ms`;
});

// --- Save on change ---
function save() {
  const s = {};
  KEYS.forEach(k => { s[k] = document.getElementById(k).checked; });
  s.hoverDelay = parseInt(document.getElementById('hoverDelay').value, 10);
  chrome.storage.sync.set({ intuneLensSettings: s });
}

KEYS.forEach(k => document.getElementById(k).addEventListener('change', save));

const slider = document.getElementById('hoverDelay');
slider.addEventListener('input', () => {
  document.getElementById('hoverDelayVal').textContent = `${slider.value} ms`;
});
slider.addEventListener('change', save);

// --- Clear cache ---
document.getElementById('clearCache').addEventListener('click', () => {
  chrome.runtime.sendMessage({ type: 'clearCache' }, () => {
    const btn = document.getElementById('clearCache');
    btn.textContent = '✓ Cleared!';
    setTimeout(() => { btn.textContent = '🗑️ Clear Cache'; }, 1500);
  });
});

// --- Status ---
chrome.runtime.sendMessage({ type: 'getStatus' }, (r) => {
  const el = document.getElementById('status');
  if (!r) { el.textContent = '⚠ No connection'; el.className = 'p-status bad'; return; }
  if (r.hasToken) {
    el.textContent = `✓ Token OK · ${r.cacheSize} cached`;
    el.className = 'p-status ok';
  } else {
    el.textContent = '⚠ No token — browse Intune first';
    el.className = 'p-status bad';
  }
});
