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

// ============================================================
// 🥚 Easter Eggs
// ============================================================

// 1. Konami Code: ↑↑↓↓←→←→BA
const KONAMI = ['ArrowUp','ArrowUp','ArrowDown','ArrowDown','ArrowLeft','ArrowRight','ArrowLeft','ArrowRight','b','a'];
let konamiIdx = 0;
document.addEventListener('keydown', (e) => {
  const key = e.key.length === 1 ? e.key.toLowerCase() : e.key;
  console.log(`🎮 Key: "${key}" | expecting: "${KONAMI[konamiIdx]}" | progress: ${konamiIdx}/${KONAMI.length}`);
  if (key === KONAMI[konamiIdx]) {
    konamiIdx++;
    if (konamiIdx === KONAMI.length) {
      console.log('🎮 KONAMI CODE ACTIVATED! 🎉');
      konamiIdx = 0;
      showClippy();
    }
  } else {
    konamiIdx = 0;
  }
});

function showClippy() {
  const existing = document.getElementById('il-egg');
  if (existing) { existing.remove(); return; }

  const phrases = [
    "It looks like you're trying to manage devices.\nWould you like help enrolling 10,000 of them?",
    "Pro tip: the 'Sync' button doesn't work faster\nif you click it 47 times. I counted.",
    "Your BitLocker key rotated successfully!\n...just kidding. Check the audit logs 😏",
    "Congratulations! You've found me.\nI've been hiding here since Windows XP.",
    "Fun fact: the Intune portal loads faster\nif you believe hard enough ✨",
    "Your Autopilot profile failed to apply.\nThe device identifies as a toaster 🍞",
    "That device hasn't checked in for 90 days.\nIt's on a spiritual journey 🧘",
    "You assigned 12 policies to one device.\nThat device now needs therapy.",
    "Every time you create a dynamic group,\nan Azure AD engineer gets a headache 🤕",
    "Your SCEP certificate deployment failed.\nNobody understands SCEP. Nobody.",
    "'Pending' is Intune's way of saying\n'I'll get to it when I get to it' 🐌",
  ];

  const el = document.createElement('div');
  el.id = 'il-egg';
  el.innerHTML = `
    <div class="egg-clippy">📎</div>
    <div class="egg-bubble">
      <div class="egg-text">${phrases[Math.floor(Math.random() * phrases.length)]}</div>
      <div class="egg-dismiss">click to dismiss</div>
    </div>
  `;
  el.addEventListener('click', () => el.remove());
  document.body.appendChild(el);
}

// 2. Triple-click on version footer → matrix rain
let footerClicks = 0;
let footerTimer = null;
document.querySelector('.p-footer')?.addEventListener('click', () => {
  footerClicks++;
  clearTimeout(footerTimer);
  footerTimer = setTimeout(() => { footerClicks = 0; }, 500);
  if (footerClicks >= 3) {
    footerClicks = 0;
    showMatrix();
  }
});

// 3. Click 5x on 🔍 logo → Clippy
let logoClicks = 0;
let logoTimer = null;
document.getElementById('p-logo')?.addEventListener('click', () => {
  logoClicks++;
  clearTimeout(logoTimer);
  logoTimer = setTimeout(() => { logoClicks = 0; }, 1500);
  if (logoClicks >= 5) {
    logoClicks = 0;
    showClippy();
  }
});

function showMatrix() {
  const existing = document.getElementById('il-matrix');
  if (existing) { existing.remove(); return; }

  const canvas = document.createElement('canvas');
  canvas.id = 'il-matrix';
  canvas.style.cssText = 'position:fixed;inset:0;z-index:99999;pointer-events:none;opacity:0.85;';
  document.body.appendChild(canvas);
  const ctx = canvas.getContext('2d');
  canvas.width = 300; canvas.height = 400;

  const chars = 'INTUNE01COMPLIANTSYNCDEPLOY🔒📱💻📋✓✗'.split('');
  const cols = Math.floor(canvas.width / 14);
  const drops = Array(cols).fill(1);

  const interval = setInterval(() => {
    ctx.fillStyle = 'rgba(0,0,0,0.05)';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = '#0078d4';
    ctx.font = '13px monospace';
    for (let i = 0; i < cols; i++) {
      ctx.fillText(chars[Math.floor(Math.random() * chars.length)], i * 14, drops[i] * 14);
      if (drops[i] * 14 > canvas.height && Math.random() > 0.975) drops[i] = 0;
      drops[i]++;
    }
  }, 50);

  setTimeout(() => { clearInterval(interval); canvas.remove(); }, 5000);
}
