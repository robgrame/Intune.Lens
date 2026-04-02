# 🔍 Intune Lens

> Rich hover cards for the Microsoft Intune admin center — see device details, compliance, primary user, and more without clicking.

<p align="center">
  <img src="https://img.shields.io/badge/manifest-v3-blue" alt="Manifest V3" />
  <img src="https://img.shields.io/badge/edge-%E2%9C%93-green" alt="Edge" />
  <img src="https://img.shields.io/badge/chrome-%E2%9C%93-green" alt="Chrome" />
  <img src="https://img.shields.io/badge/license-MIT-lightgrey" alt="MIT" />
  <img src="https://img.shields.io/badge/build-no%20build%20step-brightgreen" alt="No build step" />
</p>

---

## The Problem

Managing devices in Intune means **endless clicking** — open a device, wait for the blade to load, check compliance, go back, repeat. There's no quick way to glance at device health, primary user, or last check-in without leaving the list view.

## The Solution

**Intune Lens** adds Adaptive-Card-style hover cards directly into the Intune portal. Hover over a device name in any list view and instantly see:

| Object | Card shows |
|--------|-----------|
| 💻 **Device** | Name, OS & version, model, serial number, compliance state, encryption status, management agent, primary user, last check-in, enrollment date, storage usage |
| 👤 **User** | Display name, UPN, job title, department, office, phone, account status, list of managed devices with compliance indicators |
| 📱 **App** | Name, publisher, description, created / modified dates |
| 📋 **Policy** | Name, version, description, created / modified dates |

### Highlights

- **Zero infrastructure** — no backend server, no Azure Functions, no App Registration. The extension reuses the token from your active Intune session.
- **Works inside Intune's iframes** — the portal renders content inside sandboxed React blade iframes; Intune Lens injects into those frames automatically.
- **5-minute response cache** — fast repeat hovers without extra Graph calls.
- **Dark mode** — follows your system preference.
- **Configurable** — toggle card types on/off, adjust hover delay via the popup.

---

## 📦 Installation

No store listing yet — sideload the extension in developer mode.

### Microsoft Edge (recommended)

1. Clone this repo or download the ZIP
2. Open **`edge://extensions/`**
3. Enable **Developer mode** (bottom-left toggle)
4. Click **Load unpacked** → select the `Intune.Lens` folder
5. Navigate to [intune.microsoft.com](https://intune.microsoft.com) and start hovering!

### Google Chrome

1. Open **`chrome://extensions/`**
2. Enable **Developer mode** (top-right toggle)
3. Click **Load unpacked** → select the `Intune.Lens` folder

> **Note:** After updating the extension files, click the 🔄 button on the extension card **and** refresh the Intune tab (F5).

---

## 🏗️ Architecture

The Intune admin center is an Azure-portal-based SPA that renders blade content inside **cross-origin sandboxed iframes** (`sandbox-*.reactblade.portal.azure.net`) and makes Graph API calls from **Web Workers**. This makes traditional DOM scraping and fetch interception ineffective. Intune Lens works around these constraints:

```
┌─ intune.microsoft.com (main frame) ──────────────────┐
│                                                       │
│  inject.js (MAIN world, document_start)               │
│     • monkey-patches fetch/XHR for token capture      │
│     • intercepts Response.prototype.json for data     │
│                                                       │
│  content.js (ISOLATED world)                          │
│     • URL-based list-page detection                   │
│     • FAB (floating action button) on detail pages    │
│     • relays captured tokens to service worker        │
│                                                       │
├─ sandbox-*.reactblade.portal.azure.net (blade iframe) │
│                                                       │
│  content.js (ISOLATED world, all_frames)              │
│     • fetches device/user list via Graph API           │
│     • scans all leaf DOM elements for name matching   │
│     • renders hover cards on mouseenter               │
│     • MutationObserver for virtual-scroll re-renders  │
│                                                       │
├─ background.js (service worker) ─────────────────────┤
│  • webRequest.onSendHeaders → captures Bearer token   │
│  • chrome.storage.session → volatile token storage    │
│  • Graph API proxy with 5-min TTL cache (Map)         │
│  • message hub for all content script instances       │
└───────────────────────────────────────────────────────┘
```

### Token capture

The extension captures the Bearer token that Intune already uses for its own Graph API calls — **no App Registration or admin consent required**.

| Method | Where | How |
|--------|-------|-----|
| `webRequest.onSendHeaders` | Service worker | Reads `Authorization` header from Graph requests (primary) |
| `inject.js` fetch wrapper | Main frame (MAIN world) | Monkey-patches `window.fetch` to relay tokens via `postMessage` (fallback) |

The token is stored in `chrome.storage.session` (volatile — cleared on browser restart). **No tokens are ever sent to external servers.**

### Data flow for hover cards

1. Content script detects a list page (e.g. `~/allDevices`) via URL hash pattern
2. Sends `graphQuery` message to service worker → `GET /deviceManagement/managedDevices?$top=50`
3. Service worker uses captured token, calls Graph, caches response
4. Content script builds a `name → { id, type }` lookup table
5. `scanGridCells()` iterates all leaf DOM elements, matches `textContent` against known names
6. Matched elements get hover handlers → `mouseenter` triggers card fetch & render
7. Cards use cached data when available; otherwise make a per-object Graph call

---

## ⚙️ Configuration

Click the Intune Lens icon in the browser toolbar:

| Setting | Default | Description |
|---------|---------|-------------|
| Extension Enabled | ✓ | Master on/off toggle |
| Device cards | ✓ | Show hover cards on device names |
| User cards | ✓ | Show hover cards on user names |
| App cards | ✓ | Show hover cards on app names |
| Policy cards | ✓ | Show hover cards on policy names |
| Hover delay | 400 ms | Delay before the card appears (100–1500 ms) |

---

## 🔒 Permissions

| Permission | Why |
|-----------|-----|
| `storage` | Persist user settings; store session token |
| `webRequest` | Observe Graph API requests to capture the auth token |
| `intune.microsoft.com` | Inject content scripts + token interceptor |
| `graph.microsoft.com` | Make Graph API calls from the service worker |
| `hosting.portal.azure.net` | Inject content scripts into Intune blade iframes |
| `*.reactblade.portal.azure.net` | Inject content scripts into sandbox blade iframes |

---

## 🛠️ Development

```bash
# No build step — pure vanilla JS/CSS
git clone https://github.com/robgrame/Intune.Lens.git
# Load the folder as an unpacked extension
# Edit → reload extension → refresh Intune tab
```

### Project structure

```
Intune.Lens/
├── manifest.json     # MV3 extension manifest
├── background.js     # Service worker — Graph API proxy, token, cache
├── inject.js         # Page-context token + data interceptor (MAIN world)
├── content.js        # DOM detection + hover cards (main frame + blade iframes)
├── content.css       # Fluent-UI-inspired card styles (light + dark)
├── popup.html        # Settings popup
├── popup.js          # Settings logic
├── popup.css         # Popup styles
├── icon{16,48,128}.png
└── README.md
```

### Debugging

Open DevTools (F12) on the Intune tab and filter the console by **`[IL]`**:

```
[IL] 🚀 Intune Lens v1.8.0 — Main frame on intune.microsoft.com/...
[IL] 🚀 Intune Lens v1.8.0 — Blade iframe on sandbox-2.reactblade.portal.azure.net/...
[IL] 📋 MATCH → device list. Calling Graph…
[IL] 📋 Got 12 device(s) from Graph ✓
[IL] 📇 Lookup table: 14 entries. Scanning grid…
[IL] GridScan: matched 74 cells ✓
[IL] Hover → device a1b2c3d4-...
```

---

## 📋 Roadmap

- [ ] Group membership list in device cards
- [ ] Configuration profile assignment status
- [ ] App installation status per device
- [ ] Keyboard shortcut to pin/dismiss cards
- [ ] Export card data to clipboard (JSON / Markdown)
- [ ] Support for compliance policy detail pages
- [ ] Chrome Web Store / Edge Add-ons listing

---

## 🤝 Contributing

Contributions are welcome! Please open an issue first to discuss what you'd like to change.

## 📄 License

[MIT](LICENSE)
