# 🔍 Intune Lens

A browser extension (Edge / Chrome) that adds **rich hover cards** to the Microsoft Intune portal. Hover over any device, user, app, or policy link and instantly see detailed information — powered by Microsoft Graph.

## ✨ Features

| Object | Card shows |
|--------|-----------|
| 💻 **Device** | Name, OS & version, model, serial, compliance state, encryption, primary user, last check-in, enrollment date, storage |
| 👤 **User** | Display name, UPN, job title, department, office, phone, account status, managed devices list with compliance dots |
| 📱 **App** | Name, publisher, description, created / modified dates |
| 📋 **Policy** | Name, version, description, created / modified dates |

### Additional highlights

- **Zero infrastructure** — no backend server; the extension reuses the token from your active Intune session
- **Automatic token capture** — intercepts the bearer token that the Intune SPA already sends to `graph.microsoft.com`
- **5-minute response cache** — fast repeat hovers without hammering Graph
- **Dark mode** — follows system preference
- **Configurable** — toggle card types on/off, adjust hover delay

## 🏗️ Architecture

```
┌── intune.microsoft.com ──────────────────────────────┐
│                                                       │
│  inject.js (page context)                             │
│     ↓  intercepts fetch/XHR → captures Bearer token   │
│     ↓  window.postMessage                             │
│                                                       │
│  content.js (content script)                          │
│     • MutationObserver → detects device/user/… links  │
│     • mouseenter → shows hover card                   │
│     • relays token to service worker                  │
│     ↕  chrome.runtime.sendMessage                     │
│                                                       │
│  background.js (service worker)                       │
│     • stores token in chrome.storage.session          │
│     • makes Graph API calls                           │
│     • caches responses (Map, 5 min TTL)               │
│                                                       │
│  popup.html/js/css (extension popup)                  │
│     • toggle cards on/off                             │
│     • adjust hover delay                              │
│     • clear cache, view token status                  │
└───────────────────────────────────────────────────────┘
```

## 📦 Installation (sideload)

### Edge

1. Open `edge://extensions/`
2. Enable **Developer mode** (toggle in bottom-left)
3. Click **Load unpacked**
4. Select the `Intune.Lens` folder
5. Navigate to [intune.microsoft.com](https://intune.microsoft.com) and start hovering!

### Chrome

1. Open `chrome://extensions/`
2. Enable **Developer mode** (toggle in top-right)
3. Click **Load unpacked**
4. Select the `Intune.Lens` folder

## 🔑 How token capture works

The Intune portal is a React SPA that calls `graph.microsoft.com` using a bearer token obtained via MSAL. Intune Lens captures this token in two complementary ways:

1. **`inject.js`** — injected into the page context, monkey-patches `fetch()` and `XMLHttpRequest.setRequestHeader()` to intercept the `Authorization` header on Graph calls, then relays the token via `window.postMessage`.
2. **`background.js`** — uses `chrome.webRequest.onSendHeaders` as a fallback to read the header directly from outbound requests.

The captured token is stored in `chrome.storage.session` (volatile — cleared on browser restart). **No tokens are ever sent to external servers.**

## ⚙️ Configuration

Click the extension icon in the toolbar to open settings:

| Setting | Default | Description |
|---------|---------|-------------|
| Extension Enabled | ✓ | Master on/off |
| Device cards | ✓ | Show cards on device links |
| User cards | ✓ | Show cards on user links |
| App cards | ✓ | Show cards on app links |
| Policy cards | ✓ | Show cards on policy links |
| Hover delay | 400 ms | Time before the card appears |

## 🔒 Permissions explained

| Permission | Why |
|-----------|-----|
| `storage` | Save settings and session token |
| `webRequest` | Observe Graph API requests to capture the auth token |
| `host_permissions: intune.microsoft.com` | Inject content scripts into the Intune portal |
| `host_permissions: graph.microsoft.com` | Make Graph API calls from the service worker |

## 🛠️ Development

```bash
# No build step required — pure vanilla JS/CSS
# Edit files, then reload the extension in edge://extensions/
```

### File structure

```
Intune.Lens/
├── manifest.json    # MV3 extension manifest
├── background.js    # Service worker (Graph API + cache + token)
├── inject.js        # Page-context token interceptor
├── content.js       # DOM detection + hover cards
├── content.css      # Card styles (light + dark)
├── popup.html       # Settings UI
├── popup.js         # Settings logic
├── popup.css        # Settings styles
└── README.md        # This file
```

### Adding icons

Place PNG icons in the project root and add to `manifest.json`:

```json
"icons": {
  "16": "icon16.png",
  "48": "icon48.png",
  "128": "icon128.png"
}
```

## 📋 Roadmap

- [ ] Group membership card for devices
- [ ] Configuration profile assignment status
- [ ] App installation status per device
- [ ] Keyboard shortcut to pin a card
- [ ] Export card data to clipboard
- [ ] Custom icons

## 📄 License

MIT
