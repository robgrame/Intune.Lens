# 🔍 Intune Lens

> Rich hover cards for the Microsoft Intune admin center — see device details, compliance, primary user, apps, policies, and more without clicking.

<p align="center">
  <img src="https://img.shields.io/badge/version-3.2.0-blue" alt="Version 3.2.0" />
  <img src="https://img.shields.io/badge/manifest-v3-blue" alt="Manifest V3" />
  <img src="https://img.shields.io/badge/edge-%E2%9C%93-green" alt="Edge" />
  <img src="https://img.shields.io/badge/chrome-%E2%9C%93-green" alt="Chrome" />
  <img src="https://img.shields.io/badge/license-MIT-lightgrey" alt="MIT" />
  <img src="https://img.shields.io/badge/build-no%20build%20step-brightgreen" alt="No build step" />
</p>

---

## The Problem

Managing devices in Intune means **endless clicking** — open a device, wait for the blade to load, check compliance, go back, repeat. There's no quick way to glance at device health, primary user, app failures, or policy assignments without leaving the list view.

## The Solution

**Intune Lens** adds Adaptive-Card-style hover cards directly into the Intune portal. Hover over any object name in a list view and instantly see a rich summary card.

### 💻 Device Card

| Section | Details |
|---------|---------|
| **Device Info** | Name, OS & version, model, serial, category, encryption, management agent, Wi-Fi/Ethernet MAC |
| **User & Activity** | Primary user, last check-in, enrollment date, enrollment type |
| **Compliance Policies** | Compliant / non-compliant / other counts, non-compliant policy names |
| **Configuration Profiles** | OK / error / other counts, failed profile names |
| **Groups** | All Azure AD group memberships (transitive), with dynamic tag |
| **Managed Apps** | Installed / failed / pending counts, failed & pending app names with versions |
| **Hardware Details** | Battery health, charge cycles, IPv4, subnet, BIOS, TPM, Credential Guard, VBS |
| **Windows Autopatch** | Deployment ring, update status, OS versions, policy, hotpatch enrollment |

### 👤 User Card

| Section | Details |
|---------|---------|
| **Profile** | First/last name, UPN, email, job title, department, office, phone, last sign-in, created date |
| **Entra Roles** | All directory roles (Global Admin, Intune Admin, etc.) |
| **Managed Devices** | All devices with compliance status indicators |
| **Groups** | All group memberships (transitive), with dynamic tag |

### 📱 App Card

| Section | Details |
|---------|---------|
| **App Info** | Name, publisher, description, created/modified dates |
| **Install Status** | Installed / failed / not installed / pending device counts |
| **Assignments** | All assignments with intent (required 🔴 / available 🟢), resolved group names |

### 📋 Policy Card

| Section | Details |
|---------|---------|
| **Policy Info** | Name, type (Configuration Profile / Settings Catalog / Admin Template / Endpoint Security / Compliance), version, description, created/modified dates |
| **Device Status** | Success / failed / error / conflict / N/A counts, last report date |
| **Included Groups** | Assignment groups with filter type |
| **Excluded Groups** | Exclusion groups |

### Card Features

- **📌 Pin** — click the pin button or move mouse to card to keep it open
- **⠿ Drag** — grab the toolbar to reposition the card anywhere
- **✕ Close** — dismiss the card with the close button
- **🖱️ Scroll** — mouse wheel scrolls the card content
- **🌙 Dark mode** — follows system preference

### Highlights

- **Zero infrastructure** — no backend server, no Azure Functions, no App Registration
- **Dual token capture** — intercepts both Graph and Autopatch Bearer tokens from the active session
- **Works inside Intune's iframes** — injects into sandboxed React blade iframes automatically
- **Multi-API support** — Microsoft Graph v1.0, Graph beta, and Windows Autopatch API
- **5-minute response cache** — fast repeat hovers without extra API calls
- **Group name resolution** — assignment groups show display names, not GUIDs
- **Configurable** — toggle card types on/off, adjust hover delay via the popup

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

The Intune admin center is an Azure-portal-based SPA that renders blade content inside **cross-origin sandboxed iframes** (`sandbox-*.reactblade.portal.azure.net`) and makes API calls from **Web Workers**. This makes traditional DOM scraping and fetch interception ineffective. Intune Lens works around these constraints:

```
┌─ intune.microsoft.com (main frame) ──────────────────┐
│                                                       │
│  inject.js (MAIN world, document_start)               │
│     • monkey-patches fetch/XHR for token capture      │
│     • intercepts Response.prototype.json for data     │
│                                                       │
│  content.js (ISOLATED world)                          │
│     • URL-based list-page detection                   │
│     • proactive Graph API fetch (devices/apps/policies│
│     • shares lookup data with blade iframes           │
│     • FAB (floating action button) on detail pages    │
│                                                       │
├─ sandbox-*.reactblade.portal.azure.net (blade iframe) │
│                                                       │
│  content.js (ISOLATED world, all_frames)              │
│     • polls service worker for shared lookup data     │
│     • scans all leaf DOM elements for name matching   │
│     • renders hover cards on mouseenter               │
│     • MutationObserver for virtual-scroll re-renders  │
│                                                       │
├─ background.js (service worker) ─────────────────────┤
│  • webRequest.onSendHeaders → captures Graph +        │
│    Autopatch Bearer tokens separately                 │
│  • chrome.storage.session → volatile token storage    │
│  • Graph API proxy with 5-min TTL cache (Map)         │
│  • Autopatch API proxy with separate token            │
│  • message hub for all content script instances       │
│  • shared lookup data store (main ↔ blade iframes)    │
└───────────────────────────────────────────────────────┘
```

### Token capture

The extension captures Bearer tokens that the Intune portal already uses — **no App Registration or admin consent required**.

| Token | Source | Captured via |
|-------|--------|-------------|
| **Graph** | `graph.microsoft.com` | `webRequest.onSendHeaders` (primary) + `inject.js` fetch wrapper (fallback) |
| **Autopatch** | `services.autopatch.microsoft.com` | `webRequest.onSendHeaders` (requires visiting Autopatch section once) |

Tokens are stored in `chrome.storage.session` (volatile — cleared on browser restart). **No tokens are ever sent to external servers.**

### Supported list pages

| Page | URL pattern | Data source |
|------|------------|-------------|
| All Devices | `~/allDevices`, `~/windowsDevices` | `deviceManagement/managedDevices` |
| Device Overview | `~/overview` | `deviceManagement/managedDevices` |
| Apps | `AppsMenu`, `AppsWindowsMenu` | `deviceAppManagement/mobileApps` |
| Configuration | `~/configuration` | `deviceConfigurations` + `configurationPolicies` (beta) + `groupPolicyConfigurations` (beta) + `intents` (beta) |
| Compliance | `~/compliance` | `deviceCompliancePolicies` |
| Users | `UserManagementMenuBlade` | `/users` |

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
| `storage` | Persist user settings; store session tokens |
| `webRequest` | Observe API requests to capture auth tokens |
| `intune.microsoft.com` | Inject content scripts + token interceptor |
| `graph.microsoft.com` | Make Graph API calls from the service worker |
| `services.autopatch.microsoft.com` | Make Autopatch API calls |
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
├── background.js     # Service worker — Graph + Autopatch API proxy, tokens, cache
├── inject.js         # Page-context token + data interceptor (MAIN world)
├── content.js        # DOM detection + hover cards (main frame + blade iframes)
├── content.css       # Fluent-UI-inspired card styles (light + dark)
├── popup.html        # Settings popup
├── popup.js          # Settings logic + easter eggs 🥚
├── popup.css         # Popup styles
├── icon{16,48,128}.png
├── LICENSE
└── README.md
```

### Debugging

Open DevTools (F12) on the Intune tab and filter the console by **`[IL]`**:

```
[IL] 🚀 Intune Lens v3.2.0 — Main frame on intune.microsoft.com/...
[IL] 🚀 Intune Lens v3.2.0 — Blade iframe on sandbox-2.reactblade.portal.azure.net/...
[IL] 📋 MATCH → device list. Calling Graph…
[IL] 📋 Got 12 from /deviceManagement/managedDevices...
[IL] 📇 Total: 12 device(s), lookup: 14 entries
[IL] 📤 Shared 14 entries via service worker
[IL] 🔲 Blade: got 14 entries from service worker (poll #3)
[IL] GridScan: matched 136 cells ✓
[IL] showCard: device d321c626-... in blade frame
[IL] 📱 Device apps: 50 detected
[IL] 🔄 Autopatch data received
[IL] showCard: card rendered ✓ size=402x602
```

---

## 📋 Roadmap

- [ ] Keyboard shortcut to pin/dismiss cards
- [ ] Export card data to clipboard (JSON / Markdown)
- [ ] Chrome Web Store / Edge Add-ons listing
- [ ] Script & remediation execution status per device
- [ ] Conditional Access policy evaluation status

---

## 🤝 Contributing

Contributions are welcome! Please open an issue first to discuss what you'd like to change.

## 📄 License

[MIT](LICENSE)

---

<p align="center">Made with ❤️ by <strong>ROBGRAME</strong> for Intune admins everywhere</p>
