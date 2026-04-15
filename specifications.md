# Pi4Word — specification

## Product

**Pi4Word** is a Word task pane add-in that runs an AI assistant powered by [`@mariozechner/pi-agent-core`](https://github.com/badlogic/pi-mono/tree/main/packages/agent) (`pi-agent-core` in pi-mono). The assistant can read the selection and insert text into the document via Word JavaScript API tools.

## Current status

- **Implemented:** Task pane UI built entirely in JavaScript (`renderApp()` in `src/index.js`); static shell in `public/index.html` with a single mount node **`#app-root`** plus Office.js and the bundled script; chat transcript with streaming assistant text; connection form using **provider** and **model** `<select>` lists filled from `@mariozechner/pi-ai` (`getProviders()` / `getModels()`), API key, optional `streamProxy` URL + token persisted in `localStorage`; Pi Agent integration with Word tools `word_get_selection` and `word_insert_text`; build bundles `src/index.js` with esbuild to `public/index.min.js`.
- **Bootstrapping:** `scheduleOfficeBoot()` waits for DOM readiness, registers `Office.onReady` when available, and uses a **timeout fallback** so the UI still mounts if `onReady` is late or Office.js is restricted; host detection avoids touching `Office.HostType` when `Office` is undefined; `localStorage` read errors are caught when applying saved settings. `<select>` options are built with DOM node removal (no `innerHTML` on `<select>`) for compatibility with the Word task-pane WebView.
- **Not implemented:** Backend proxy server (client only supports configuring proxy mode); production key handling beyond local storage (use a secure proxy for real deployments).

## Key files

| Area | Path |
|------|------|
| Entry / UI | `src/index.js` |
| Agent + prompts | `src/pi-assistant.js` |
| Word tools | `src/word-tools.js` |
| Settings persistence | `src/settings-storage.js` |
| Task pane shell (mount point only) | `public/index.html` |
| Styles | `public/index.css` |
| Manifest | `manifest.xml` |

## Commands

- `npm run cert` — dev HTTPS cert (see repo scripts).
- `npm run serve` — static server for the add-in.
- `npm run word` — load add-in in Word with debugging.
- `npm run build` — tests + esbuild bundle to `public/index.min.js`.

## Changelog

- **2026-04-15:** Spec updated: minimal `index.html` with **`#app-root`**; UI and provider/model lists implemented in **`src/index.js`** via `@mariozechner/pi-ai`; Office bootstrap with **`Office.onReady`** + fallback timeout; no catalog JSON or generate step.

- **2026-04-15:** Add-in display name and manifest strings updated to **Pi4Word** (ribbon, Get Started, descriptions).
- **2026-04-15:** Integrated `@mariozechner/pi-agent-core` with chat UI, Word tools, and settings storage.
