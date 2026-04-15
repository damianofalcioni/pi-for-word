# Pi4Word — specification

## Product

**Pi4Word** is a Word task pane add-in that runs an AI assistant powered by [`@mariozechner/pi-agent-core`](https://github.com/badlogic/pi-mono/tree/main/packages/agent) (`pi-agent-core` in pi-mono). The assistant can read the selection and insert text into the document via Word JavaScript API tools.

## Current status

- **Implemented:** Code is organized under **`src/features/<name>/`** per **`AGENT.md`**: **`src/features/settings/`** (legacy `localStorage` shape + migration helpers), **`src/features/assistant/`** (Pi Agent + Word tools), **`src/features/task-pane/`** (Office bootstrap and **`@mariozechner/pi-web-ui`** `ChatPanel`). The bundle entry **`src/index.js`** imports **`initializeTaskPane()`** from **`src/features/task-pane/task-pane.init.js`**, which awaits **IndexedDB** `AppStorage` (`setAppStorage`), runs optional **one-time migration** from legacy `localStorage` into provider keys / Pi4Word stream-proxy settings, mounts **`pi-chat-panel`**, and opens **Settings** / **Sessions** via pi-web-ui dialogs. Model, API keys, CORS proxy, and Pi **`streamProxy`** (URL + Bearer token) are configured in **Settings** (including a **Pi4Word proxy** tab for `streamProxy`). Styles: `public/index.css` plus copied **`pi-web-ui-app.css`** (see `npm run copy:css`). Word tools `word_get_selection` and `word_insert_text`; esbuild bundles `src/index.js` to `public/index.min.js` with a **`process` shim** for browser. Colocated tests (`*.test.js`).
- **Bootstrapping:** **`initializeTaskPane()`** uses internal `scheduleOfficeBoot()` (DOM ready, **`Office.onReady`**, **2.5s fallback**). **`initPiWebStorage()`** initializes IndexedDB before UI. Chat uses **`defaultConvertToLlm`** and a composite **`streamFn`** (`streamProxy` when Pi4Word proxy is enabled, otherwise pi-web-ui CORS **`createStreamFn`**).
- **Not implemented:** Backend proxy server (client configures `streamProxy` URL + token only); production key handling beyond IndexedDB (use a secure proxy for real deployments).

## Key files

| Area | Path |
|------|------|
| Entry (esbuild bundle only; calls task pane init) | `src/index.js` |
| Task pane bootstrap (Office + agent + chat wiring) | `src/features/task-pane/task-pane.init.js` |
| Task pane shell (header, toolbar, chat mount) | `src/features/task-pane/task-pane.boundary.js` |
| Pi-web-ui storage bootstrap + migration | `src/features/task-pane/pi-web-storage.js` |
| Pi4Word streamProxy tab (Settings) | `src/features/settings/pi4word-proxy-tab.js` |
| Agent + prompts | `src/features/assistant/pi-assistant.js` |
| Word tools | `src/features/assistant/word-tools.js` |
| Settings persistence | `src/features/settings/settings-storage.js` |
| Feature docs (per `AGENT.md`) | `src/features/*/README.md` |
| Task pane shell (mount point only) | `public/index.html` |
| Styles | `public/index.css`, `public/pi-web-ui-app.css` (copied from `@mariozechner/pi-web-ui`) |
| Manifest | `manifest.xml` |

## Commands

- `npm run cert` — dev HTTPS cert (see repo scripts).
- `npm run serve` — static server for the add-in.
- `npm run word` — load add-in in Word with debugging.
- `npm run build` — tests + copy pi-web-ui CSS + esbuild bundle to `public/index.min.js`.

## Changelog

- **2026-04-15:** Moved feature roots to **`src/features/`** (was repo-root `features/`); updated `npm test` paths and entry imports.
- **2026-04-15:** Adopted **`features/`** layout: **`settings`**, **`assistant`**, **`task-pane`** with colocated `*.test.js` and feature **`README.md`**; `npm test` runs **`node --test`** with explicit paths in `package.json`.
- **2026-04-15:** Split task-pane UI into boundary module; **`src/index.js`** entry/orchestration; tests for `validateLegacySettingsForRun`.
- **2026-04-15:** Moved orchestration from **`src/index.js`** to **`src/features/task-pane/task-pane.init.js`** (`initializeTaskPane`); bundle entry only calls that feature.
- **2026-04-15:** Spec updated: minimal `index.html` with **`#app-root`**; UI in **`src/features/task-pane/`**; Office bootstrap with **`Office.onReady`** + fallback timeout; no catalog JSON or generate step.

- **2026-04-15:** Add-in display name and manifest strings updated to **Pi4Word** (ribbon, Get Started, descriptions).
- **2026-04-15:** Integrated `@mariozechner/pi-agent-core` with chat UI, Word tools, and settings storage.
- **2026-04-15:** Replaced custom chat/connection form with **`@mariozechner/pi-web-ui`** (`ChatPanel`, IndexedDB `AppStorage`, Settings/Sessions dialogs); added **`Pi4Word proxy`** tab and legacy settings migration.
