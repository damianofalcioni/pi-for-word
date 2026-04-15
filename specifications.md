# Pi4Word — specification

## Product

**Pi4Word** is a Word task pane add-in that runs an AI assistant powered by [`@mariozechner/pi-agent-core`](https://github.com/badlogic/pi-mono/tree/main/packages/agent) (`pi-agent-core` in pi-mono). The assistant can read the selection and insert text into the document via Word JavaScript API tools.

## Current status

- **Implemented:** Code is organized under **`src/features/<name>/`** per **`AGENT.md`**: **`src/features/settings/`** (localStorage serialization), **`src/features/assistant/`** (Pi Agent + Word tools), **`src/features/task-pane/`** (task-pane DOM, form boundary, and UI bootstrap). The bundle entry **`src/index.js`** only imports and calls **`initializeTaskPane()`** from **`src/features/task-pane/task-pane.init.js`**, which runs Office bootstrap, agent lifecycle, and event wiring. Static shell in `public/index.html` with **`#app-root`** plus Office.js and the bundled script; chat transcript with streaming assistant text; connection form using **provider** and **model** `<select>` lists from `@mariozechner/pi-ai`, API key, optional `streamProxy` URL + token in `localStorage`; Word tools `word_get_selection` and `word_insert_text`; esbuild bundles `src/index.js` to `public/index.min.js`. Colocated tests live next to each feature (`*.test.js`).
- **Bootstrapping:** **`initializeTaskPane()`** (in **`task-pane.init.js`**) calls internal `scheduleOfficeBoot()`, which waits for DOM readiness, registers `Office.onReady` when available, and uses a **timeout fallback** so the UI still mounts if `onReady` is late or Office.js is restricted; host detection avoids touching `Office.HostType` when `Office` is undefined; `localStorage` read errors are caught when applying saved settings. `<select>` options are built with DOM node removal (no `innerHTML` on `<select>`) for compatibility with the Word task-pane WebView.
- **Not implemented:** Backend proxy server (client only supports configuring proxy mode); production key handling beyond local storage (use a secure proxy for real deployments).

## Key files

| Area | Path |
|------|------|
| Entry (esbuild bundle only; calls task pane init) | `src/index.js` |
| Task pane bootstrap (Office + agent + chat wiring) | `src/features/task-pane/task-pane.init.js` |
| Task pane UI (DOM, selects, chat lines) | `src/features/task-pane/task-pane.boundary.js` |
| Agent + prompts | `src/features/assistant/pi-assistant.js` |
| Word tools | `src/features/assistant/word-tools.js` |
| Settings persistence | `src/features/settings/settings-storage.js` |
| Feature docs (per `AGENT.md`) | `src/features/*/README.md` |
| Task pane shell (mount point only) | `public/index.html` |
| Styles | `public/index.css` |
| Manifest | `manifest.xml` |

## Commands

- `npm run cert` — dev HTTPS cert (see repo scripts).
- `npm run serve` — static server for the add-in.
- `npm run word` — load add-in in Word with debugging.
- `npm run build` — tests + esbuild bundle to `public/index.min.js`.

## Changelog

- **2026-04-15:** Moved feature roots to **`src/features/`** (was repo-root `features/`); updated `npm test` paths and entry imports.
- **2026-04-15:** Adopted **`features/`** layout: **`settings`**, **`assistant`**, **`task-pane`** with colocated `*.test.js` and feature **`README.md`**; `npm test` runs **`node --test`** with explicit paths in `package.json`.
- **2026-04-15:** Split task-pane UI into boundary module; **`src/index.js`** entry/orchestration; tests for `validateSettingsForRun`.
- **2026-04-15:** Moved orchestration from **`src/index.js`** to **`src/features/task-pane/task-pane.init.js`** (`initializeTaskPane`); bundle entry only calls that feature.
- **2026-04-15:** Spec updated: minimal `index.html` with **`#app-root`**; UI in **`src/features/task-pane/`**; Office bootstrap with **`Office.onReady`** + fallback timeout; no catalog JSON or generate step.

- **2026-04-15:** Add-in display name and manifest strings updated to **Pi4Word** (ribbon, Get Started, descriptions).
- **2026-04-15:** Integrated `@mariozechner/pi-agent-core` with chat UI, Word tools, and settings storage.
