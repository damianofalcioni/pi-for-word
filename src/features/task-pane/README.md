# task-pane

Purpose: **`task-pane.boundary.js`** builds the minimal shell (`renderApp`, `setStatus`) — header, status strip, toolbar (**Settings**, **Sessions**, **New chat**), and **`#chatMount`** for **`pi-chat-panel`**. **`task-pane.init.js`** is the entry: **`initializeTaskPane()`** runs Office bootstrap, **`await initPiWebStorage()`**, **`migrateLegacyLocalStorageOnce()`**, **`ChatPanel`** + **`setAgent()`**, session autosave, and dialog wiring.

**`pi-web-storage.js`** — IndexedDB `AppStorage` + `setAppStorage`, optional one-time migration from legacy `localStorage` (`SETTINGS_KEY`).

Public API — boundary: `renderApp`, `setStatus`. Entry: `initializeTaskPane`.

Side effects: **`task-pane.init.js`** — Office.js, IndexedDB, **`@mariozechner/pi-web-ui`** dialogs, Pi agent `subscribe` for session save.

Flow: `initializeTaskPane` → `initPiWebStorage` → `renderApp` → mount `ChatPanel` → `setAgent` with Word tools + `ApiKeyPromptDialog` → Settings / Sessions buttons.

Tests: `src/features/settings/legacy-settings-validation.test.js` (legacy `PiWordSettings` validation used by migration semantics).
