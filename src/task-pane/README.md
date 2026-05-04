# task-pane

Purpose: **`task-pane.boundary.js`** builds the minimal shell (`renderApp`, `setStatus`) — header, status strip, toolbar (**Settings**, **Sessions**, **New chat**), and **`#chatMount`** for **`pi-chat-panel`**. **`task-pane.init.js`** is the composition entry: **`initializeTaskPane()`** runs Office bootstrap, awaits **`initPiWebStorage()`** / **`migrateLegacyLocalStorageOnce()`** (from **`settings`**, see `../settings/pi-web-bootstrap.js`), **`ChatPanel`** + **`setAgent()`**, session autosave (**`task-pane.session.js`**, with titles from **`task-pane.chat-title.model.js`**), and dialog wiring (**`task-pane.wiring.js`**). **`task-pane.office.js`** holds Office host detection and **`Office.onReady`** / fallback scheduling.

Public API — boundary: `renderApp`, `setStatus`. Entry: `initializeTaskPane`.

Side effects: **`task-pane.init.js`** — Office.js, IndexedDB (via settings bootstrap), **`@mariozechner/pi-web-ui`** dialogs, Pi agent `subscribe` for session save.

Flow: `initializeTaskPane` → storage bootstrap → `renderApp` → mount `ChatPanel` → `setAgent` with Word tools + `ApiKeyPromptDialog` → Settings / Sessions buttons.

Cross-feature imports: only through **`../assistant/index.js`** and **`../settings/index.js`** (composition root).

Tests: `task-pane.chat-title.model.test.js` (title preview strings); `src/settings/legacy-settings-validation.test.js` (legacy settings validation used by migration).
