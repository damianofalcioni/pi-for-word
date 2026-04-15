# task-pane

Purpose: Task-pane **boundary** (`task-pane.boundary.js`) — build DOM (`renderApp`), provider/model selects, form ⇄ `PiWordSettings`, chat line elements, connection validation messages. **`task-pane.init.js`** is the **entry** for the pane: `initializeTaskPane()` runs Office bootstrap and wires the agent and chat to those controls.

Public API — boundary: `renderApp`, `setStatus`, `fillProviderSelect`, `fillModelSelect`, `readSettingsFromForm`, `applySettingsToForm`, `appendUserMessage`, `appendAssistantShell`, `appendSystemLine`, `validateSettingsForRun`. Entry: `initializeTaskPane`.

Side effects: **`task-pane.boundary.js`** — `document` and `@mariozechner/pi-ai` catalog lists. **`task-pane.init.js`** — Office.js, `localStorage` (via settings), Pi agent subscribe/abort.

State: None global; reads/writes existing elements by id after `renderApp`.

Flow: `initializeTaskPane` → mount + `renderApp` → user input → `readSettingsFromForm` / `validateSettingsForRun` → agent (`pi-assistant`).

Tests: `task-pane.boundary.test.js` (`validateSettingsForRun`).
