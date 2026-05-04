# assistant

Purpose: Build a **pi-agent-core** `Agent` for Word: model from `pi-ai`, **`defaultConvertToLlm`** from **`@mariozechner/pi-web-ui`**, composite **`streamFn`** (Pi **`streamProxy`** when Pi4Word proxy settings are enabled, otherwise pi-web-ui CORS **`createStreamFn`**), **`getApiKey`** from **`getAppStorage().providerKeys`**, Word tools supplied via **`ChatPanel`** `toolsFactory`.

**`index.js`** re-exports the stable public surface for other features (`createWordAgent`, `createWordAgentFromSession`, `getDefaultWordModel`, `createWordTools`, `createPi4WordStreamFn`, `SYSTEM_PROMPT`). Implementations live in **`pi-assistant.js`** and **`word-tools.js`**.

Public API: same as **`index.js`** exports; see **`pi-assistant.js`** for details.

Side effects: Network via agent streaming; Word API only inside tool execution when hosted in Word.

State: None in module; credentials and proxy flags come from **`AppStorage`** (IndexedDB).

Flow: `createWordAgent` → `new Agent` → **`ChatPanel.setAgent`** sets artifacts + `toolsFactory` → Word tools merged with artifact tool.

Tests: none in this feature (browser/Office-dependent); Word tools exercised at runtime in Word.
