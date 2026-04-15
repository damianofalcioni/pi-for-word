# assistant

Purpose: Build a **pi-agent-core** `Agent` for Word: model from `pi-ai`, **`defaultConvertToLlm`** from **`@mariozechner/pi-web-ui`**, composite **`streamFn`** (Pi **`streamProxy`** when Pi4Word proxy settings are enabled, otherwise pi-web-ui CORS **`createStreamFn`**), **`getApiKey`** from **`getAppStorage().providerKeys`**, Word tools supplied via **`ChatPanel`** `toolsFactory`.

Public API: `createWordAgent(model)`, `createWordAgentFromSession(partial)`, `createPi4WordStreamFn()`, `getDefaultWordModel()`, `SYSTEM_PROMPT`; `createWordTools` from `word-tools.js`.

Side effects: Network via agent streaming; Word API only inside tool execution when hosted in Word.

State: None in module; credentials and proxy flags come from **`AppStorage`** (IndexedDB).

Flow: `createWordAgent` → `new Agent` → **`ChatPanel.setAgent`** sets artifacts + `toolsFactory` → Word tools merged with artifact tool.

Tests: none in this feature (browser/Office-dependent); Word tools exercised at runtime in Word.
