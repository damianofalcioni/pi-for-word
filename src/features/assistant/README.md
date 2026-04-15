# assistant

Purpose: Build a **pi-agent-core** `Agent` for Word: model from `pi-ai`, optional `streamProxy`, Word tools when `Word` is defined.

Public API: `createWordAgent(settings)` → `Agent`; `SYSTEM_PROMPT` export for tests or reuse.

Side effects: Network via agent streaming; Word API only inside tool execution when hosted in Word.

State: None in module; credentials come from the `settings` argument per call.

Flow: `settings` → `getModel` + tools → `new Agent` → returned instance.

Tests: none yet (browser/Office-dependent); Word tools exercised at runtime in Word.
