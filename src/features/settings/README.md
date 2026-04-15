# settings

Purpose: Serialize and persist **PiWordSettings** (provider, model, API key, optional proxy) in `localStorage`.

Public API: `SETTINGS_KEY`, `DEFAULT_SETTINGS`, `parseSettings`, `serializeSettings`, `loadSettings`, `saveSettings` — inputs/outputs are JSON strings and `PiWordSettings` objects.

Side effects: read/write `Storage` only via the `storage` argument passed in.

State: None global; callers inject `localStorage` (or test doubles).

Flow: `loadSettings` → `parseSettings` → merged object; `saveSettings` → `serializeSettings` → `setItem`.

Tests: `settings-storage.test.js` (parse/serialize/roundtrip).
