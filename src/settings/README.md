# settings

Purpose: **Legacy** `PiWordSettings` shape and `localStorage` helpers (`SETTINGS_KEY`, `parseSettings`, `serializeSettings`, etc.) used for **one-time migration** into **`@mariozechner/pi-web-ui`** IndexedDB stores via **`pi-web-bootstrap.js`** (`initPiWebStorage`, **`migrateLegacyLocalStorageOnce`**). **`legacy-settings-validation.js`** holds the same validation rules as the old connection form for tests and reference.

**`pi4word-proxy-tab.js`** — custom Settings tab for Pi Agent **`streamProxy`** (base URL + Bearer token), separate from the pi-web-ui **CORS Proxy** tab.

**`index.js`** — stable exports for the composition root (`Pi4WordProxyTab`, storage bootstrap).

Public API: `SETTINGS_KEY`, `DEFAULT_SETTINGS`, `parseSettings`, `serializeSettings`, `loadSettings`, `saveSettings`; `validateLegacySettingsForRun` in `legacy-settings-validation.js`; **`initPiWebStorage`**, **`migrateLegacyLocalStorageOnce`** from **`pi-web-bootstrap.js`**.

Side effects: read/write `Storage` only via the `storage` argument passed in (legacy path); IndexedDB via **`pi-web-bootstrap`**.

Tests: `settings-storage.test.js`, `legacy-settings-validation.test.js`.
