/**
 * Settings surface for the composition root: Providers tab label + Proxy tab (CORS + Pi `streamProxy`),
 * preferred chat model helpers, and pi-web-ui IndexedDB bootstrap plus one-time legacy
 * `localStorage` migration (`initPiWebStorage`, `migrateLegacyLocalStorageOnce`).
 *
 * Side effects: IndexedDB via bootstrap; legacy path reads/writes only through injected storage;
 * preferred thinking level mirrors to localStorage as an unload fallback.
 *
 * @module settings
 */
export { Pi4WordProvidersTab } from "./providers-tab.js";
export { Pi4WordProxySettingsTab } from "./pi4word-proxy-tab.js";
export {
  loadPreferredChatModel,
  persistPreferredChatModel,
  PREFERRED_MODEL_ID_KEY,
  PREFERRED_PROVIDER_KEY,
} from "./preferred-chat-model.js";
export {
  attachPreferredThinkingLevelPersistence,
  attachPreferredThinkingLevelUnloadSync,
  isThinkingLevel,
  loadPreferredThinkingLevel,
  peekPreferredThinkingLevelLocal,
  persistPreferredThinkingLevel,
  PREFERRED_THINKING_LEVEL_KEY,
  stashPreferredThinkingLevelLocal,
} from "./preferred-thinking-level.js";
export { DEFAULT_SETTINGS } from "./settings-storage.js";
export { initPiWebStorage, migrateLegacyLocalStorageOnce } from "./pi-web-bootstrap.js";
