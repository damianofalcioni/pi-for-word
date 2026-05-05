/**
 * Settings surface for the composition root: Pi `streamProxy` Settings tab (`Pi4WordProxyTab`),
 * preferred chat model helpers, and pi-web-ui IndexedDB bootstrap plus one-time legacy
 * `localStorage` migration (`initPiWebStorage`, `migrateLegacyLocalStorageOnce`).
 *
 * Side effects: IndexedDB via bootstrap; legacy path reads/writes only through injected storage.
 * Other settings helpers (e.g. legacy key shape) live in sibling modules and are not re-exported here.
 *
 * @module settings
 */
export { Pi4WordProxyTab } from "./pi4word-proxy-tab.js";
export {
  loadPreferredChatModel,
  persistPreferredChatModel,
  PREFERRED_MODEL_ID_KEY,
  PREFERRED_PROVIDER_KEY,
} from "./preferred-chat-model.js";
export { initPiWebStorage, migrateLegacyLocalStorageOnce } from "./pi-web-bootstrap.js";
