import {
  AppStorage,
  CustomProvidersStore,
  IndexedDBStorageBackend,
  ProviderKeysStore,
  SessionsStore,
  SettingsStore,
  setAppStorage,
} from "@mariozechner/pi-web-ui";
import { getModel } from "@mariozechner/pi-ai";
import { parseSettings, SETTINGS_KEY } from "./settings-storage.js";

const DB_NAME = "pi4word";
const DB_VERSION = 1;
const MIGRATION_FLAG = "pi4word.migration.v1";

function isMigrationMarkedDone() {
  return typeof sessionStorage !== "undefined" && sessionStorage.getItem(MIGRATION_FLAG);
}

function markMigrationDone() {
  if (typeof sessionStorage !== "undefined") {
    sessionStorage.setItem(MIGRATION_FLAG, "1");
  }
}

/** @returns {{ blocked: true } | { blocked: false, raw: string | null }} */
function readLegacySettingsRaw() {
  try {
    return { blocked: false, raw: localStorage.getItem(SETTINGS_KEY) };
  } catch {
    return { blocked: true };
  }
}

/**
 * @param {import("@mariozechner/pi-web-ui").AppStorage} storage
 * @param {ReturnType<typeof parseSettings>} legacy
 */
async function applyLegacySettingsToStorage(storage, legacy) {
  const key = String(legacy.apiKey || "").trim();
  if (key) {
    await storage.providerKeys.set(legacy.provider, key);
  }
  await storage.settings.set("pi4word.streamProxy.enabled", legacy.useProxy);
  await storage.settings.set("pi4word.streamProxy.url", legacy.proxyUrl);
  await storage.settings.set("pi4word.streamProxy.token", legacy.proxyToken);
}

/** @param {ReturnType<typeof parseSettings>} legacy */
function legacyModelOrUndefined(legacy) {
  try {
    return getModel(
      /** @type {import("@mariozechner/pi-ai").KnownProvider} */ (legacy.provider),
      /** @type {any} */ (legacy.modelId),
    );
  } catch {
    return undefined;
  }
}

/**
 * Creates IndexedDB-backed storage and registers the global AppStorage instance.
 * @returns {Promise<import("@mariozechner/pi-web-ui").AppStorage>}
 */
export async function initPiWebStorage() {
  const settings = new SettingsStore();
  const providerKeys = new ProviderKeysStore();
  const sessions = new SessionsStore();
  const customProviders = new CustomProvidersStore();

  const stores = [
    settings.getConfig(),
    SessionsStore.getMetadataConfig(),
    providerKeys.getConfig(),
    customProviders.getConfig(),
    sessions.getConfig(),
  ];

  const backend = new IndexedDBStorageBackend({
    dbName: DB_NAME,
    version: DB_VERSION,
    stores,
  });

  settings.setBackend(backend);
  providerKeys.setBackend(backend);
  sessions.setBackend(backend);
  customProviders.setBackend(backend);

  const storage = new AppStorage(settings, providerKeys, sessions, customProviders, backend);
  setAppStorage(storage);
  return storage;
}

/**
 * One-time migration from legacy localStorage `SETTINGS_KEY` into IndexedDB.
 * @param {import("@mariozechner/pi-web-ui").AppStorage} storage
 * @returns {Promise<import("@mariozechner/pi-ai").Model | undefined>}
 */
export async function migrateLegacyLocalStorageOnce(storage) {
  if (isMigrationMarkedDone()) {
    return undefined;
  }
  const read = readLegacySettingsRaw();
  if (read.blocked) {
    return undefined;
  }
  if (!read.raw) {
    markMigrationDone();
    return undefined;
  }

  const legacy = parseSettings(read.raw);
  try {
    await applyLegacySettingsToStorage(storage, legacy);
  } catch (e) {
    console.warn("[pi4word] could not migrate legacy settings:", e);
    return undefined;
  }

  markMigrationDone();
  return legacyModelOrUndefined(legacy);
}
