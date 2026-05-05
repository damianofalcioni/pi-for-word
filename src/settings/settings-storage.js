/** @typedef {{
 *   provider: string,
 *   modelId: string,
 *   apiKey: string,
 *   useProxy: boolean,
 *   proxyUrl: string,
 *   proxyToken: string,
 * }} PiWordSettings */

export const SETTINGS_KEY = "pi4word.settings.v1";

/** @type {PiWordSettings} */
export const DEFAULT_SETTINGS = {
  provider: "openai",
  modelId: "gpt-5.5",
  apiKey: "",
  useProxy: false,
  proxyUrl: "",
  proxyToken: "",
};

/**
 * @param {string | null | undefined} json
 * @returns {PiWordSettings}
 */
export function parseSettings(json) {
  if (!json) {
    return { ...DEFAULT_SETTINGS };
  }
  try {
    const data = JSON.parse(json);
    return {
      ...DEFAULT_SETTINGS,
      ...data,
      provider: typeof data.provider === "string" ? data.provider : DEFAULT_SETTINGS.provider,
      modelId: typeof data.modelId === "string" ? data.modelId : DEFAULT_SETTINGS.modelId,
      apiKey: typeof data.apiKey === "string" ? data.apiKey : "",
      useProxy: Boolean(data.useProxy),
      proxyUrl: typeof data.proxyUrl === "string" ? data.proxyUrl : "",
      proxyToken: typeof data.proxyToken === "string" ? data.proxyToken : "",
    };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

/**
 * @param {PiWordSettings} settings
 * @returns {string}
 */
export function serializeSettings(settings) {
  return JSON.stringify({
    provider: settings.provider,
    modelId: settings.modelId,
    apiKey: settings.apiKey,
    useProxy: settings.useProxy,
    proxyUrl: settings.proxyUrl,
    proxyToken: settings.proxyToken,
  });
}

/**
 * @param {Storage} storage
 * @returns {PiWordSettings}
 */
export function loadSettings(storage) {
  return parseSettings(storage.getItem(SETTINGS_KEY));
}

/**
 * @param {Storage} storage
 * @param {PiWordSettings} settings
 */
export function saveSettings(storage, settings) {
  storage.setItem(SETTINGS_KEY, serializeSettings(settings));
}
