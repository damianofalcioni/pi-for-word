/**
 * Validation for legacy `PiWordSettings` (localStorage migration / docs).
 * @param {import("./settings-storage.js").PiWordSettings} settings
 * @returns {string | null} Error message, or null if OK to run against providers.
 */
export function validateLegacySettingsForRun(settings) {
  const useProxy =
    settings.useProxy &&
    settings.proxyUrl.trim().length > 0 &&
    settings.proxyToken.trim().length > 0;
  if (!useProxy && settings.apiKey.trim().length === 0) {
    return "Add an API key, or enable proxy mode with URL and token.";
  }
  if (settings.useProxy && (!settings.proxyUrl.trim() || !settings.proxyToken.trim())) {
    return "Proxy mode needs both proxy URL and token.";
  }
  return null;
}
