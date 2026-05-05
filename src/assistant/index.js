/**
 * Pi-agent-core Agent for Word: model selection, `defaultConvertToLlm`, composite `streamFn`
 * (Pi `streamProxy` when Pi4Word proxy settings are enabled, otherwise pi-web-ui CORS streaming),
 * API keys from `getAppStorage().providerKeys`, and Word tools for ChatPanel `toolsFactory`.
 *
 * Side effects: network via streaming; Word.js only inside tool execution when hosted in Word.
 * No module state — credentials come from IndexedDB-backed AppStorage.
 *
 * @module assistant
 */
export {
  SYSTEM_PROMPT,
  createPi4WordStreamFn,
  createWordAgent,
  createWordAgentFromSession,
  getDefaultWordModel,
  createWordTools,
} from "./pi-assistant.js";
