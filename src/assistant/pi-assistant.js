import { Agent, streamProxy } from "@mariozechner/pi-agent-core";
import { getModel } from "@mariozechner/pi-ai";
import { createStreamFn, defaultConvertToLlm, getAppStorage } from "@mariozechner/pi-web-ui";
import {
  attachPreferredThinkingLevelPersistence,
  DEFAULT_SETTINGS,
  isThinkingLevel,
} from "../settings/index.js";

export const SYSTEM_PROMPT = `You are an AI assistant embedded in Microsoft Word. You help the user draft and edit documents.
Be concise and prefer to use the available tools when possible.`;

/**
 * Builds a stream function that prefers Pi Agent `streamProxy` when enabled in settings,
 * otherwise uses pi-web-ui CORS proxy handling via `createStreamFn`.
 */
export function createPi4WordStreamFn() {
  const corsStreamFn = createStreamFn(async () => {
    const storage = getAppStorage();
    const enabled = await storage.settings.get("proxy.enabled");
    return enabled ? (await storage.settings.get("proxy.url")) || undefined : undefined;
  });
  return async (model, context, options) => {
    const storage = getAppStorage();
    const useStreamProxy = await storage.settings.get("pi4word.streamProxy.enabled");
    const proxyUrl = String((await storage.settings.get("pi4word.streamProxy.url")) ?? "").trim();
    const proxyToken = String((await storage.settings.get("pi4word.streamProxy.token")) ?? "").trim();
    if (useStreamProxy && proxyUrl && proxyToken) {
      return streamProxy(model, context, {
        ...options,
        proxyUrl: proxyUrl.replace(/\/$/, ""),
        authToken: proxyToken,
      });
    }
    return corsStreamFn(model, context, options);
  };
}

/**
 * @param {import("@mariozechner/pi-ai").Model} model
 * @param {{ thinkingLevel?: import("@mariozechner/pi-agent-core").ThinkingLevel }} [opts]
 * @returns {import("@mariozechner/pi-agent-core").Agent}
 */
export function createWordAgent(model, opts = {}) {
  const streamFn = createPi4WordStreamFn();
  const agent = new Agent({
    initialState: {
      systemPrompt: SYSTEM_PROMPT,
      model,
      thinkingLevel: isThinkingLevel(opts.thinkingLevel) ? opts.thinkingLevel : "off",
      messages: [],
      tools: [],
    },
    convertToLlm: defaultConvertToLlm,
    streamFn,
    getApiKey: async (provider) => {
      const key = await getAppStorage().providerKeys.get(provider);
      return key ?? undefined;
    },
  });
  attachPreferredThinkingLevelPersistence(agent);
  return agent;
}

/**
 * Default model when no migration or session provides one.
 */
export function getDefaultWordModel() {
  return getModel(
    /** @type {import("@mariozechner/pi-ai").KnownProvider} */ (DEFAULT_SETTINGS.provider),
    /** @type {any} */ (DEFAULT_SETTINGS.modelId),
  );
}

/**
 * Restore an agent from persisted session data (e.g. SessionsStore).
 * @param {{ model: import("@mariozechner/pi-ai").Model, thinkingLevel?: import("@mariozechner/pi-agent-core").ThinkingLevel, messages?: import("@mariozechner/pi-agent-core").AgentMessage[] }} partial
 * @returns {import("@mariozechner/pi-agent-core").Agent}
 */
export function createWordAgentFromSession(partial) {
  const streamFn = createPi4WordStreamFn();
  const agent = new Agent({
    initialState: {
      systemPrompt: SYSTEM_PROMPT,
      model: partial.model,
      thinkingLevel: isThinkingLevel(partial.thinkingLevel) ? partial.thinkingLevel : "off",
      messages: partial.messages ?? [],
      tools: [],
    },
    convertToLlm: defaultConvertToLlm,
    streamFn,
    getApiKey: async (provider) => {
      const key = await getAppStorage().providerKeys.get(provider);
      return key ?? undefined;
    },
  });
  attachPreferredThinkingLevelPersistence(agent);
  return agent;
}

export { createWordTools } from "../assistant-tools/index.js";
