import { Agent, streamProxy } from "@mariozechner/pi-agent-core";
import { getModel } from "@mariozechner/pi-ai";
import { createStreamFn, defaultConvertToLlm, getAppStorage } from "@mariozechner/pi-web-ui";
import { createWordTools } from "./word-tools.js";

export const SYSTEM_PROMPT = `You are an AI assistant embedded in Microsoft Word. You help the user draft and edit documents.
Use the provided tools to read the current selection and to insert or replace text when the user asks you to change the document.
Be concise; prefer applying edits via tools when appropriate instead of only describing what to type.`;

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
 * @returns {import("@mariozechner/pi-agent-core").Agent}
 */
export function createWordAgent(model) {
  const streamFn = createPi4WordStreamFn();
  return new Agent({
    initialState: {
      systemPrompt: SYSTEM_PROMPT,
      model,
      thinkingLevel: "off",
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
}

/**
 * Default model when no migration or session provides one.
 */
export function getDefaultWordModel() {
  return getModel("anthropic", "claude-sonnet-4-20250514");
}

/**
 * Restore an agent from persisted session data (e.g. SessionsStore).
 * @param {{ model: import("@mariozechner/pi-ai").Model, thinkingLevel?: import("@mariozechner/pi-agent-core").ThinkingLevel, messages?: import("@mariozechner/pi-agent-core").AgentMessage[] }} partial
 * @returns {import("@mariozechner/pi-agent-core").Agent}
 */
export function createWordAgentFromSession(partial) {
  const streamFn = createPi4WordStreamFn();
  return new Agent({
    initialState: {
      systemPrompt: SYSTEM_PROMPT,
      model: partial.model,
      thinkingLevel: partial.thinkingLevel ?? "off",
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
}

export { createWordTools };
