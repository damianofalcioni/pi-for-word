import { Agent, streamProxy } from "@mariozechner/pi-agent-core";
import { getModel } from "@mariozechner/pi-ai";
import { createWordTools } from "./word-tools.js";

const SYSTEM_PROMPT = `You are an AI assistant embedded in Microsoft Word. You help the user draft and edit documents.
Use the provided tools to read the current selection and to insert or replace text when the user asks you to change the document.
Be concise; prefer applying edits via tools when appropriate instead of only describing what to type.`;

/**
 * @param {import("./settings-storage.js").PiWordSettings} settings
 * @returns {import("@mariozechner/pi-agent-core").Agent}
 */
export function createWordAgent(settings) {
  const model = getModel(
    /** @type {import("@mariozechner/pi-ai").KnownProvider} */ (settings.provider),
    /** @type {any} */ (settings.modelId),
  );

  const useProxy =
    settings.useProxy &&
    String(settings.proxyUrl || "").trim().length > 0 &&
    String(settings.proxyToken || "").trim().length > 0;

  /** @type {import("@mariozechner/pi-agent-core").StreamFn | undefined} */
  let streamFn;
  if (useProxy) {
    const proxyUrl = String(settings.proxyUrl).trim().replace(/\/$/, "");
    const authToken = String(settings.proxyToken).trim();
    streamFn = (m, ctx, opt) =>
      streamProxy(m, ctx, {
        ...opt,
        authToken,
        proxyUrl,
      });
  }

  const tools =
    typeof Word !== "undefined" ? createWordTools() : [];

  const agent = new Agent({
    initialState: {
      systemPrompt: SYSTEM_PROMPT,
      model,
      tools,
    },
    streamFn,
    getApiKey: () => {
      if (useProxy) {
        return undefined;
      }
      const key = String(settings.apiKey || "").trim();
      return key || undefined;
    },
  });

  return agent;
}

export { SYSTEM_PROMPT };
