import { getModel } from "@mariozechner/pi-ai";
import { getAppStorage } from "@mariozechner/pi-web-ui";

export const PREFERRED_PROVIDER_KEY = "pi4word.chat.preferredProvider";
export const PREFERRED_MODEL_ID_KEY = "pi4word.chat.preferredModelId";

/**
 * @param {import("@mariozechner/pi-web-ui").AppStorage} storage
 * @param {string} provider
 * @param {string} modelId
 */
export async function setPreferredChatModelKeys(storage, provider, modelId) {
  await storage.settings.set(PREFERRED_PROVIDER_KEY, provider);
  await storage.settings.set(PREFERRED_MODEL_ID_KEY, modelId);
}

/**
 * @returns {Promise<import("@mariozechner/pi-ai").Model | undefined>}
 */
export async function loadPreferredChatModel() {
  const storage = getAppStorage();
  const provider = await storage.settings.get(PREFERRED_PROVIDER_KEY);
  const modelId = await storage.settings.get(PREFERRED_MODEL_ID_KEY);
  if (typeof provider !== "string" || !provider.trim()) {
    return undefined;
  }
  if (typeof modelId !== "string" || !modelId.trim()) {
    return undefined;
  }
  try {
    return getModel(
      /** @type {import("@mariozechner/pi-ai").KnownProvider} */ (provider),
      /** @type {any} */ (modelId),
    );
  } catch {
    return undefined;
  }
}

/**
 * @param {import("@mariozechner/pi-ai").Model} model
 */
export async function persistPreferredChatModel(model) {
  await setPreferredChatModelKeys(getAppStorage(), String(model.provider), model.id);
}
