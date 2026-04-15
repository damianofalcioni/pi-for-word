import { getModels, getProviders } from "@mariozechner/pi-ai";
import { createWordAgent } from "./pi-assistant.js";
import {
  DEFAULT_SETTINGS,
  loadSettings,
  saveSettings,
  serializeSettings,
} from "./settings-storage.js";

/**
 * @param {HTMLSelectElement} select
 */
function clearSelect(select) {
  while (select.firstChild) {
    select.removeChild(select.firstChild);
  }
}

/**
 * @param {string} id
 */
function formatProviderLabel(id) {
  return id
    .split("-")
    .map((w) => w.slice(0, 1).toUpperCase() + w.slice(1))
    .join(" ");
}

/**
 * @param {HTMLSelectElement} select
 * @param {string} [preferredId]
 */
function fillProviderSelect(select, preferredId) {
  const providers = getProviders();
  clearSelect(select);
  for (const id of providers) {
    const opt = document.createElement("option");
    opt.value = id;
    opt.textContent = formatProviderLabel(id);
    select.appendChild(opt);
  }
  const chosen = providers.includes(preferredId ?? "")
    ? preferredId
    : providers.includes(DEFAULT_SETTINGS.provider)
      ? DEFAULT_SETTINGS.provider
      : providers[0];
  if (chosen) {
    select.value = chosen;
  }
}

/**
 * @param {HTMLSelectElement} select
 * @param {string} provider
 * @param {string} [preferredModelId] If omitted, selects the first model for the provider.
 */
function fillModelSelect(select, provider, preferredModelId) {
  const models = getModels(/** @type {import("@mariozechner/pi-ai").KnownProvider} */ (provider));
  const ids = new Set(models.map((m) => m.id));
  clearSelect(select);
  for (const m of models) {
    const opt = document.createElement("option");
    opt.value = m.id;
    opt.textContent = m.name && m.name !== m.id ? `${m.name} — ${m.id}` : m.id;
    select.appendChild(opt);
  }
  if (preferredModelId && !ids.has(preferredModelId)) {
    const opt = document.createElement("option");
    opt.value = preferredModelId;
    opt.textContent = `${preferredModelId} (saved; not in catalog)`;
    select.insertBefore(opt, select.firstChild);
    select.value = preferredModelId;
    return;
  }
  const pick =
    preferredModelId && ids.has(preferredModelId) ? preferredModelId : (models[0]?.id ?? "");
  if (pick) {
    select.value = pick;
  }
}

function setStatus(message, isError) {
  const status = document.getElementById("hostStatus");
  if (!status) {
    return;
  }
  status.textContent = message;
  status.style.background = isError ? "#f7d8d2" : "#edf4f4";
  status.style.color = isError ? "#6a1b16" : "#134243";
}

/**
 * Builds the task-pane DOM (stylesheet is linked from index.html).
 */
function renderApp() {
  const header = document.createElement("header");
  header.className = "app-header";
  const headerLeft = document.createElement("div");
  const title = document.createElement("div");
  title.className = "app-title";
  title.textContent = "Pi4Word";
  const subtitle = document.createElement("div");
  subtitle.className = "app-subtitle";
  subtitle.textContent = "Pi Agent assistant for Word";
  headerLeft.append(title, subtitle);
  const hostStatus = document.createElement("div");
  hostStatus.className = "status";
  hostStatus.id = "hostStatus";
  hostStatus.textContent = "Loading Office...";
  header.append(headerLeft, hostStatus);

  const main = document.createElement("main");
  main.id = "mainDiv";
  main.className = "hello-main";

  const details = document.createElement("details");
  details.className = "panel connection-details";
  const summary = document.createElement("summary");
  summary.textContent = "Connection (model & API)";
  const hint = document.createElement("p");
  hint.className = "connection-hint";
  hint.append(
    "Direct mode sends requests from this pane using your key (suitable for local dev only). For production, use a proxy and ",
  );
  const code = document.createElement("code");
  code.textContent = "streamProxy";
  hint.append(code, " — see ");
  const a = document.createElement("a");
  a.href = "https://github.com/badlogic/pi-mono/tree/main/packages/agent";
  a.target = "_blank";
  a.rel = "noreferrer";
  a.textContent = "pi-agent-core";
  hint.append(a, ".");

  const labProv = document.createElement("label");
  labProv.append("Provider\n");
  const fieldProvider = document.createElement("select");
  fieldProvider.id = "fieldProvider";
  labProv.appendChild(fieldProvider);

  const labModel = document.createElement("label");
  labModel.append("Model\n");
  const fieldModelId = document.createElement("select");
  fieldModelId.id = "fieldModelId";
  labModel.appendChild(fieldModelId);

  const labKey = document.createElement("label");
  labKey.append("API key (direct mode)\n");
  const fieldApiKey = document.createElement("input");
  fieldApiKey.id = "fieldApiKey";
  fieldApiKey.type = "password";
  fieldApiKey.autocomplete = "off";
  fieldApiKey.placeholder = "Required if not using proxy";
  labKey.appendChild(fieldApiKey);

  const labProxy = document.createElement("label");
  labProxy.className = "row-check";
  const fieldUseProxy = document.createElement("input");
  fieldUseProxy.id = "fieldUseProxy";
  fieldUseProxy.type = "checkbox";
  labProxy.append(fieldUseProxy, " Use proxy (");
  const c2 = document.createElement("code");
  c2.textContent = "streamProxy";
  labProxy.append(c2, ")");

  const labProxyUrl = document.createElement("label");
  labProxyUrl.append("Proxy base URL\n");
  const fieldProxyUrl = document.createElement("input");
  fieldProxyUrl.id = "fieldProxyUrl";
  fieldProxyUrl.type = "url";
  fieldProxyUrl.autocomplete = "off";
  fieldProxyUrl.placeholder = "https://your-proxy.example.com";
  labProxyUrl.appendChild(fieldProxyUrl);

  const labProxyTok = document.createElement("label");
  labProxyTok.append("Proxy token\n");
  const fieldProxyToken = document.createElement("input");
  fieldProxyToken.id = "fieldProxyToken";
  fieldProxyToken.type = "password";
  fieldProxyToken.autocomplete = "off";
  labProxyTok.appendChild(fieldProxyToken);

  const saveRow = document.createElement("div");
  saveRow.className = "actions";
  const saveConnectionBtn = document.createElement("button");
  saveConnectionBtn.type = "button";
  saveConnectionBtn.id = "saveConnectionBtn";
  saveConnectionBtn.className = "primary";
  saveConnectionBtn.textContent = "Save connection";
  saveRow.appendChild(saveConnectionBtn);

  details.append(
    summary,
    hint,
    labProv,
    labModel,
    labKey,
    labProxy,
    labProxyUrl,
    labProxyTok,
    saveRow,
  );

  const chatSection = document.createElement("section");
  chatSection.className = "panel chat-section";
  chatSection.setAttribute("aria-label", "Chat");
  const h2 = document.createElement("h2");
  h2.textContent = "Assistant";
  const chatLog = document.createElement("div");
  chatLog.id = "chatLog";
  chatLog.className = "chat-log";
  chatLog.setAttribute("role", "log");
  chatLog.setAttribute("aria-live", "polite");
  const labMsg = document.createElement("label");
  labMsg.htmlFor = "chatInput";
  labMsg.textContent = "Message";
  const chatInput = document.createElement("textarea");
  chatInput.id = "chatInput";
  chatInput.rows = 3;
  chatInput.placeholder = "Ask a question or request an edit…";
  const actions = document.createElement("div");
  actions.className = "actions";
  const sendBtn = document.createElement("button");
  sendBtn.type = "button";
  sendBtn.id = "sendBtn";
  sendBtn.className = "primary";
  sendBtn.textContent = "Send";
  const stopBtn = document.createElement("button");
  stopBtn.type = "button";
  stopBtn.id = "stopBtn";
  stopBtn.disabled = true;
  stopBtn.textContent = "Stop";
  const clearChatBtn = document.createElement("button");
  clearChatBtn.type = "button";
  clearChatBtn.id = "clearChatBtn";
  clearChatBtn.textContent = "Clear chat";
  actions.append(sendBtn, stopBtn, clearChatBtn);

  chatSection.append(h2, chatLog, labMsg, chatInput, actions);
  main.append(details, chatSection);

  const mount = document.getElementById("app-root");
  if (!mount) {
    throw new Error("Pi4Word: #app-root is missing from index.html");
  }
  mount.append(header, main);
}

/**
 * @returns {import("./settings-storage.js").PiWordSettings}
 */
function readSettingsFromForm() {
  const provider = /** @type {HTMLSelectElement} */ (document.getElementById("fieldProvider"))
    .value;
  const modelId = /** @type {HTMLSelectElement} */ (document.getElementById("fieldModelId")).value
    .trim();
  const apiKey = /** @type {HTMLInputElement} */ (document.getElementById("fieldApiKey")).value;
  const useProxy = /** @type {HTMLInputElement} */ (document.getElementById("fieldUseProxy")).checked;
  const proxyUrl = /** @type {HTMLInputElement} */ (document.getElementById("fieldProxyUrl")).value;
  const proxyToken = /** @type {HTMLInputElement} */ (
    document.getElementById("fieldProxyToken")
  ).value;
  return {
    provider: provider || DEFAULT_SETTINGS.provider,
    modelId: modelId || DEFAULT_SETTINGS.modelId,
    apiKey,
    useProxy,
    proxyUrl,
    proxyToken,
  };
}

/**
 * @param {import("./settings-storage.js").PiWordSettings} settings
 */
function applySettingsToForm(settings) {
  const fieldProvider = /** @type {HTMLSelectElement} */ (document.getElementById("fieldProvider"));
  const fieldModelId = /** @type {HTMLSelectElement} */ (document.getElementById("fieldModelId"));
  /** @type {HTMLInputElement} */ (document.getElementById("fieldApiKey")).value = settings.apiKey;
  /** @type {HTMLInputElement} */ (document.getElementById("fieldUseProxy")).checked =
    settings.useProxy;
  /** @type {HTMLInputElement} */ (document.getElementById("fieldProxyUrl")).value =
    settings.proxyUrl;
  /** @type {HTMLInputElement} */ (document.getElementById("fieldProxyToken")).value =
    settings.proxyToken;
  if (!fieldProvider || !fieldModelId) {
    return;
  }
  fillProviderSelect(fieldProvider, settings.provider);
  fillModelSelect(fieldModelId, fieldProvider.value, settings.modelId);
}

/**
 * @param {HTMLElement} log
 */
function appendUserMessage(log, text) {
  const el = document.createElement("div");
  el.className = "chat-msg chat-msg-user";
  el.textContent = text;
  log.appendChild(el);
  log.scrollTop = log.scrollHeight;
  return el;
}

/**
 * @param {HTMLElement} log
 */
function appendAssistantShell(log) {
  const el = document.createElement("div");
  el.className = "chat-msg chat-msg-assistant";
  log.appendChild(el);
  log.scrollTop = log.scrollHeight;
  return el;
}

/**
 * @param {HTMLElement} log
 */
function appendSystemLine(log, text) {
  const el = document.createElement("div");
  el.className = "chat-msg chat-msg-tool";
  el.textContent = text;
  log.appendChild(el);
  log.scrollTop = log.scrollHeight;
}

function validateSettingsForRun(settings) {
  const useProxy =
    settings.useProxy &&
    settings.proxyUrl.trim().length > 0 &&
    settings.proxyToken.trim().length > 0;
  if (!useProxy && settings.apiKey.trim().length === 0) {
    return "Add an API key, or enable proxy mode with URL and token.";
  }
  if (settings.useProxy && (!settings.proxyUrl.trim() || !settings.proxyToken.trim())) {
    return "Proxy mode needs both proxy URL and proxy token.";
  }
  return null;
}

/** @type {boolean} */
let appStarted = false;

/**
 * @param {{ host?: unknown } | null | undefined} info
 */
function init(info) {
  const mount = document.getElementById("app-root");
  if (!mount) {
    throw new Error("Pi4Word: #app-root is missing from index.html");
  }
  mount.replaceChildren();
  renderApp();

  const inWord = Boolean(
    info &&
      typeof Office !== "undefined" &&
      typeof Office.HostType !== "undefined" &&
      info.host === Office.HostType.Word,
  );
  setStatus(
    inWord
      ? "Connected to Word — configure the model, then chat below."
      : "Open this add-in in Word for document tools; chat still works for testing.",
    false,
  );

  const main = document.getElementById("mainDiv");
  const chatLog = document.getElementById("chatLog");
  const input = /** @type {HTMLTextAreaElement} */ (document.getElementById("chatInput"));
  const sendBtn = /** @type {HTMLButtonElement} */ (document.getElementById("sendBtn"));
  const stopBtn = /** @type {HTMLButtonElement} */ (document.getElementById("stopBtn"));
  const clearBtn = /** @type {HTMLButtonElement} */ (document.getElementById("clearChatBtn"));
  const saveConnBtn = /** @type {HTMLButtonElement} */ (
    document.getElementById("saveConnectionBtn")
  );
  const fieldProvider = /** @type {HTMLSelectElement} */ (document.getElementById("fieldProvider"));
  const fieldModelId = /** @type {HTMLSelectElement} */ (document.getElementById("fieldModelId"));

  if (!main || !chatLog || !input || !sendBtn || !stopBtn || !clearBtn || !fieldProvider || !fieldModelId) {
    throw new Error("Pi4Word: expected task pane controls were not found after renderApp()");
  }

  fieldProvider.addEventListener("change", () => {
    fillModelSelect(fieldModelId, fieldProvider.value, undefined);
  });

  let stored;
  try {
    stored = loadSettings(localStorage);
  } catch (e) {
    console.warn("[pi4word] could not read settings from storage:", e);
    stored = { ...DEFAULT_SETTINGS };
  }
  applySettingsToForm(stored);

  /** @type {import("@mariozechner/pi-agent-core").Agent | null} */
  let agent = null;
  let settingsSig = "";
  let unsubscribe = () => {};

  /** @returns {import("@mariozechner/pi-agent-core").Agent} */
  function ensureAgent() {
    const settings = readSettingsFromForm();
    const sig = serializeSettings(settings);
    if (agent && sig === settingsSig) {
      return agent;
    }
    try {
      agent = createWordAgent(settings);
      settingsSig = sig;
      unsubscribe();
      unsubscribe = attachAgentListeners(agent);
      return agent;
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      setStatus(`Model error: ${msg}`, true);
      agent = null;
      settingsSig = "";
      throw err;
    }
  }

  /** @type {HTMLElement | null} */
  let currentAssistantEl = null;

  function attachAgentListeners(a) {
    return a.subscribe((event) => {
      if (event.type === "message_start" && event.message.role === "assistant") {
        currentAssistantEl = appendAssistantShell(chatLog);
        currentAssistantEl.textContent = "";
      }
      if (event.type === "message_update" && event.assistantMessageEvent?.type === "text_delta") {
        if (currentAssistantEl) {
          currentAssistantEl.textContent += event.assistantMessageEvent.delta;
          chatLog.scrollTop = chatLog.scrollHeight;
        }
      }
      if (event.type === "tool_execution_start") {
        appendSystemLine(chatLog, `Tool: ${event.toolName}`);
      }
      if (event.type === "agent_end") {
        currentAssistantEl = null;
      }
    });
  }

  try {
    ensureAgent();
  } catch {
    // Invalid model until user fixes connection
  }

  saveConnBtn?.addEventListener("click", () => {
    const settings = readSettingsFromForm();
    saveSettings(localStorage, settings);
    try {
      settingsSig = "";
      ensureAgent();
      setStatus("Connection settings saved.", false);
    } catch {
      // status set in ensureAgent
    }
  });

  function setRunning(running) {
    sendBtn.disabled = running;
    stopBtn.disabled = !running;
    input.disabled = running;
  }

  stopBtn.addEventListener("click", () => {
    agent?.abort();
  });

  clearBtn.addEventListener("click", () => {
    agent?.reset();
    while (chatLog.firstChild) {
      chatLog.removeChild(chatLog.firstChild);
    }
    setStatus("Conversation cleared.", false);
  });

  async function onSend() {
    const text = input.value.trim();
    if (!text) {
      return;
    }
    const settings = readSettingsFromForm();
    const err = validateSettingsForRun(settings);
    if (err) {
      setStatus(err, true);
      return;
    }

    try {
      ensureAgent();
    } catch {
      return;
    }

    if (!agent) {
      return;
    }

    input.value = "";
    appendUserMessage(chatLog, text);
    setRunning(true);
    setStatus("Thinking…", false);

    try {
      await agent.prompt(text);
      const errMsg = agent.state.errorMessage;
      if (errMsg) {
        setStatus(errMsg, true);
      } else {
        setStatus("Ready.", false);
      }
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      setStatus(msg, true);
    } finally {
      setRunning(false);
    }
  }

  sendBtn.addEventListener("click", () => {
    void onSend();
  });
  input.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter" && !ev.shiftKey) {
      ev.preventDefault();
      void onSend();
    }
  });
}

/**
 * Ensures the UI mounts even when Office.js is blocked (e.g. tracking prevention) or onReady is late.
 * @param {{ host?: unknown } | null | undefined} info
 */
function startApp(info) {
  if (appStarted) {
    return;
  }
  try {
    init(info);
    appStarted = true;
  } catch (e) {
    console.error("[pi4word] init failed:", e);
  }
}

function scheduleOfficeBoot() {
  const runOfficeReady = () => {
    if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
      Office.onReady((info) => startApp(info));
    } else {
      startApp(null);
    }
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", runOfficeReady, { once: true });
  } else {
    runOfficeReady();
  }

  window.setTimeout(() => {
    if (!appStarted) {
      console.warn(
        "[pi4word] Office.onReady did not finish; showing UI without host context (check Office.js / tracking prevention).",
      );
      startApp(null);
    }
  }, 2500);
}

scheduleOfficeBoot();
