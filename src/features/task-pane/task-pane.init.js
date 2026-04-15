import { createWordAgent } from "../assistant/pi-assistant.js";
import {
  DEFAULT_SETTINGS,
  loadSettings,
  saveSettings,
  serializeSettings,
} from "../settings/settings-storage.js";
import {
  appendAssistantShell,
  appendSystemLine,
  appendUserMessage,
  applySettingsToForm,
  fillModelSelect,
  readSettingsFromForm,
  renderApp,
  setStatus,
  validateSettingsForRun,
} from "./task-pane.boundary.js";

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
    chatLog.replaceChildren();
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

/** Mounts the task pane, loads settings, wires the agent and Office.js bootstrap. */
export function initializeTaskPane() {
  scheduleOfficeBoot();
}
