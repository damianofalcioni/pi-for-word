import {
  ApiKeyPromptDialog,
  ApiKeysTab,
  ChatPanel,
  getAppStorage,
  ProxyTab,
  ProvidersModelsTab,
  SessionListDialog,
  SettingsDialog,
} from "@mariozechner/pi-web-ui";
import { Pi4WordProxyTab } from "../settings/pi4word-proxy-tab.js";
import {
  createWordAgent,
  createWordAgentFromSession,
  createWordTools,
  getDefaultWordModel,
} from "../assistant/pi-assistant.js";
import { initPiWebStorage, migrateLegacyLocalStorageOnce } from "./pi-web-storage.js";
import { renderApp, setStatus } from "./task-pane.boundary.js";

/** @type {boolean} */
let appStarted = false;

/** @type {{ current: string | undefined, title: string }} */
const sessionRef = { current: undefined, title: "" };

/** @type {(() => void) | undefined} */
let unsubscribeAgent = () => {};

/**
 * @param {import("@mariozechner/pi-agent-core").AgentMessage[]} messages
 */
function shouldSaveSession(messages) {
  const hasUser = messages.some((m) => m.role === "user" || m.role === "user-with-attachments");
  const hasAssistant = messages.some((m) => m.role === "assistant");
  return hasUser && hasAssistant;
}

/**
 * @param {import("@mariozechner/pi-agent-core").AgentMessage[]} messages
 */
function firstUserLikeMessage(messages) {
  return messages.find((m) => m.role === "user" || m.role === "user-with-attachments");
}

/**
 * @param {import("@mariozechner/pi-agent-core").AgentMessage | undefined} first
 */
function textFromUserLikeMessage(first) {
  if (!first) {
    return "";
  }
  if (first.role === "user") {
    const c = first.content;
    if (typeof c === "string") {
      return c;
    }
    if (Array.isArray(c)) {
      return c
        .filter((x) => x && x.type === "text")
        .map((x) => x.text || "")
        .join(" ");
    }
    return "";
  }
  return String(first.content ?? "");
}

/** @param {string} text */
function shortenChatTitle(text) {
  const trimmed = text.trim();
  if (!trimmed) {
    return "";
  }
  const end = trimmed.search(/[.!?]/);
  if (end > 0 && end <= 50) {
    return trimmed.substring(0, end + 1);
  }
  return trimmed.length <= 50 ? trimmed : `${trimmed.substring(0, 47)}…`;
}

/**
 * @param {import("@mariozechner/pi-agent-core").AgentMessage[]} messages
 */
function generateTitle(messages) {
  return shortenChatTitle(textFromUserLikeMessage(firstUserLikeMessage(messages)));
}

/**
 * @param {import("@mariozechner/pi-agent-core").Agent} agent
 */
async function saveCurrentSession(agent) {
  const storage = getAppStorage();
  if (!storage.sessions) {
    return;
  }
  const state = agent.state;
  if (!shouldSaveSession(state.messages)) {
    return;
  }

  if (!sessionRef.current) {
    sessionRef.current = crypto.randomUUID();
  }
  if (!sessionRef.title) {
    sessionRef.title = generateTitle(state.messages);
  }

  const now = new Date().toISOString();
  const sessionData = {
    id: sessionRef.current,
    title: sessionRef.title,
    model: state.model,
    thinkingLevel: state.thinkingLevel,
    messages: state.messages,
    createdAt: now,
    lastModified: now,
  };
  const metadata = {
    id: sessionRef.current,
    title: sessionRef.title,
    createdAt: now,
    lastModified: now,
    messageCount: state.messages.length,
    usage: {
      input: 0,
      output: 0,
      cacheRead: 0,
      cacheWrite: 0,
      totalTokens: 0,
      cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
    },
    modelId: state.model?.id ?? null,
    thinkingLevel: state.thinkingLevel,
    preview: generateTitle(state.messages),
  };
  await storage.sessions.save(sessionData, metadata);
}

/**
 * @param {import("@mariozechner/pi-agent-core").Agent} agent
 */
function attachSessionAutosave(agent) {
  unsubscribeAgent();
  unsubscribeAgent = agent.subscribe(async (event) => {
    if (event.type === "agent_end") {
      try {
        await saveCurrentSession(agent);
      } catch (e) {
        console.warn("[pi4word] session save failed:", e);
      }
    }
  });
}

/** @param {{ host?: unknown } | null | undefined} info */
function computeWordHost(info) {
  return Boolean(
    info &&
      typeof Office !== "undefined" &&
      typeof Office.HostType !== "undefined" &&
      info.host === Office.HostType.Word,
  );
}

function queryTaskPaneControls() {
  const chatMount = document.getElementById("chatMount");
  const settingsBtn = /** @type {HTMLButtonElement | null} */ (document.getElementById("settingsBtn"));
  const sessionsBtn = /** @type {HTMLButtonElement | null} */ (document.getElementById("sessionsBtn"));
  const newSessionBtn = /** @type {HTMLButtonElement | null} */ (document.getElementById("newSessionBtn"));

  if (!chatMount || !settingsBtn || !sessionsBtn || !newSessionBtn) {
    throw new Error("Pi4Word: expected task pane controls were not found after renderApp()");
  }
  return { chatMount, settingsBtn, sessionsBtn, newSessionBtn };
}

/**
 * @param {{ host?: unknown } | null | undefined} info
 * @returns {Promise<{
 *   controls: ReturnType<typeof queryTaskPaneControls>,
 *   agentHolder: {
 *     agent: import("@mariozechner/pi-agent-core").Agent,
 *     bindChatPanel?: () => Promise<void>,
 *   },
 * }>}
 */
async function bootstrapTaskPane(info) {
  const mount = document.getElementById("app-root");
  if (!mount) {
    throw new Error("Pi4Word: #app-root is missing from index.html");
  }
  mount.replaceChildren();

  await initPiWebStorage();
  const migratedModel = await migrateLegacyLocalStorageOnce(getAppStorage());

  renderApp();

  const inWord = computeWordHost(info);
  setStatus(
    inWord
      ? "Connected to Word — use Settings for API keys and model."
      : "Open in Word for document tools; chat works in the browser for testing.",
    false,
  );

  const controls = queryTaskPaneControls();
  const agentHolder = {
    agent: createWordAgent(migratedModel ?? getDefaultWordModel()),
  };

  return { controls, agentHolder };
}

/**
 * @param {HTMLElement} chatMount
 * @param {{ agent: import("@mariozechner/pi-agent-core").Agent, bindChatPanel?: () => Promise<void> }} agentHolder
 */
async function mountChatPanel(chatMount, agentHolder) {
  const chatPanel = new ChatPanel();
  chatMount.appendChild(chatPanel);
  const bindChatPanel = async () => {
    await chatPanel.setAgent(agentHolder.agent, {
      onApiKeyRequired: (provider) => ApiKeyPromptDialog.prompt(provider),
      toolsFactory: () => createWordTools(),
    });
  };
  agentHolder.bindChatPanel = bindChatPanel;
  await bindChatPanel();
}

/** @param {HTMLButtonElement} settingsBtn */
function wireSettingsButton(settingsBtn) {
  settingsBtn.addEventListener("click", () => {
    void SettingsDialog.open([
      new ProvidersModelsTab(),
      new ProxyTab(),
      new ApiKeysTab(),
      new Pi4WordProxyTab(),
    ]);
  });
}

/**
 * @param {ReturnType<typeof queryTaskPaneControls>} controls
 * @param {{ agent: import("@mariozechner/pi-agent-core").Agent, bindChatPanel?: () => Promise<void> }} agentHolder
 */
function wireSessionsButton(controls, agentHolder) {
  controls.sessionsBtn.addEventListener("click", () => {
    void SessionListDialog.open(
      async (sessionId) => {
        const storage = getAppStorage();
        const data = await storage.sessions.get(sessionId);
        if (!data) {
          setStatus("Session not found.", true);
          return;
        }
        sessionRef.current = sessionId;
        sessionRef.title = data.title || "";
        agentHolder.agent = createWordAgentFromSession({
          model: data.model,
          thinkingLevel: data.thinkingLevel,
          messages: data.messages,
        });
        await agentHolder.bindChatPanel();
        attachSessionAutosave(agentHolder.agent);
        setStatus("Session loaded.", false);
      },
      (deletedId) => {
        if (deletedId === sessionRef.current) {
          sessionRef.current = undefined;
          sessionRef.title = "";
        }
      },
    );
  });
}

/**
 * @param {ReturnType<typeof queryTaskPaneControls>} controls
 * @param {{ agent: import("@mariozechner/pi-agent-core").Agent, bindChatPanel?: () => Promise<void> }} agentHolder
 */
function wireNewSessionButton(controls, agentHolder) {
  controls.newSessionBtn.addEventListener("click", async () => {
    sessionRef.current = undefined;
    sessionRef.title = "";
    agentHolder.agent = createWordAgent(getDefaultWordModel());
    await agentHolder.bindChatPanel();
    attachSessionAutosave(agentHolder.agent);
    setStatus("New chat.", false);
  });
}

/**
 * @param {{ host?: unknown } | null | undefined} info
 */
async function init(info) {
  const { controls, agentHolder } = await bootstrapTaskPane(info);
  await mountChatPanel(controls.chatMount, agentHolder);
  attachSessionAutosave(agentHolder.agent);
  wireSettingsButton(controls.settingsBtn);
  wireSessionsButton(controls, agentHolder);
  wireNewSessionButton(controls, agentHolder);
}

/**
 * Ensures the UI mounts even when Office.js is blocked or onReady is late.
 * @param {{ host?: unknown } | null | undefined} info
 */
function startApp(info) {
  if (appStarted) {
    return;
  }
  void init(info)
    .then(() => {
      appStarted = true;
    })
    .catch((e) => {
      console.error("[pi4word] init failed:", e);
      setStatus(e instanceof Error ? e.message : String(e), true);
    });
}

function scheduleOfficeBoot() {
  const runOfficeReady = () => {
    if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
      Office.onReady((officeInfo) => startApp(officeInfo));
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
