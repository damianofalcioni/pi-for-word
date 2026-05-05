import {
  ApiKeyPromptDialog,
  ApiKeysTab,
  ChatPanel,
  getAppStorage,
  ModelSelector,
  ProxyTab,
  ProvidersModelsTab,
  SessionListDialog,
  SettingsDialog,
} from "@mariozechner/pi-web-ui";
import { createPi4WordJavaScriptReplTool } from "../assistant/javascript-repl-tool.js";
import {
  createWordAgent,
  createWordAgentFromSession,
  createWordTools,
  getDefaultWordModel,
} from "../assistant/index.js";
import {
  loadPreferredChatModel,
  loadPreferredThinkingLevel,
  persistPreferredChatModel,
  persistPreferredThinkingLevel,
  Pi4WordProxyTab,
} from "../settings/index.js";
import { attachSessionAutosave, sessionRef } from "./task-pane.session.js";
import { setStatus } from "./task-pane.boundary.js";

export function queryTaskPaneControls() {
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
 * @param {HTMLElement} chatMount
 * @param {{ agent: import("@mariozechner/pi-agent-core").Agent, bindChatPanel?: () => Promise<void> }} agentHolder
 */
export async function mountChatPanel(chatMount, agentHolder) {
  const chatPanel = new ChatPanel();
  chatMount.appendChild(chatPanel);
  const bindChatPanel = async () => {
    await chatPanel.setAgent(agentHolder.agent, {
      onApiKeyRequired: (provider) => ApiKeyPromptDialog.prompt(provider),
      toolsFactory: (_agent, _agentInterface, artifactsPanel, runtimeProvidersFactory) => {
        const replTool = createPi4WordJavaScriptReplTool();
        replTool.runtimeProvidersFactory = runtimeProvidersFactory;
        replTool.sandboxUrlProvider = artifactsPanel.sandboxUrlProvider;
        return [...createWordTools(), replTool];
      },
      onBeforeSend: async () => {
        const agent = agentHolder.agent;
        await persistPreferredThinkingLevel(agent.state.thinkingLevel);
      },
      // ModelSelector updates agent.state.model without re-rendering AgentInterface (pi-web-ui).
      onModelSelect: () => {
        const agent = agentHolder.agent;
        void ModelSelector.open(agent.state.model, (model) => {
          agent.state.model = model;
          void persistPreferredChatModel(model);
          const iface = chatPanel.querySelector("agent-interface");
          iface?.requestUpdate();
        });
      },
    });
  };
  agentHolder.bindChatPanel = bindChatPanel;
  await bindChatPanel();
}

/** @param {HTMLButtonElement} settingsBtn */
export function wireSettingsButton(settingsBtn) {
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
export function wireSessionsButton(controls, agentHolder) {
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
export function wireNewSessionButton(controls, agentHolder) {
  controls.newSessionBtn.addEventListener("click", async () => {
    sessionRef.current = undefined;
    sessionRef.title = "";
    const preferred = await loadPreferredChatModel();
    const preferredThinking = await loadPreferredThinkingLevel();
    agentHolder.agent = createWordAgent(preferred ?? getDefaultWordModel(), {
      thinkingLevel: preferredThinking,
    });
    await agentHolder.bindChatPanel();
    attachSessionAutosave(agentHolder.agent);
    setStatus("New chat.", false);
  });
}
