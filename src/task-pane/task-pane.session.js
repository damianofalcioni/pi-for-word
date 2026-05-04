import { getAppStorage } from "@mariozechner/pi-web-ui";
import { generateTitle } from "./task-pane.chat-title.model.js";

/** @type {{ current: string | undefined, title: string }} */
export const sessionRef = { current: undefined, title: "" };

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
 * @param {import("@mariozechner/pi-agent-core").Agent} agent
 */
export async function saveCurrentSession(agent) {
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

function findAgentInterfaceElement() {
  const scope = document.getElementById("chatMount");
  return scope?.querySelector("agent-interface") ?? document.querySelector("agent-interface");
}

/**
 * @param {Element} iface
 * @param {import("@mariozechner/pi-agent-core").Agent} agent
 */
function syncChatComposerDom(iface, agent) {
  if (typeof iface.requestUpdate !== "function") {
    return;
  }
  const editor = iface.querySelector("message-editor");
  if (editor && "isStreaming" in editor) {
    editor.isStreaming = Boolean(agent.state.isStreaming);
  }
  iface.requestUpdate();
  if (editor && typeof editor.requestUpdate === "function") {
    editor.requestUpdate();
  }
}

/**
 * Composer shows Stop (square) while `session.state.isStreaming`; pi-web-ui can paint once before
 * `finishRun()` clears it. Session autosave previously `await`ed IndexedDB inside `agent_end`, which
 * delayed `finishRun()` and made `setTimeout(0)` refresh misleading. Re-sync after the run is idle.
 *
 * @param {import("@mariozechner/pi-agent-core").Agent} agent
 */
function refreshChatComposerAfterIdle(agent) {
  void agent.waitForIdle().then(() => {
    const iface = findAgentInterfaceElement();
    if (!iface) {
      return;
    }
    syncChatComposerDom(iface, agent);
  });
}

/**
 * @param {import("@mariozechner/pi-agent-core").Agent} agent
 */
export function attachSessionAutosave(agent) {
  unsubscribeAgent();
  unsubscribeAgent = agent.subscribe((event) => {
    if (event.type === "agent_end") {
      refreshChatComposerAfterIdle(agent);
      void saveCurrentSession(agent).catch((e) => {
        console.warn("[pi4word] session save failed:", e);
      });
    }
  });
}
