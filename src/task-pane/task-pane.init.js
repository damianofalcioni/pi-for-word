import { getAppStorage } from "@mariozechner/pi-web-ui";
import { createWordAgent, getDefaultWordModel } from "../assistant/index.js";
import { initPiWebStorage, migrateLegacyLocalStorageOnce } from "../settings/index.js";
import { attachSessionAutosave } from "./task-pane.session.js";
import { computeWordHost, scheduleOfficeBoot } from "./task-pane.office.js";
import {
  mountChatPanel,
  queryTaskPaneControls,
  wireNewSessionButton,
  wireSessionsButton,
  wireSettingsButton,
} from "./task-pane.wiring.js";
import { renderApp, setStatus } from "./task-pane.boundary.js";

/** @type {boolean} */
let appStarted = false;

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

/** Mounts the task pane, loads settings, wires the agent and Office.js bootstrap. */
export function initializeTaskPane() {
  scheduleOfficeBoot(startApp, { getStarted: () => appStarted });
}
