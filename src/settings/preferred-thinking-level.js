import { getAppStorage } from "@mariozechner/pi-web-ui";

/** @typedef {import("@mariozechner/pi-agent-core").ThinkingLevel} ThinkingLevel */

export const PREFERRED_THINKING_LEVEL_KEY = "pi4word.chat.preferredThinkingLevel";

const ALLOWED_LEVELS =
  /** @type {readonly ThinkingLevel[]} */ (["off", "minimal", "low", "medium", "high", "xhigh"]);

const ALLOWED = new Set(ALLOWED_LEVELS);

/**
 * @param {unknown} value
 * @returns {value is ThinkingLevel}
 */
export function isThinkingLevel(value) {
  return typeof value === "string" && ALLOWED.has(/** @type {string} */ (value));
}

/**
 * Synchronous stash so reload/refresh still sees the last choice even when an IndexedDB
 * write does not finish before the task pane unloads (common in embedded hosts).
 * @param {ThinkingLevel} level
 */
export function stashPreferredThinkingLevelLocal(level) {
  try {
    localStorage.setItem(PREFERRED_THINKING_LEVEL_KEY, level);
  } catch {
    /* private mode / disabled storage — IndexedDB path may still work */
  }
}

/**
 * @returns {ThinkingLevel | undefined}
 */
export function peekPreferredThinkingLevelLocal() {
  try {
    const ls = localStorage.getItem(PREFERRED_THINKING_LEVEL_KEY);
    return isThinkingLevel(ls) ? ls : undefined;
  } catch {
    return undefined;
  }
}

/**
 * @returns {Promise<ThinkingLevel | undefined>}
 */
export async function loadPreferredThinkingLevel() {
  const cached = peekPreferredThinkingLevelLocal();
  if (cached !== undefined) {
    return cached;
  }

  let raw;
  try {
    raw = await getAppStorage().settings.get(PREFERRED_THINKING_LEVEL_KEY);
  } catch (e) {
    console.warn("[pi4word] IndexedDB read for thinking level failed:", e);
    return undefined;
  }

  if (!isThinkingLevel(raw)) {
    return undefined;
  }

  stashPreferredThinkingLevelLocal(raw);
  return raw;
}

/**
 * @param {ThinkingLevel} v
 */
async function persistThinkingLevelIndexedDB(v) {
  try {
    await getAppStorage().settings.set(PREFERRED_THINKING_LEVEL_KEY, v);
  } catch (e) {
    console.warn("[pi4word] IndexedDB persist for thinking level failed:", e);
  }
}

/**
 * Mirror to {@link stashPreferredThinkingLevelLocal} immediately, then IndexedDB when available.
 * @param {ThinkingLevel | string} level
 */
export async function persistPreferredThinkingLevel(level) {
  const v = isThinkingLevel(level) ? level : "off";
  stashPreferredThinkingLevelLocal(v);
  await persistThinkingLevelIndexedDB(v);
}

/**
 * Persist thinking level whenever it changes (composer assigns `agent.state.thinkingLevel`).
 * @param {import("@mariozechner/pi-agent-core").Agent} agent
 */
export function attachPreferredThinkingLevelPersistence(agent) {
  const state = agent.state;
  let current = isThinkingLevel(state.thinkingLevel) ? state.thinkingLevel : "off";
  delete state.thinkingLevel;
  Object.defineProperty(state, "thinkingLevel", {
    get() {
      return current;
    },
    set(next) {
      const level = isThinkingLevel(next) ? next : "off";
      if (level === current) {
        return;
      }
      current = level;
      // localStorage is synchronous; IndexedDB may not commit before the task pane unloads.
      stashPreferredThinkingLevelLocal(level);
      void persistThinkingLevelIndexedDB(level);
    },
    enumerable: true,
    configurable: true,
  });
}

/**
 * Last-chance sync for embedded browsers that tear down before IndexedDB transactions settle.
 * @param {{ agent: import("@mariozechner/pi-agent-core").Agent }} agentHolder
 */
export function attachPreferredThinkingLevelUnloadSync(agentHolder) {
  const flush = () => {
    try {
      const tl = agentHolder.agent?.state.thinkingLevel;
      if (isThinkingLevel(tl)) {
        stashPreferredThinkingLevelLocal(tl);
      }
    } catch {
      /* ignore */
    }
  };
  window.addEventListener("pagehide", flush);
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") {
      flush();
    }
  });
}
