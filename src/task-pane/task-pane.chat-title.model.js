/**
 * Chat session title derivation from agent messages (pure logic).
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
export function generateTitle(messages) {
  return shortenChatTitle(textFromUserLikeMessage(firstUserLikeMessage(messages)));
}
