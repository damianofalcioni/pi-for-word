import { test } from "node:test";
import assert from "node:assert/strict";
import { generateTitle } from "./task-pane.chat-title.model.js";

test("generateTitle prefers first user message text (string content)", () => {
  const title = generateTitle([
    { role: "user", content: "Draft the intro. With care." },
    { role: "assistant", content: "Done." },
  ]);
  assert.equal(title, "Draft the intro.");
});

test("generateTitle shortens long text without sentence end", () => {
  const long = "a".repeat(60);
  const title = generateTitle([{ role: "user", content: long }, { role: "assistant", content: "ok" }]);
  // First 47 chars + Unicode ellipsis (see shortenChatTitle)
  assert.equal(title.length, 48);
  assert.ok(title.endsWith("…"));
});
