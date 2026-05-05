import test from "node:test";
import assert from "node:assert/strict";
import {
  parseSettings,
  serializeSettings,
  DEFAULT_SETTINGS,
} from "./settings-storage.js";

test("parseSettings returns defaults for null input", () => {
  const s = parseSettings(null);
  assert.deepEqual(s, DEFAULT_SETTINGS);
});

test("parseSettings merges known fields", () => {
  const raw = JSON.stringify({
    provider: "openai",
    modelId: "gpt-4.1",
    apiKey: "secret",
    useProxy: true,
    proxyUrl: "https://x.example.com",
    proxyToken: "tok",
  });
  const s = parseSettings(raw);
  assert.equal(s.provider, "openai");
  assert.equal(s.modelId, "gpt-4.1");
  assert.equal(s.apiKey, "secret");
  assert.equal(s.useProxy, true);
  assert.equal(s.proxyUrl, "https://x.example.com");
  assert.equal(s.proxyToken, "tok");
});

test("serializeSettings roundtrips parseSettings", () => {
  const a = {
    provider: "openai",
    modelId: "gpt-5.5",
    apiKey: "k",
    useProxy: false,
    proxyUrl: "",
    proxyToken: "",
  };
  const b = parseSettings(serializeSettings(a));
  assert.deepEqual(b, a);
});
