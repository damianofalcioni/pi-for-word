import test from "node:test";
import assert from "node:assert/strict";
import { validateSettingsForRun } from "./task-pane.boundary.js";

const base = {
  provider: "openai",
  modelId: "gpt-4.1",
  apiKey: "",
  useProxy: false,
  proxyUrl: "",
  proxyToken: "",
};

test("validateSettingsForRun accepts API key in direct mode", () => {
  assert.equal(
    validateSettingsForRun({ ...base, apiKey: "sk-test", useProxy: false }),
    null,
  );
});

test("validateSettingsForRun rejects direct mode without key or proxy", () => {
  const msg = validateSettingsForRun({ ...base, apiKey: "", useProxy: false });
  assert.ok(msg && msg.length > 0);
});

test("validateSettingsForRun accepts proxy when URL and token set", () => {
  assert.equal(
    validateSettingsForRun({
      ...base,
      apiKey: "",
      useProxy: true,
      proxyUrl: "https://proxy.example.com",
      proxyToken: "tok",
    }),
    null,
  );
});

test("validateSettingsForRun rejects proxy mode with missing URL or token", () => {
  assert.ok(
    validateSettingsForRun({
      ...base,
      apiKey: "",
      useProxy: true,
      proxyUrl: "",
      proxyToken: "tok",
    }),
  );
  assert.ok(
    validateSettingsForRun({
      ...base,
      apiKey: "",
      useProxy: true,
      proxyUrl: "https://proxy.example.com",
      proxyToken: "",
    }),
  );
});
