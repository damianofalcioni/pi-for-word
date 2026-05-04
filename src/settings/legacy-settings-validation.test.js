import test from "node:test";
import assert from "node:assert/strict";
import { validateLegacySettingsForRun } from "./legacy-settings-validation.js";

const base = {
  provider: "openai",
  modelId: "gpt-4.1",
  apiKey: "",
  useProxy: false,
  proxyUrl: "",
  proxyToken: "",
};

test("validateLegacySettingsForRun accepts API key in direct mode", () => {
  assert.equal(
    validateLegacySettingsForRun({ ...base, apiKey: "sk-test", useProxy: false }),
    null,
  );
});

test("validateLegacySettingsForRun rejects direct mode without key or proxy", () => {
  const msg = validateLegacySettingsForRun({ ...base, apiKey: "", useProxy: false });
  assert.ok(msg && msg.length > 0);
});

test("validateLegacySettingsForRun accepts proxy when URL and token set", () => {
  assert.equal(
    validateLegacySettingsForRun({
      ...base,
      apiKey: "",
      useProxy: true,
      proxyUrl: "https://proxy.example.com",
      proxyToken: "tok",
    }),
    null,
  );
});

test("validateLegacySettingsForRun rejects proxy mode with missing URL or token", () => {
  assert.ok(
    validateLegacySettingsForRun({
      ...base,
      apiKey: "",
      useProxy: true,
      proxyUrl: "",
      proxyToken: "tok",
    }),
  );
  assert.ok(
    validateLegacySettingsForRun({
      ...base,
      apiKey: "",
      useProxy: true,
      proxyUrl: "https://proxy.example.com",
      proxyToken: "",
    }),
  );
});
