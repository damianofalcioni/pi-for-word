/** Minimal `process` stub for browser bundles; `--inject` installs `globalThis.process`, `--alias:process=` resolves imports here. */
if (typeof globalThis.process === "undefined") {
  globalThis.process = { env: {} };
} else if (!globalThis.process.env) {
  globalThis.process.env = {};
}

export default globalThis.process;
