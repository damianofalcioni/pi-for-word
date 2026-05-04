/* Injected first by esbuild (`--inject`); some bundled deps expect `globalThis.process.env`. */
if (typeof globalThis.process === "undefined") {
  globalThis.process = { env: {} };
} else if (!globalThis.process.env) {
  globalThis.process.env = {};
}
