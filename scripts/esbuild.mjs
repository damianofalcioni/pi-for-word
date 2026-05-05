import * as esbuild from "esbuild";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const katexFontsDir = path.join(repoRoot, "node_modules", "katex", "dist", "fonts");

/** @returns {import("esbuild").Plugin} */
function resolveKatexFontsInPiWebUiCss() {
  return {
    name: "resolve-katex-fonts-pi-web-ui",
    setup(build) {
      build.onResolve({ filter: /^fonts\/KaTeX_/ }, (args) => {
        if (args.kind !== "url-token") {
          return undefined;
        }
        const basename = args.path.startsWith("fonts/")
          ? args.path.slice("fonts/".length)
          : args.path;
        return { path: path.join(katexFontsDir, basename) };
      });
    },
  };
}

await esbuild.build({
  entryPoints: [path.join(repoRoot, "src", "index.js")],
  bundle: true,
  minify: true,
  sourcemap: true,
  platform: "browser",
  format: "esm",
  outfile: path.join(repoRoot, "public", "index.min.js"),
  external: ["office-js"],
  alias: {
    process: path.join(repoRoot, "src", "shims", "process.js"),
  },
  inject: [path.join(repoRoot, "src", "shims", "process.js")],
  assetNames: "assets/[name]-[hash]",
  loader: {
    ".ttf": "file",
    ".woff": "file",
    ".woff2": "file",
  },
  plugins: [resolveKatexFontsInPiWebUiCss()],
});

// pi-web-ui resolves the PDF.js worker as ./pdfjs-dist/build/pdf.worker.min.mjs
// relative to the bundled script (public/index.min.js). Copy the worker from npm.
const pdfWorkerSrc = path.join(
  repoRoot,
  "node_modules",
  "pdfjs-dist",
  "build",
  "pdf.worker.min.mjs",
);
const pdfWorkerDestDir = path.join(repoRoot, "public", "pdfjs-dist", "build");
const pdfWorkerDest = path.join(pdfWorkerDestDir, "pdf.worker.min.mjs");
fs.mkdirSync(pdfWorkerDestDir, { recursive: true });
fs.copyFileSync(pdfWorkerSrc, pdfWorkerDest);
