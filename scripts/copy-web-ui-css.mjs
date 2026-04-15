import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(__dirname, "..");
const src = path.join(root, "node_modules", "@mariozechner", "pi-web-ui", "dist", "app.css");
const dest = path.join(root, "public", "pi-web-ui-app.css");

if (!fs.existsSync(src)) {
  console.error("Missing:", src);
  process.exit(1);
}
fs.copyFileSync(src, dest);
console.log("Copied pi-web-ui app.css -> public/pi-web-ui-app.css");
