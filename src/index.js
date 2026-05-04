import "@mariozechner/pi-web-ui/app.css";
import "./index.css";

import { initializeTaskPane } from "./task-pane/task-pane.init.js";

function showBootError(msg) {
  const root = document.getElementById("app-root");
  if (!root) {
    return;
  }
  root.replaceChildren();
  const wrap = document.createElement("div");
  wrap.className = "boot-error";
  wrap.textContent = String(msg);
  root.appendChild(wrap);
}

window.addEventListener("error", (ev) => {
  const m =
    ev.error?.message ??
    ev.message ??
    "Script error";
  showBootError(`Error: ${m}`);
});
window.addEventListener("unhandledrejection", (ev) => {
  const r = ev.reason;
  const m = r?.message ?? String(r ?? "unknown");
  showBootError(`Error (promise): ${m}`);
});

initializeTaskPane();
