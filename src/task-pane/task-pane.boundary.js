/**
 * @param {string} message
 * @param {boolean} isError
 */
export function setStatus(message, isError) {
  const status = document.getElementById("hostStatus");
  if (!status) {
    return;
  }
  status.textContent = message;
  status.classList.toggle("is-error", Boolean(isError));
}

function createAppHeader() {
  const header = document.createElement("header");
  header.className = "app-header";
  const headerLeft = document.createElement("div");
  const title = document.createElement("div");
  title.className = "app-title";
  title.textContent = "Pi4Word";
  const subtitle = document.createElement("div");
  subtitle.className = "app-subtitle";
  subtitle.textContent = "Pi Agent assistant for Word";
  headerLeft.append(title, subtitle);

  const hostStatus = document.createElement("div");
  hostStatus.className = "status";
  hostStatus.id = "hostStatus";
  hostStatus.textContent = "Loading…";

  header.append(headerLeft, hostStatus);
  return header;
}

/**
 * @param {HTMLButtonElement} button
 * @param {string} label Accessible name and tooltip (e.g. "Settings").
 * @param {string} svgChildren Inner markup for SVG paths/lines (viewBox 0 0 24 24).
 */
function setToolbarIconButton(button, label, svgChildren) {
  button.textContent = "";
  button.setAttribute("aria-label", label);
  button.title = label;
  button.classList.add("toolbar-icon-btn");
  button.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${svgChildren}</svg>`;
}

function createAppToolbar() {
  const toolbar = document.createElement("div");
  toolbar.className = "app-toolbar";
  const settingsBtn = document.createElement("button");
  settingsBtn.type = "button";
  settingsBtn.id = "settingsBtn";
  setToolbarIconButton(
    settingsBtn,
    "Settings",
    '<path d="M9.594 3.94c.09-.542.56-.94 1.11-.94h2.593c.55 0 1.02.398 1.11.94l.213 1.281c.063.374.313.686.645.87.074.04.147.083.22.127.324.196.72.257 1.075.124l1.217-.456a1.125 1.125 0 0 1 1.37.49l1.296 2.247a1.125 1.125 0 0 1-.26 1.431l-1.003.827c-.293.24-.438.613-.431.992a6.759 6.759 0 0 1 0 .255c-.008.378.137.75.431.99l1.005.828c.424.35.534.954.26 1.43l-1.298 2.247a1.125 1.125 0 0 1-1.369.491l-1.217-.456c-.355-.133-.75-.072-1.076.124a6.57 6.57 0 0 1-.22.128c-.331.183-.581.495-.644.869l-.213 1.28c-.09.543-.56.941-1.11.941h-2.594c-.55 0-1.02-.398-1.11-.94l-.213-1.281c-.062-.374-.312-.686-.644-.87a6.52 6.52 0 0 1-.22-.127c-.325-.196-.72-.257-1.076-.124l-1.217.456a1.125 1.125 0 0 1-1.369-.49l-1.297-2.247a1.125 1.125 0 0 1 .26-1.431l1.004-.827c.292-.24.437-.613.43-.992a6.932 6.932 0 0 1 0-.255c.007-.378-.138-.75-.43-.99l-1.004-.828a1.125 1.125 0 0 1-.26-1.43l1.297-2.247a1.125 1.125 0 0 1 1.37-.491l1.216.456c.356.133.751.072 1.076-.124.072-.044.146-.087.22-.128.332-.183.582-.495.644-.869l.214-1.281z"/><path d="M15 12a3 3 0 1 1-6 0 3 3 0 0 1 6 0z"/>',
  );
  const sessionsBtn = document.createElement("button");
  sessionsBtn.type = "button";
  sessionsBtn.id = "sessionsBtn";
  setToolbarIconButton(
    sessionsBtn,
    "Sessions",
    '<path d="M8 6h13"/><path d="M8 12h13"/><path d="M8 18h13"/><path d="M3 6h.01"/><path d="M3 12h.01"/><path d="M3 18h.01"/>',
  );
  const newSessionBtn = document.createElement("button");
  newSessionBtn.type = "button";
  newSessionBtn.id = "newSessionBtn";
  setToolbarIconButton(
    newSessionBtn,
    "New chat",
    '<circle cx="12" cy="12" r="10"/><path d="M12 8v8"/><path d="M8 12h8"/>',
  );
  toolbar.append(settingsBtn, sessionsBtn, newSessionBtn);
  return toolbar;
}

function createAppMain() {
  const main = document.createElement("main");
  main.id = "mainDiv";
  main.className = "app-main";

  const chatMount = document.createElement("div");
  chatMount.id = "chatMount";
  chatMount.className = "chat-mount";
  chatMount.setAttribute("aria-label", "Chat");

  main.append(createAppToolbar(), chatMount);
  return main;
}

/**
 * Builds the task-pane shell: header, toolbar, and mount point for pi-web-ui ChatPanel.
 * Styles: `src/index.css` + `@mariozechner/pi-web-ui/app.css` (imported from `src/index.js`, esbuild → `public/index.min.css`).
 */
export function renderApp() {
  const mount = document.getElementById("app-root");
  if (!mount) {
    throw new Error("Pi4Word: #app-root is missing from index.html");
  }
  mount.append(createAppHeader(), createAppMain());
}
