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

function createAppToolbar() {
  const toolbar = document.createElement("div");
  toolbar.className = "app-toolbar";
  const settingsBtn = document.createElement("button");
  settingsBtn.type = "button";
  settingsBtn.id = "settingsBtn";
  settingsBtn.className = "primary";
  settingsBtn.textContent = "Settings";
  const sessionsBtn = document.createElement("button");
  sessionsBtn.type = "button";
  sessionsBtn.id = "sessionsBtn";
  sessionsBtn.textContent = "Sessions";
  const newSessionBtn = document.createElement("button");
  newSessionBtn.type = "button";
  newSessionBtn.id = "newSessionBtn";
  newSessionBtn.textContent = "New chat";
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
