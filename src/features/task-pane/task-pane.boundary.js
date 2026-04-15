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
  status.style.background = isError ? "#f7d8d2" : "#edf4f4";
  status.style.color = isError ? "#6a1b16" : "#134243";
}

/**
 * Builds the task-pane shell: header, toolbar, and mount point for pi-web-ui ChatPanel.
 * Stylesheets: `index.css`, `pi-web-ui-app.css` (linked from index.html).
 */
export function renderApp() {
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

  const main = document.createElement("main");
  main.id = "mainDiv";
  main.className = "app-main";

  const chatMount = document.createElement("div");
  chatMount.id = "chatMount";
  chatMount.className = "chat-mount";
  chatMount.setAttribute("aria-label", "Chat");

  main.append(toolbar, chatMount);

  const mount = document.getElementById("app-root");
  if (!mount) {
    throw new Error("Pi4Word: #app-root is missing from index.html");
  }
  mount.append(header, main);
}
