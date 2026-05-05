# Pi4Word — specification

## Product summary

**Pi4Word** is a Word task pane add-in that runs an AI assistant powered by [`@mariozechner/pi-agent-core`](https://github.com/badlogic/pi-mono/tree/main/packages/agent) (`pi-agent-core` in pi-mono). The assistant reads the selection and inserts text into the document through Word JavaScript API tools. The UI is **`@mariozechner/pi-web-ui`**’s **`ChatPanel`** (mounted in **`#chatMount`**), composed from packages under **`src/`** (`settings`, `assistant`, `task-pane`) per **`AGENT.md`**.

## Core product constraints

- **Hosting:** Runs as an Office Word task pane; bootstrap uses **`Office.onReady`** with a bounded DOM-ready fallback timing.
- **Storage:** Conversation and settings persistence use **IndexedDB** via pi-web-ui **`AppStorage`**; a **one-time migration** may move legacy **`localStorage`** shapes into provider keys and Pi4Word **`streamProxy`** settings.
- **Networking:** LLM traffic uses a **composite stream function** — when Pi4Word proxy settings are enabled, requests go through configured **`streamProxy`** (URL + Bearer token); otherwise streaming uses pi-web-ui **CORS** **`createStreamFn`**. The repository does **not** ship a backend; deployments supply their own **`streamProxy`** endpoint for secure key handling in production.
- **Build/runtime:** Entry **`src/index.js`** imports **`src/index.css`** and **`@mariozechner/pi-web-ui/app.css`**; **esbuild** bundles JS to **`public/index.min.js`** (with **`process`** inject/alias, **`office-js`** external) and CSS to **`public/index.min.css`**, plus font assets under **`public/assets/`**. **`public/index.html`** loads **`index.min.css`** and **`index.min.js`**.
- **Modularity:** Feature code lives under **`src/<feature>/`** with a single composition root in **`src/task-pane/`**; **`assistant`** and **`settings`** are consumed only via their **`index.js`** barrels. The app bundle entry **`src/index.js`** imports **`task-pane`** only (through **`src/task-pane/index.js`**).

## Functional specification

- **Task pane startup:** **`initializeTaskPane()`** schedules **`Office.onReady`** (or starts without host context if Office.js is missing) and a **2.5s** fallback if the app has not started. **`bootstrapTaskPane`** runs **`initPiWebStorage()`** (IndexedDB stores and **`setAppStorage`**), **`migrateLegacyLocalStorageOnce`** when applicable, **`renderApp`**, **`mountChatPanel`** + **`setAgent`** (Word tools, API key prompt, model selector), **`attachSessionAutosave`**, and toolbar handlers for **Settings**, **Sessions**, and **New chat** (pi-web-ui dialogs / session list).
- **Settings:** Users configure model, API keys, CORS proxy, and Pi **`streamProxy`** (including a **Pi4Word proxy** tab for URL and token).
- **Assistant:** Pi agent wiring, prompts, and **`defaultConvertToLlm`** integrate **Word tools** **`word_get_selection`** and **`word_insert_markdown`** with the chat stream.
- **Sessions:** Session titles and autosave behavior for titles and persistence are handled in the task-pane layer (preview/save helpers).
- **Static shell:** **`public/index.html`** provides the mount point; **`manifest.xml`** defines the add-in manifest for Word.
