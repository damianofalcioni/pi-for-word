# Pi4Word Specification

## Product Contract

Pi4Word is a Microsoft Word task pane add-in that embeds a Pi Agent assistant powered by [`@mariozechner/pi-agent-core`](https://github.com/badlogic/pi-mono/tree/main/packages/agent).

The assistant can read the current Word selection as GitHub-flavored Markdown, reason over the document through Word tools, and insert or replace content by rendering Markdown into sanitized HTML for the Word JavaScript API.

Core user-visible surfaces:

- **Chat UI:** `@mariozechner/pi-web-ui` `ChatPanel`, mounted into `#chatMount`.
- **Word reading:** `Range.getHtml()` -> Turndown + `turndown-plugin-gfm`, with text fallback.
- **Word writing:** Markdown -> `marked` -> DOMPurify -> Word `insertHtml`.
- **Settings:** model, provider API keys, CORS proxy, and Pi stream proxy.
- **Sessions:** pi-web-ui session persistence with task-pane autosave/title helpers.

## Runtime Model

- **Entry:** `src/index.js` imports `./shims/office-alert.js` first, then pi-web-ui CSS, `src/index.css`, and `initializeTaskPane`.
- **Boot:** `initializeTaskPane()` calls `scheduleOfficeBoot()`, which waits for `DOMContentLoaded` if needed, then `Office.onReady`.
- **Fallback:** if `Office.onReady` is unavailable or does not complete within 2.5s, the UI still mounts with missing host context.
- **Shell:** `renderApp()` clears `#app-root`, builds the header, toolbar, and `#chatMount`.
- **Errors:** `src/index.js` installs `window` `error` and `unhandledrejection` handlers that surface boot failures in `#app-root`.
- **Static HTML:** `public/index.html` provides `#app-root`, loads Office.js from `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`, then loads `./index.min.js` as an ES module and `./index.min.css` in `head`.
- **Manifest:** `manifest.xml` targets Word and points `SourceLocation` to `https://localhost:3000/public/index.html`.

## Assistant Contract

- **Prompt:** `SYSTEM_PROMPT` lives in `src/assistant/pi-assistant.js`.
- **Agent:** `createWordAgent()` and `createWordAgentFromSession()` create pi-agent-core `Agent` instances using `defaultConvertToLlm`.
- **Model default:** `getDefaultWordModel()` uses `DEFAULT_SETTINGS` from settings.
- **API keys:** provider keys come from `getAppStorage().providerKeys`.
- **Chat wiring:** `mountChatPanel()` creates `new ChatPanel()`, appends it to `#chatMount`, and calls `chatPanel.setAgent()`.
- **Tool order:** `toolsFactory` returns `...createWordTools()` first, then `createPi4WordJavaScriptReplTool()`.
- **Tool rendering:** `registerCollapsibleWordToolRenderers()` registers collapsible renderers for the Word tool names.
- **Model picker:** `ModelSelector` updates `agent.state.model` and persists the preferred model.
- **API-key prompt:** `ApiKeyPromptDialog.prompt(provider)` handles missing provider keys.

Tool descriptions and parameter schemas in source are authoritative.

## Word Tools

`createWordTools()` lives in `src/assistant-tools/index.js` and returns these Word tools in order.

| Tool | Source | Contract |
| --- | --- | --- |
| `word_get_selection` | `word-tool-get-selection.js` | Runs `Word.run`, loads selection text, queues `range.getHtml()`, syncs, converts Word HTML to GFM Markdown, and falls back to `range.text` if Markdown is empty but text is present. |
| `word_get_document_outline` | `word-tool-get-document-outline.js` | Returns a table-of-contents-style outline from document body paragraphs using built-in Heading 1-9 styles and/or outline levels 1-9. |
| `word_search_text` | `word-tool-search-text.js` | Searches the document body with Word `search`, accepts optional `search_options`, and returns match count plus a preview for the selected `match_index`. |
| `word_insert_markdown` | `word-tool-insert-markdown.js` | Converts Markdown to sanitized HTML and inserts it using Word `insertHtml` at the requested placement or search anchor. |

### `word_insert_markdown`

Required parameters:

- `markdown`: Markdown to render and insert. Plain prose counts as Markdown.
- `where`: `after_selection`, `before_selection`, `replace_selection`, or `end_of_document`.

Accepted `where` aliases:

- `after` -> `after_selection`
- `before` -> `before_selection`
- `replace` -> `replace_selection`
- `end` / `document_end` -> `end_of_document`

Optional search anchoring parameters:

- `anchor_search_text`
- `anchor_match_index`
- `anchor_search_options`

When `anchor_search_text` is non-empty and `where` is not `end_of_document`, insertion targets the Nth body search hit instead of the current selection.

### Search Options

`word_search_text.search_options` and `word_insert_markdown.anchor_search_options` accept Word `SearchOptions` flags in snake_case:

- `ignore_punct`
- `ignore_space`
- `match_case`
- `match_prefix`
- `match_suffix`
- `match_whole_word`
- `match_wildcards`

When an insert is anchored from a prior search, `anchor_search_options` must mirror the original `search_options`; mismatched flags can target a different occurrence. Word wildcard and search behavior follows Microsoft's [search-option guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/word/search-option-guidance).

### Markdown And HTML

- Markdown insertion uses `marked` and DOMPurify.
- GFM task-list `<input type="checkbox">` is allowlisted for insertion.
- Markdown readback from Word is best-effort because Word HTML is approximate, not an OOXML round trip.

## JavaScript REPL Tool

`createPi4WordJavaScriptReplTool()` lives in `src/assistant-tools/javascript-repl-tool.js` and wraps pi-web-ui's `javascript_repl`.

Contract differences from the upstream tool:

- The schema remains a single root object.
- `title` is required.
- `code` is accepted.
- `script` is accepted as an alias for `code`.
- At least one of `code` or `script` must be a non-empty string.

## Settings And Storage

- **Storage backend:** `initPiWebStorage()` creates pi-web-ui `AppStorage` with IndexedDB-backed settings, provider keys, sessions, and custom providers.
- **Legacy migration:** `migrateLegacyLocalStorageOnce()` may migrate `pi4word.settings.v1` from `localStorage` into provider keys, preferred model keys, and Pi4Word stream proxy settings.
- **Preferred model:** stored in AppStorage settings as `pi4word.chat.preferredProvider` and `pi4word.chat.preferredModelId`.
- **Preferred thinking level:** stored in AppStorage settings as `pi4word.chat.preferredThinkingLevel` and mirrored to `localStorage` as a synchronous embedded-host fallback/cache because IndexedDB writes may not finish before unload.
- **CORS proxy:** pi-web-ui proxy settings use `proxy.enabled` and `proxy.url`.
- **Pi stream proxy:** Pi4Word settings use `pi4word.streamProxy.enabled`, `pi4word.streamProxy.url`, and `pi4word.streamProxy.token`.

## Networking

LLM traffic uses `createPi4WordStreamFn()` in `src/assistant/pi-assistant.js`.

- If Pi4Word stream proxy is enabled and both URL and token are present, requests use `streamProxy` from `@mariozechner/pi-agent-core`.
- The stream proxy URL is passed as `proxyUrl`; the token is passed as `authToken`.
- Otherwise requests use pi-web-ui `createStreamFn()` with the optional CORS proxy URL.
- This repository does not ship a backend. Production deployments must provide their own stream proxy endpoint when secure server-side key handling is required.

## Build And Assets

- `scripts/esbuild.mjs` bundles `src/index.js` to `public/index.min.js`.
- CSS is emitted to `public/index.min.css`.
- `office-js` is external.
- `process` is supplied by the esbuild inject/alias path `src/shims/process.js`.
- KaTeX font assets are emitted under `public/assets/`.
- `scripts/esbuild.mjs` copies `node_modules/pdfjs-dist/build/pdf.worker.min.mjs` to `public/pdfjs-dist/build/pdf.worker.min.mjs` so pi-web-ui can load the PDF.js worker beside the bundle.
- The bundle includes `marked`, DOMPurify, Turndown, and `turndown-plugin-gfm`.

## Source Boundaries

- Feature code lives under `src/<feature>/`.
- `src/task-pane/` is the composition root for assistant, settings, and pi-web-ui.
- `src/index.js` imports `src/task-pane/index.js` as the only feature entry, plus entry shims and styles.
- `task-pane` consumes `assistant` and `settings` only through their `index.js` barrels.
- The `assistant` barrel exposes agent creation, `createWordTools()`, `createPi4WordJavaScriptReplTool()`, and `registerCollapsibleWordToolRenderers()`.
- `src/assistant-tools/index.js` exposes tool factories for tests or custom tool sets.
