# AGENT.md

## Start here

1. Read **`SPECS.md`** — product summary, constraints, functional spec only (not a backlog). Update it when **intent**, **constraints**, or **required behavior** change; skip for tiny fixes that affect none of those.
2. Prefer **`src/<feature>/`** with **one primary change type** ([below](#change-types)). Supporting edits (tests, barrel **`index.js`**, **`SPECS.md`**) do not count as a second primary type.
3. **Few files:** target **≤3**, allow up to **5** when justified (e.g. unit + test + barrel touch).
4. **Before finishing:** **`npm run eslint`** and **`npm test`** from repo root (unless the user says to skip).

If **`SPECS.md`** is missing, recreate it with the same three sections as the codebase.

---

## Repo layout

* Features live under **`src/<feature>/`** (**`assistant`**, **`settings`**, **`task-pane`**). **`task-pane`** composes **`assistant`**, **`settings`**, and **`@mariozechner/pi-web-ui`**.
* **`src/assistant-tools/`** — Word **`Word.run`** tools (**`index.js`** exports **`createWordTools`**). **`assistant`** imports **`assistant-tools`**; **`task-pane`** uses **`assistant`** only (via **`assistant/index.js`**).
* **`src/shims/`** — small entry-time modules used from **`src/index.js`** (e.g. **`office-alert.js`** for Word hosts). **`process`** is supplied via esbuild **inject**/alias, not as a normal source import. **`scripts/esbuild.mjs`** may copy third-party assets into **`public/`** (PDF.js worker for pi-web-ui); see **Build/runtime** in **`SPECS.md`**.
* **Do not import another feature’s internals** (`…/other-feature/foo.js`). **`task-pane`** may import **`assistant`** and **`settings`** only via **`src/<feature>/index.js`**; inside **`task-pane`**, use sibling **`task-pane.*.js`** as needed.
* **`assistant`** / **`settings`**: **`index.js`** is the **only** import surface for **`task-pane`** — re-exports + **short top-of-file JSDoc** (purpose, main side effects, persistence). Update when **exports**, **cross-feature usage**, **persistence**, or **major side effects** change—not for trivial internal tweaks.
* **`task-pane`**: **`index.js`** is the only **`src/<feature>/`** import from **`src/index.js`** (typically **`initializeTaskPane`**); use a **short top-of-file JSDoc** and keep it in sync when that entry contract changes.

---

## Units and layering

| Guideline | Detail |
| --- | --- |
| Size | Aim **50–150 lines** per file; split past **~150**. |
| Flow | **`entry`** (orchestration) → **`boundary` / `state` / `api` / `model`**. Boundary may call **API**; **API** / **state** must not depend on **boundary**. |
| Roles | Prefer names like **`*.init.js`**, **`*.boundary.js`**, **`*.model.js`**; avoid **`utils`** / **`manager`**. |
| State | Local or injected; no hidden globals (except documented composition singletons). |
| Errors | Each layer handles its own failures; don’t leak raw errors across boundaries. |

**Tests:** Colocated **`*.test.js`**; prioritize **model** → **state** → **boundary** → **api**. Pure helpers should have tests; Office / Word-only paths can stay manual.

---

## Change types

Exactly **one primary** type per task:

* **render** — DOM / task pane shell / static HTML / styles  
* **state** — persisted or in‑memory app/session state  
* **model** — pure validation / formatting / mapping  
* **api** — Office.js, streams, IndexedDB bootstrap, hosts  
* **contract** — barrel **`index.js`**, exports, shared types  

---

## Prefer / avoid

* Prefer **small, explicit APIs** and **limited duplication over premature abstraction**.
* Avoid mixed responsibilities, **cross‑feature internals**, hidden shared state, large abstractions, oversized files, vague tests.

**Refactor first** when the same fix would span **more than five files**, the **public surface** is unclear, tests sprawl across concerns, or a prompt needs most of the repo.

---

## LLM prompts (optional)

Pick **one primary file**; touch at most **5** unless restructuring. Mention the barrel **`index.js`** when contracts change, and cite an existing **`*.test.js`** when relevant.

**Pi4Word:** **`SYSTEM_PROMPT`** lives in **`src/assistant/pi-assistant.js`**; **Word tools** live under **`src/assistant-tools/`** (**`index.js`** exports **`createWordTools`** and per-tool factories; **`word-tool-*.js`**, **`word-tools-search.js`**, **`word-tools-shared.js`**). When you change selection/read semantics, insert placement, or model instructions, update **`SPECS.md`** (product summary and **Assistant** bullet) so the spec, prompt, and tools stay aligned.
