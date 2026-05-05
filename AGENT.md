# AGENT.md

## Required Workflow

1. Read `SPECS.md` before changing source. Treat it as the product contract, not a backlog.
2. Choose one primary change intent per task. Supporting edits to tests, barrels, or `SPECS.md` are allowed.
3. Keep changes small: target 3 files or fewer; use up to 5 when a test, barrel, or spec update is justified.
4. Run `npm run eslint` and `npm test` from the repo root before finishing, unless the user explicitly says to skip.

If `SPECS.md` is missing, recreate it from the current source using the same high-level structure: product, runtime, assistant/tools, settings/storage, networking, build/assets, and source boundaries.

## Repository Boundaries

- Features live under `src/<feature>/`.
- `src/task-pane/` is the composition root for `assistant`, `settings`, and `@mariozechner/pi-web-ui`.
- `src/index.js` imports `src/task-pane/index.js` as the only feature entry, plus entry shims and global styles.
- `task-pane` consumes `assistant` and `settings` only through their `index.js` barrels.
- Inside `task-pane`, use sibling `task-pane.*.js` modules for local composition.
- Do not import another feature's internals from `task-pane` or other feature roots.

## Feature Surfaces

- `assistant/index.js` is the public assistant contract for task-pane wiring. Keep its top-of-file JSDoc current when exports, side effects, or persistence assumptions change.
- `settings/index.js` is the public settings contract for task-pane and assistant wiring. Keep its JSDoc current when exports, storage behavior, or migration behavior changes.
- `task-pane/index.js` is the only task-pane entry consumed by `src/index.js`. Keep its JSDoc current when startup or composition responsibilities change.
- `src/assistant-tools/index.js` exposes tool factories for tests and custom tool sets. Product-level tool behavior and schemas belong in `SPECS.md`.
- `src/shims/` contains small entry-time modules used by `src/index.js`; `process` is supplied by esbuild inject/alias, not normal source imports.

## Change Discipline

Choose one primary intent:

- `render`: DOM, task pane shell, static HTML, styles
- `state`: persisted or in-memory app/session state
- `model`: pure validation, formatting, mapping
- `api`: Office.js, streams, IndexedDB bootstrap, host integration
- `contract`: barrels, exports, shared public surfaces

Prefer small explicit APIs and limited duplication over broad abstractions. Avoid mixed responsibilities, hidden shared state, oversized files, vague tests, and cross-feature internals.

Refactor first when the same fix would span more than five files, the public surface is unclear, tests sprawl across concerns, or a prompt needs most of the repo.

## File Organization

- Aim for 50-150 lines per source file; split past roughly 150 lines when the split creates a clearer responsibility.
- Prefer names such as `*.init.js`, `*.boundary.js`, `*.model.js`, `*.state.js`, and `*.api.js`.
- Avoid vague buckets like `utils` and `manager`.
- Keep flow directional: entry orchestration -> boundary/state/api/model. Boundary may call API; API and state should not depend on boundary.
- Keep state local or injected unless a composition singleton is documented.
- Each layer should handle its own expected failures and avoid leaking raw low-level errors across boundaries.

## Testing And Specs

- Place tests next to the code under test as `*.test.js`.
- Prioritize tests for model/state helpers; Office/Word-only paths may remain manual when automation is impractical.
- Update `SPECS.md` when a change affects product intent, constraints, required behavior, tool schemas, streaming/settings keys, storage behavior, build/runtime assumptions, source boundaries, or model instructions.
- Do not update `SPECS.md` for tiny internal fixes that do not change observable behavior or contracts.
