# AGENT.md

## Specifications (mandatory)

1. **At the beginning of every task or session**, read **`specifications.md`** in the repo root. It is the live project brief: product intent, current scope, implemented vs planned items, key files, and commands.
2. **Keep `specifications.md` current.** After meaningful progress (new features, removed features, new scripts, build steps, layout changes, or important decisions), update that file in the same change set when practical. If the user only asked for a small fix, at least adjust the **Current status**, **Not implemented**, or **Changelog** sections so the next reader sees reality.

If `specifications.md` is missing, create it using the same structure as the repo’s current state.

**Project map:** Product layout and key file paths live in **`specifications.md`** (especially **Key files**). Do not duplicate that table here.

---

## Agent quickstart

1. Read **`specifications.md`**.
2. Identify the **feature** under **`src/<name>/`** (e.g. **`assistant`**, **`settings`**, **`task-pane`**) and stay inside it unless you are editing the **composition root** (`task-pane` entry), **`specifications.md`**, or a barrel **`index.js`**.
3. Pick a **primary change type** (see [Change types](#change-types)). Supporting edits (tests, **`README`**, typings) are normal and **do not** count as mixing types.
4. Touch **few files** (see [Limits](#limits)).
5. Finish with:

```bash
npm run eslint
npm test
```

---

## Purpose

Optimize for:

* minimal LLM context
* strict modularity
* predictable, local changes

---

## Limits

| Rule | Detail |
| --- | --- |
| Files per change | **Target ≤3**; **allow up to 5** when the task genuinely needs e.g. unit + test + README + a barrel `index.js`. |
| Unit size | **Target 50–150 lines** per file; split if a unit grows past **~150 lines**. |
| Cross-feature imports | Feature **A** must not import **internals** of feature B (`…/feature-b/some-internal.js`). **Exception:** the **composition root** (**`src/task-pane/`** entry modules such as **`task-pane.init.js`**) may import other features **only** via their **public surface** — prefer **`src/<feature>/index.js`** when present. |

---

## Feature structure

Feature packages live **directly under `src/`** (there is no `src/features` folder).

```text
src/<feature>/
```

* Non-root features stay **inward-facing**: no imports of sibling feature **internals**.
* **`task-pane`** composes **`assistant`**, **`settings`**, and **`@mariozechner/pi-web-ui`**.

---

## Roles (single responsibility)

Each unit has one role:

| Role     | Responsibility                             |
| -------- | ------------------------------------------ |
| Entry    | orchestration                              |
| Boundary | interaction interface (UI / handler / CLI) |
| State    | state management                           |
| API      | external I/O                               |
| Model    | pure logic                                 |
| Test     | behavior verification                      |

Use **platform-idiomatic naming** while preserving roles (e.g. `*.init.js`, `*.boundary.js`).

---

## Dependency direction

```text
entry → (boundary | state | api | model) → core
```

Rules:

* boundary **may call API**
* API must not depend on boundary
* state must not depend on boundary
* no cross-feature **internal** modules (composition root exception above)

---

## Public contract

Every unit must define:

* inputs
* outputs
* side effects

Contracts must be explicit, minimal, and documented in the feature **`README.md`**.

**When to update `README`:** new or changed **exports**, **side effects**, or **persistence** — not for every one-line fix.

Public APIs are stable by default.

---

## State rule

State must be **local** or **explicitly passed/injected**.

Forbidden: hidden globals, implicit shared state (except documented module singletons used as composition state).

---

## Change types

Each task has **one primary** type:

* **render** — DOM / layout / styles (**in this repo:** task pane shell CSS hooks, static HTML, visual structure)
* **state** — persisted or in-memory app/session state
* **model** — pure logic (validation, formatting, mapping)
* **api** — Office.js, network streams, IndexedDB bootstrap, external hosts
* **contract** — public exports, `README`, shared types between layers

**Supporting** edits (tests, colocated `*.test.js`, feature `README`, `specifications.md`) accompany the primary type; they are **not** a second primary type.

---

## Boundaries

* Small explicit API
* No hidden dependencies
* No unrelated logic
* ≤150 lines

---

## Reuse

* Prefer duplication over abstraction
* Move to core only if reused + stable

---

## Tests

* Mandatory per feature where meaningful (pure **model**/**state** helpers must have tests; Office / Word-only paths may stay manual or integration)
* Colocated `*.test.js`
* Priority: model → state → boundary → api

Tests verify behavior, stay local, and avoid unnecessary cross-feature scope.

---

## Verification after implementation (mandatory)

After **every implementation**, before treating work as complete:

1. Run **`npm run eslint`** from the repo root and fix reported problems (including warnings unless the project or user explicitly allows ignoring them).
2. Run **`npm test`** from the repo root and fix failures.

If either command fails, continue iterating until both succeed. Only skip when the user **explicitly** says to omit ESLint or tests for that task.

---

## Error handling

Each layer handles its own errors. Do not leak raw errors across layers.

---

## README (mandatory)

`src/<feature>/README.md`

Max ~20 lines:

```text
# <feature>

Purpose:
<1–2 lines>

Public API:
<inputs / outputs>

Side Effects:
<if any>

State:
<short>

Flow:
input → state → processing → output

Tests:
- boundary/state/model
```

---

## Naming

Use explicit names.

Good:

* `user-profile.boundary`
* `calculate-tax.model`

Bad:

* `utils`, `helpers`, `manager`

---

## Prompt format (LLM)

Target **one primary file**; touch at most **5 files** unless refactoring to reduce scope.

Always include the feature **`README`** and a **relevant test** when one exists.

```text
Modify only:
- <file>
- <test file>

Do not change:
- unrelated features
- public APIs (unless this task is a contract change)
```

**Output expectations:** only necessary files, no unrelated refactors, tests updated, then **`npm run eslint`** and **`npm test`**.

---

## Anti-patterns

* mixed responsibilities in one unit
* importing sibling feature **internals** (except composition root + public/barrel imports)
* hidden state
* large abstractions
* oversized units
* broad or unclear tests

---

## Refactor triggers

Refactor **before** growing a change when:

* the **same concern** would require **more than 5 files** without a structural cleanup
* the **public API** is unclear or unstable
* **tests** must cover too many behaviors at once
* an LLM **prompt** needs half the repo in context

---

## Summary

This enforces:

* minimal context
* explicit contracts
* strong isolation (with a **defined composition root**)
* reliable LLM behavior

> Any change must be achievable by modifying a small, local part of the system.
