# AGENT.md

## Specifications (mandatory)

1. **At the beginning of every task or session**, read **`specifications.md`** in the repo root. It is the live project brief: product intent, current scope, implemented vs planned items, key files, and commands.
2. **Keep `specifications.md` current.** After meaningful progress (new features, removed features, new scripts, build steps, layout changes, or important decisions), update that file in the same change set when practical. If the user only asked for a small fix, at least adjust the **Current status**, **Not implemented**, or **Changelog** sections so the next reader sees reality.

If `specifications.md` is missing, create it using the same structure as the repo’s current state.

---

## Purpose

Optimize for:

* minimal LLM context
* strict modularity
* predictable, local changes

---

## Core Invariants

* A change requires **≤3–5 files**
* A feature is **modifiable in isolation**
* Every change has a **matching local test**

---

## Feature Structure

```text
features/<feature>/
```

In **this repo**, feature roots live under **`src/features/<feature>/`** (same rules).

* No cross-feature internal dependencies
* All logic for a feature is local

---

## Roles (Single Responsibility)

Each unit has one role:

| Role     | Responsibility                             |
| -------- | ------------------------------------------ |
| Entry    | orchestration                              |
| Boundary | interaction interface (UI / handler / CLI) |
| State    | state management                           |
| API      | external I/O                               |
| Model    | pure logic                                 |
| Test     | behavior verification                      |

Use **platform-idiomatic naming** while preserving roles.

---

## Dependency Direction

```text
entry → (boundary | state | api | model) → core
```

Rules:

* boundary **may call API**
* API must not depend on boundary
* state must not depend on boundary
* no cross-feature internals

---

## Public Contract

Every unit must define:

* inputs
* outputs
* side effects

Contracts must be:

* explicit
* minimal
* documented in README

Public APIs are stable by default.

---

## State Rule

State must be:

* local, or
* explicitly passed/injected

Forbidden:

* hidden globals
* implicit shared state

---

## Change Types (Mandatory)

Each task is exactly one:

* render
* state
* model
* api
* contract

Do not mix types.

---

## Boundaries

* Small explicit API
* No hidden dependencies
* No unrelated logic
* ≤150 lines

---

## File Size

* Target: **50–150 lines**
* Split if exceeded

---

## Reuse

* Prefer duplication over abstraction
* Move to core only if reused + stable

---

## Tests

* Mandatory per feature
* Colocated
* One test per unit

Priority:
model → state → boundary → api

Tests:

* verify behavior
* remain local
* avoid cross-feature scope

---

## Error Handling

Each layer handles its own errors.
Do not leak raw errors across layers.

---

## README (Mandatory)

`features/<feature>/README.md`

Max 20 lines:

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

## LLM Rules

### Prompt Scope

* Target **1 file**, max **3–5 files**
* Always include:

  * README
  * relevant test

---

### Prompt Format

```text
Modify only:
- <file>
- <test file>

Do not change:
- other files
- public APIs
```

---

### Hard Limits

* > 5 files → refactor first
* > 150 lines → split first

---

### Output

* Only changed files
* No unrelated refactors
* Tests updated

---

## Anti-Patterns

* mixed responsibilities
* cross-feature dependencies
* hidden state
* large abstractions
* oversized units
* broad or unclear tests

---

## Refactor Triggers

Refactor if:

* > 3 files needed
* unclear API
* test scope too large
* prompt grows

---

## Summary

This enforces:

* minimal context
* explicit contracts
* strong isolation
* reliable LLM behavior

> Any change must be achievable by modifying a small, local part of the system.
