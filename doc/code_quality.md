# Code Quality Reference

Short reference for code quality decisions in fastsheet (Tauri 2 +
SvelteKit + Rust + IronCalc). Reflects general engineering practice
plus the specific goals of this project.

## Primary Goals

The codebase should optimize for:

- correctness for spreadsheet operations: formula evaluation, file
  format round-trip (xlsx + xls preservation), keyboard navigation
- predictable behavior across the Tauri command boundary (frontend ↔
  backend) and across the IronCalc evaluator
- low coupling between UI composition and reusable domain logic — the
  backend should be callable both from Tauri commands and headless
  probe binaries
- maintainability as the project grows in formats, functions, and
  features
- incremental modularization without destabilizing working features
- keyboard responsiveness as a non-negotiable; the arrow-key hot path
  has explicit perf rules in CLAUDE.md that override generic style
  preferences

## Core Principles

### 1. Prefer clear ownership

Each piece of logic should have an obvious owner.

- The IronCalc `Model` (held in `AppState.model`) is the workbook's
  source of truth. Sheet data, formulas, styles, defined names — all
  live there.
- `AppState` is the backend's shared-mutable state container; each
  field is `Mutex`-wrapped and has a single clear purpose
  (`dirty`, `style_dirty`, `structural_dirty`, `loaded`, `compare`,
  `xls_preserved`, `auto_recalc`, etc.).
- Each backend module owns its domain: `workbook.rs` (open/save),
  `cells.rs` (cell read/write/style), `compare.rs` (diff sessions),
  `trace.rs` (formula tracing), `xls_load.rs` / `xls_save.rs` (BIFF
  pipeline), `atomic.rs` (atomic file replacement + backups),
  `index.rs` (recents).
- `+page.svelte` is the frontend composition root: keyboard
  dispatcher, menu wiring, top-level state. Not the default home for
  new business logic.
- Svelte components (`Grid.svelte`, `Navigator.svelte`,
  `FormulaTrace.svelte`, `CompareDiff.svelte`,
  `FormatCellsDialog.svelte`) own UI concerns and receive their data
  via props.

If ownership is unclear, coupling and duplication usually follow.

### 2. Keep coupling low

New code should depend on narrow interfaces and stable abstractions.

- Backend functions should take the specific state they need
  (`&Model`, `&CompareSession`, a path) rather than `&AppState`
  whenever possible — makes them testable in isolation and reusable
  from probe binaries.
- Svelte components should accept props for the data they need; avoid
  reaching for top-level `+page.svelte` state from a deep component.
- Avoid reaching into IronCalc internals beyond the public API + our
  vendored patches. New IronCalc-specific quirks should be wrapped in
  a focused helper rather than spread across callsites.
- Frontend code should not assume the on-the-wire shape of a Tauri
  command beyond what `types.ts` declares.

### 3. DRY, but only for real duplication

Avoid copy/paste logic, especially when behavior must stay aligned
across:

- xls and xlsx load/save paths
- Tauri command path and headless probe binaries
- frontend menu callbacks and direct keyboard shortcuts

Do not force abstractions too early. Extract shared code when:

- the behavior is already duplicated
- the same bug would need to be fixed in multiple places
- tests would otherwise duplicate the same expectations

The fastsheet preference (per CLAUDE.md) is "three similar lines is
better than a premature abstraction." Wait for the third occurrence
before generalizing.

### 4. Prefer modular cores

Reusable logic should live in modules that can be called from:

- Tauri command handlers
- the `bin/probe.rs`, `bin/save_probe.rs`, `bin/trace_probe.rs`
  headless probes
- integration tests (`tests/xls_roundtrip.rs`)

Tauri command wrappers add command-only concerns:

- mutex acquisition + release ordering
- `Result<_, String>` formatting for the wire
- AppState side effects (e.g. `dirty.clear()` after save,
  `record_open_internal` for recents)

The pure-Rust core (`xls_save::save_xls`, `compare::diff_workbooks`,
`trace::trace`) takes a `&Model` and returns a typed result — that's
what the probes and tests call.

### 5. Preserve boundaries

Two boundaries matter in this project:

**Frontend ↔ backend (Tauri command channel)**

- Cross-boundary communication uses `invoke` only — never poke at
  Tauri internals or window globals.
- Every backend command returns `Result<T, String>` where `T:
  Serialize`. The frontend deserializes and surfaces errors via
  `statusMsg`.
- Don't introduce shared mutable state across the boundary (e.g. via
  `tauri-plugin-store` for ephemeral state). Source of truth stays in
  `AppState` on the backend; the frontend mirrors snapshots.

**Mutex acquisition order in AppState**

- Acquire locks for the shortest duration possible. Drop guards
  before any IO or any call that might re-enter another command.
- Never hold an `AppState` mutex across an `await` (Rust's
  `std::sync::Mutex` is not safe across awaits anyway).
- IronCalc's `Model` is `Send` but not `Sync` — only one thread holds
  it at a time via `state.model.lock()`.

### 6. Make state flow explicit

Prefer explicit data flow over hidden mutation.

- Functions take the state they need.
- Backend commands return typed result structs (`SaveResult`,
  `CompareResult`, `WorkbookInfo`, `TraceNode`, etc.); the frontend
  doesn't have to refetch.
- Avoid "set state, then call another command" when a single command
  can take both inputs.
- Name state for the exact thing it represents. In particular,
  separate unsaved-workbook, save-patcher, style-dirty,
  structural-dirty, and recalc-dirty concepts. A generic name like
  `dirtyEdits` is only acceptable when it really drives one behavior;
  if it feeds both "unsaved changes" and "pending recalc" UI, split it.

## Project-Specific Rules

### `+page.svelte` should stay thin

`+page.svelte` may:

- own top-level reactive state (cursor, viewport, edit mode, modal
  flags)
- dispatch keyboard events
- wire menus and callbacks
- compose Svelte components

`+page.svelte` should not:

- become the default home for new business logic
- accumulate feature-specific state when a focused component or store
  could own it
- reach across into a child component's internals

When `+page.svelte` grows past comfortable size, the lift is usually
to extract a feature into a self-contained `lib/` component
(`Navigator.svelte`, `FormulaTrace.svelte`, `CompareDiff.svelte` are
prior examples).

Dialogs and panels should call semantic backend commands. A bulk
formatting dialog should say "set bold to true" / "set border sides"
rather than synthesising a sequence of toggle calls whose final value
depends on a particular cell's current state.

### Reuse the IronCalc `Model` instead of parallel state

If a feature is part of the workbook (cells, formulas, styles, sheet
names, defined names, frozen panes), it lives in the IronCalc `Model`
or in `AppState`'s side-channel maps that shadow it (e.g.
`hidden_cols` because IronCalc's `Col` doesn't carry a `hidden`
field).

What does NOT belong in the model:

- selection, cursor, viewport scroll position
- modal flags (edit mode, navigator open, trace popup, compare dock)
- menu descent state
- dirty-edit set (this lives in `AppState.dirty`, not in the model
  itself, because the model evaluates everything in place)

Different "dirty" states mean different things:

- `dirty`: cell content edits that the xlsx preservation patcher can
  project back into sheet XML.
- `style_dirty`: formatting changes that currently require full
  serializer save because styles.xml is not patched safely.
- `structural_dirty`: row/column/sheet changes that make coordinate-
  based xlsx patching unsafe.
- `workbook_dirty`: user-visible unsaved changes. This drives
  title/status markers and discard prompts, not save strategy.
- recalc-pending frontend state: user-visible stale formula values in
  manual recalculation mode only. Automatic recalculation should not
  increment this counter.

### Respect the modularization direction

Recent refactoring split monolithic logic into focused modules:

- `compare.rs`, `trace.rs`, `xls_preserve.rs`, `xls_save.rs`,
  `xls_load.rs`, `xls_biff.rs`, `hidden.rs`, `index.rs`,
  `navigator.rs`
- `lib/Navigator.svelte`, `lib/Grid.svelte`,
  `lib/FormulaTrace.svelte`, `lib/CompareDiff.svelte`,
  `lib/SheetTabs.svelte`

New code should continue that direction, not undo it. A feature with
its own state, keyboard handling, and rendering belongs in a new
component / module rather than threaded through `+page.svelte` or
`workbook.rs`.

### Keep headless probes in step with the GUI

`probe.rs`, `save_probe.rs`, `trace_probe.rs` exercise the same
backend modules the GUI calls. They are the cheapest way to reproduce
a load/save bug without spinning up the WebView. New backend logic
that operates on a `Model` should be callable from a probe; if it
isn't, that's a sign too much logic ended up in the Tauri command
wrapper instead of the core module.

### The keyboard hot path overrides generic style preferences

CLAUDE.md documents seven load-bearing techniques behind arrow-key
responsiveness (rAF-coalesced moveSel, cumulative-offset arrays,
viewport row virtualization, overlay-based selection, event
delegation, measured `colhdrH`, `flushSync` before scroll). These
techniques violate "prefer pure transformations over mutation" on
purpose — re-allocating prefix-sum arrays per keystroke or
rebinding `class:sel` on every cell would tank perf.

When working in `Grid.svelte` or the `+page.svelte` keyboard
dispatcher: re-read the playbook before refactoring. Do not undo any
of the seven in isolation.

### Save atomically and back up deliberately

Workbook saves must preserve the user's existing file on failure.

- Never use `remove_file(target); write(target)` for workbook output.
  Use `atomic::write`, which writes a sibling temp file, fsyncs it,
  then renames it over the destination.
- When a save path may lose unsupported workbook features, create a
  recoverable backup before the lossy write. Use `atomic::backup_if_exists`
  so `.bak`, `.bak.1`, `.bak.2`, ... are preserved instead of
  overwritten.
- Avoid double-backup flows. If a command explicitly creates a backup
  and then calls a save path that can also create a backup, pass that
  fact through the API or split the save helper so only one backup is
  produced and reported.
- Keep AppState locks out of filesystem work where practical. Snapshot
  the model data, preservation bundle, and side-channel maps under
  lock, drop the guards, then do backup/write/rename.
- `SaveResult` should report the actual mode and backup path so the
  frontend does not infer preservation or lossiness from file
  extensions.

## Control Flow and Complexity

### 7. Crash early, fail fast

A dead program causes less damage than a crippled one — but in Rust
"crash" usually means returning `Err` early, not panicking.

- validate preconditions at the top: invalid input → return `Err`
  immediately rather than guarding 200 lines deep
- use `?` and `let-else` to keep happy paths flat
- use `assert!` / `debug_assert!` for invariants that should never be
  violated in correct code (sheet index in range, prefix-sum array
  consistency, etc.)
- prefer early returns over deep nesting

```rust
// BAD — deep nesting
fn process(data: Option<&Data>, mode: Mode) -> Result<Out, Error> {
    if let Some(data) = data {
        if mode == Mode::Fast {
            // ...
        } else {
            // ...
        }
    } else {
        // ...
    }
}

// GOOD — guard clauses, flat structure
fn process(data: Option<&Data>, mode: Mode) -> Result<Out, Error> {
    let data = data.ok_or(Error::MissingData)?;
    if mode == Mode::Fast {
        return fast_path(data);
    }
    slow_path(data)
}
```

In TypeScript / Svelte, the equivalent is `throw new Error(...)` or
returning early from a callback after surfacing `statusMsg`.

### 8. Design by contract

Every function has a contract: what it expects (preconditions) and
what it guarantees (postconditions). Make these explicit through the
type system whenever possible.

- prefer typed inputs (`u32` sheet index, `&Path`, an enum) over
  loosely-typed (`&str`, `i32`)
- prefer typed outputs (`Result<SaveResult, String>`,
  `Option<RecentDir>`) over `()` with side effects
- when a function's contract is clear, callers don't need defensive
  checks
- when the contract is violated, fail immediately (see 7)

In Rust, the type system carries most of the contract. Reserve
runtime checks for cases the type system can't express
(`row >= 1 && row <= MAX_ROW`, formula text length limits, etc.).

### 9. Prefer transformations over state mutation — except in the hot path

Treat data processing as a pipeline of transformations, not a series
of mutations to shared state.

- pure functions that take input and return output are easier to
  test, compose, and reason about than methods that mutate `self` /
  `&mut state`
- move computation out of deep loops into separate functions
- when state must be mutated, confine it to one place and keep the
  transformation logic stateless

```rust
// BAD — mutation inside a loop with branching
let mut out = Vec::new();
for diff in diffs {
    if diff.kind == "value" {
        diff.label = format_value(&diff);
    } else if diff.kind == "formula" {
        diff.label = format_formula(&diff);
    }
    out.push(diff);
}

// GOOD — pure transformation
fn label_for(diff: &Diff) -> String {
    match diff.kind {
        "value"   => format_value(diff),
        "formula" => format_formula(diff),
        _         => String::new(),
    }
}
let out: Vec<_> = diffs.into_iter()
    .map(|mut d| { d.label = label_for(&d); d })
    .collect();
```

**Hot-path exception:** `Grid.svelte` and the `+page.svelte` arrow-key
dispatcher use mutation deliberately (rAF accumulation, prefix-sum
geometry arrays, overlay style updates). Don't apply this rule to
those files without re-reading the keyboard perf playbook.

### 10. Law of Demeter — talk only to friends

A function should only call methods on:

- its own object / module
- objects passed as parameters
- objects it creates
- its direct components

Avoid reaching through chains of objects:

```ts
// BAD — deep coupling
state.compare.session.right_model.workbook.worksheets[0].sheet_data.get(1)

// GOOD — ask the object you have
session.right_value_at(left_model, sheet, row, col)
```

In `+page.svelte`, the common violation is reaching from a callback
into deeply-nested grid state. Prefer extracting a helper that takes
the values it needs as arguments.

In Rust, reaching through `state.compare.lock().unwrap().as_ref()...`
chains is OK at the Tauri command boundary (that's what the boundary
exists to do), but the pure-Rust core inside should take a
`&CompareSession` and not know about `AppState`.

### 11. Keep functions short and focused

A function should do one thing. Signs it does too much:

- more than 3 levels of indentation
- more than ~40 lines (Rust functions can grow past this when they
  carry typed setup; use judgment)
- multiple unrelated blocks separated by blank-line "section
  headers"
- a name that needs "and" to describe what it does

Split along natural seams. In this codebase the natural seams are
the existing module boundaries — if a function is doing both BIFF
encoding and IronCalc Node walking, that's two functions.

### 12. Flatten conditional logic

Deep nesting is the primary source of fragility.

- replace nested `if/else` with guard clauses and early returns
- replace long `if/else if` chains over a discriminant with `match`
  (Rust) or a dispatch table / `switch` (TS)
- replace flag variables (`let mut found = false; for ...; if found`)
  with `iter().find()` / `next()` / extracted helpers

```rust
// BAD — nested branches
if let Some(session) = compare.as_ref() {
    if let Some(left) = model.as_ref() {
        if let Some(name) = left.workbook.worksheets.get(idx).map(|w| &w.name) {
            // ...
        }
    }
}

// GOOD — guard clauses
let Some(session) = compare.as_ref() else { return None };
let Some(left) = model.as_ref() else { return None };
let Some(ws) = left.workbook.worksheets.get(idx) else { return None };
let name = &ws.name;
// ...
```

## Standard Engineering Rules

### Favor readability over cleverness

- choose straightforward control flow
- prefer descriptive names over short opaque names
- keep functions focused
- comment only where intent is non-obvious from the code itself
  (per CLAUDE.md: default to no comments; the WHY-not-WHAT rule is
  strict here)

### Keep APIs narrow

- Tauri commands should be the smallest set that the frontend
  actually uses; don't expose internal helpers as commands
- types in `types.ts` should be the minimum surface
- prefer typed serializable structs over loose `serde_json::Value`
- don't leak internal implementation details into the wire format

### Fail clearly

- raise explicit errors (`return Err("...")` in Rust, `throw new
  Error("...")` or `statusMsg = ...` in TS) for invalid state or
  unsupported operations
- avoid silent failure unless it is intentionally best-effort (the
  recents `record_open_internal` is best-effort because a SQLite
  hiccup shouldn't fail an open; that's documented at the call site)
- include enough context in errors to debug the real problem (the
  path that failed, the cell address, the sheet name)

### Write tests at the right level

Prefer tests that verify behavior, not implementation trivia.

- the existing `tests/xls_roundtrip.rs` is the canonical example: it
  builds a `Model`, saves it, reloads it, and asserts cell-by-cell
  equality. It doesn't care which intermediate functions ran.
- in-tree `#[test]` modules in `xls_save.rs` cover framing
  invariants (BOUNDSHEET8 lbPlyPos, required globals records) that
  external readers don't catch.
- frontend tests are sparse currently; behavior is verified by
  building the app and exercising it. Be explicit when reporting "I
  didn't test the UI."

### Keep docs aligned with reality

`CLAUDE.md` is the project's source-of-truth doc and is loaded into
every Claude session. Update it when:

- a new module is added (description line under "Rust backend" / "Svelte frontend")
- a non-obvious behavior is discovered (add to "Non-obvious behaviours" or "Gotchas")
- a perf-critical pattern is established (add to playbook)

`QUICKSTART.md` is the user-facing docs. Update it when a new
keyboard shortcut, menu item, or save behavior is added.

`memory/` is auto-loaded across conversations. Use it for cross-
session reference (API quirks, build recipes), not for ephemeral
work-in-progress notes.

### Refactor opportunistically — but don't over-correct

Per CLAUDE.md: "Don't add features, refactor, or introduce
abstractions beyond what the task requires." Combined with refactor-
when-the-cost-is-low: if you're already touching a file and notice
its conditional nesting could flatten cleanly without expanding the
diff, do it in the same commit. If the refactor would balloon the
diff, file a follow-up — don't bundle a working fix with a
speculative cleanup.

## Practical Review Checklist

When adding or reviewing code, ask:

1. Does this introduce unnecessary coupling to `+page.svelte` /
   `AppState`?
2. Is there duplicated logic that should move into a shared module
   (`xls_*`, `compare`, `trace`, `lib/*.svelte`)?
3. Does this respect the frontend ↔ backend boundary? Does it hold
   any mutex across an await?
4. Is the state source of truth clear (Model? AppState side-channel?
   Frontend `$state`?)
5. Could this Tauri command surface be narrower?
6. Are error paths surfaced explicitly via `Result<_, String>` /
   `statusMsg`?
7. Do the round-trip tests still pass? Should a new one be added
   (`tests/xls_roundtrip.rs`)?
8. Does CLAUDE.md / QUICKSTART.md need updating because of this
   change?
9. More than 3 levels of nesting? Can guard clauses or `let-else`
   flatten it?
10. Function exceeds ~40 lines? Can it split along a natural seam?
11. Long method chains violating Demeter? (Especially deep
    `state.X.lock().unwrap().as_ref()...` outside the immediate
    Tauri command boundary.)
12. Could a loop body or branch be a pure transformation function?
13. Are preconditions validated at the top, or buried in deep
    branches?
14. Does a bulk UI action call semantic "set" commands, or is it
    composing toggles whose result depends on an arbitrary seed cell?
15. Does every workbook save use `atomic::write` and at most one
    deliberate backup path?
16. Are save-time mutex guards dropped before backup / filesystem
    writes where a snapshot would suffice?
17. Are save-dirty, style-dirty, structural-dirty, and recalc-dirty
    state kept separate?
18. **Hot path check:** if the change touches `Grid.svelte` or the
    arrow-key dispatcher, does it preserve the seven load-bearing
    techniques?

## Rule of Thumb

Good code in this project is:

- explicit
- testable (incl. via `bin/*` probes when applicable)
- low-coupling (modules don't reach for `AppState` when a focused
  argument would do)
- thread-safe at the AppState boundary (short-lived locks, no holds
  across awaits)
- reusable across the GUI command path and headless probes where
  appropriate
- flat (guard clauses, not deep nesting)
- contractual (clear preconditions and postconditions, expressed in
  the type system where possible)
- transformational (pure functions over state mutation) **except in
  the keyboard hot path**, where the perf playbook wins

If a change makes the feature work but pushes more hidden behavior
into `+page.svelte`, duplicates logic across xls/xlsx or
GUI/probe paths, weakens AppState lock discipline, or adds deep
conditional nesting, it is probably moving the codebase in the wrong
direction.
