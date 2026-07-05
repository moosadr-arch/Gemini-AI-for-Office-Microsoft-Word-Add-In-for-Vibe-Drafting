# System Architecture and Project Map

Role: system map for the external reconciliation package and its runtime consumers.

## Project Map

This repository consumes one reusable external package.

```mermaid
graph TD
    Core[@ansonlai/docx-redline-js\nexternal dependency] --> Word[Word Add-in\nsrc/taskpane]
    Core --> Demo[Browser Demo\nbrowser-demo]
    Core --> MCP[MCP docx server\nmcp/docx-server]
```

## Subprojects

1. `src/taskpane`: Word add-in host.
2. `src/taskpane/modules/docx-redline-js-integration`: add-in-only bridge layer around Word APIs.
3. `browser-demo`: browser runtime for manual validation and demo workflows.
4. `mcp/docx-server`: Node runtime exposing reconciliation as MCP tools.

## Core Boundary and Entrypoints

- Host-agnostic engine now comes from `@ansonlai/docx-redline-js`.
- Add-in local bridge entrypoint is `src/taskpane/modules/docx-redline-js-integration/index.js`.
- `src/taskpane/modules/reconciliation/` was removed after extraction.

## Import Path Conventions

- Word add-in command modules import the bridge:
  - `src/taskpane/modules/commands/agentic-tools.js` -> `../docx-redline-js-integration/index.js`
- Browser and MCP consumers import package entrypoints directly:
  - `browser-demo/demo.js` -> `@ansonlai/docx-redline-js`
  - `mcp/docx-server/src/services/docx-redline-js-service.mjs` -> `@ansonlai/docx-redline-js`

## Reconciliation Engine Overview

The engine reconciles text/markdown edits into Word-compatible OOXML with track changes.
Major paths include format-only, formatting removal, surgical edits, reconstruction, list generation, and table reconciliation.

Pipeline stages:

1. Ingestion
2. Markdown preprocessing
3. Word-level diffing
4. Patching
5. Serialization

## Two-stage LLM pipeline

A document edit request runs through two distinct model calls, not one:

1. **Tool selection (chat model).** The agentic loop in `src/taskpane/taskpane.js`
   sends the conversation + document context to Gemini with function-calling tools
   (`apply_redlines`, `edit_list`, `edit_table`, `insert_comment`, ...). The model
   chooses a tool and arguments.
2. **Change-set generation (structured-output model).** For `apply_redlines`, the
   tool (`executeRedline` in `modules/commands/agentic-tools.js`) makes a *second*
   Gemini call via `callGeminiForDiffs` using the shared prompt + schema in
   `modules/commands/redline-prompt.js` and `responseMimeType: "application/json"`
   with a `responseSchema`. This returns a JSON array of changes, each targeting a
   paragraph by `[P#]` index plus an `anchorText` snippet. The call has a 90s abort
   timeout (`DIFF_CALL_TIMEOUT_MS`) and a tighter `diffMaxOutputTokens` budget
   (model profile) so a degenerate/repetitive generation cannot hang the UI. If the
   returned JSON is truncated mid-stream, `repairTruncatedJsonArray` salvages the
   complete leading change objects rather than failing outright.
3. **Validation + application.** The change set is sanitized (`sanitizeChangeSet`)
   and anchor-verified (`verifyAnchor`) in `modules/commands/change-validation.js` —
   this also repairs wrong-field payloads (e.g. content written to `newContent`
   instead of `content`) and strips fields that don't belong to the chosen
   operation, so a model can never leak reasoning/commentary into document text via
   an unused field. Valid changes are applied via `applyRedlineChangeSet` →
   `applyRedlineChangesToWordContext` (the Word-API bridge). Off-by-one indices are
   auto-corrected; invalid changes are dropped. Targeting an empty paragraph, or the
   paragraph index one past the last one (append-at-end-of-document), is handled by
   the bridge with native `insertText`/`insertParagraph` calls
   (`insertContentAsNativeParagraphs`) instead of OOXML reconciliation, since Word
   rejects OOXML built from an empty diff target.
4. **In-tool corrective retry.** If a change set is rejected in full (nothing
   applied, so the document is untouched), `executeRedline` re-asks the diff
   generator once via `buildCorrectiveRetryPrompt` — showing it its own invalid
   output plus the machine-generated rejection reasons — before giving up. This
   catches the common "valid tool call, malformed diff JSON" failure mode without
   spending a chat-loop iteration or a loop-guard strike.
5. **Failure feedback.** If both diff attempts fail, rejected/failed changes are
   returned to the chat model as structured `TOOL_FAILURE …` messages (with the
   actual paragraph text, capped at 8 reasons + a summary line), so the next
   chat-level attempt is informed rather than blind.

The offline eval harness (`tests/evals/run-evals.mjs`) exercises stages 2–3 against
real models using the same shared prompt/schema and validation pipeline.

## Reliability layers

Defenses against per-model variance and brittleness, roughly outermost to innermost:

1. **Per-model profiles** (`modules/config/model-profiles.js`) — tokens (including a
   separate, tighter `diffMaxOutputTokens`), temperature, retries, and
   preview-throttle flags per model, instead of inline model-name checks.
2. **Structured output schema** — the diff call is constrained by a `responseSchema`.
3. **Change-set sanitizer** (`change-validation.js`) — enforces the prompt rules in
   code (valid operation, index range incl. append-at-end, `[P#]` stripping,
   wrong-field repair, empty/oversized/structural content guards, pseudo-table and
   schema-leak rejection, dedupe, stripping fields that don't belong to the chosen
   operation).
4. **Anchor verification** (`change-validation.js`) — an edit cannot silently land on
   the wrong paragraph; off-by-one is auto-corrected, ambiguous matches are rejected.
5. **Truncated-response salvage** (`repairTruncatedJsonArray`) — recovers complete
   change objects from a diff response cut off mid-stream (e.g. a model repetition
   loop exhausting its token budget) instead of failing the whole change set.
6. **Diff-call timeout** (`DIFF_CALL_TIMEOUT_MS`, `AbortController`) — bounds how
   long a single diff-generation call can run.
7. **Native-API fallback for empty targets** (`insertContentAsNativeParagraphs`) —
   inserting into an empty paragraph or appending at end-of-document uses Word's
   native, natively-tracked insertion instead of OOXML reconciliation, which Word
   rejects for empty diff targets.
8. **Auto-checkpoints** (`modules/storage/checkpoint-store.js`) — an IndexedDB snapshot
   is taken before each mutating tool, so any edit is recoverable.
9. **In-tool corrective retry** (`buildCorrectiveRetryPrompt`) — one extra diff-only
   retry with the model's own invalid output + rejection reasons, before failing the
   tool call up to the chat model.
10. **Informative retry feedback** — structured `TOOL_FAILURE` messages (capped size)
    drive corrected chat-level retries.
11. **History invariants** (`modules/chat/chat-history.js`, `appendFunctionExchange`) —
    function-call/response pairs are validated atomically before entering history.
12. **Loop guard** (`taskpane.js`) — stops repeated no-progress mutation cycles.

## Portability Status

Core reconciliation code is host-agnostic and consumed as a package:

- No required Office.js or Word API references in package core modules.
- Runtime defaults (author/platform) are caller-configurable.
- Node hosts use `@xmldom/xmldom`; browser hosts use native DOM APIs.

## Operational Guidance

1. Prefer OOXML-first implementations for document manipulation.
2. Keep Word API usage inside `src/taskpane/modules/docx-redline-js-integration/`.
3. Route reusable logic to `@ansonlai/docx-redline-js` rather than command-layer duplication.
4. Update `STATE.md` and `ROADMAP.md` whenever package boundaries or entrypoints change.
