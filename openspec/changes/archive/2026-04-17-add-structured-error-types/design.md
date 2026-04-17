## Context

docx-editor is used as an MCP-style tool by LLM agents. When a tool call fails, the agent's only recovery signal is the exception class and message. `HashMismatchError` is the one error that already carries structured fields (`paragraph_index`, `expected_hash`, `actual_hash`, `paragraph_preview`) and encodes recovery instructions in its message. Agents using it can retry in-loop without re-reading the document.

Every other LLM-facing error today is a bare string:
- `TextNotFoundError` — no way to know which paragraph was searched or what it actually contains.
- Bad paragraph indices raise stdlib `IndexError` — no `index` or `total_paragraphs` attribute, so the agent must parse the message.
- Batch validation raises stdlib `ValueError` — on a 10-operation batch, the agent cannot extract which operation failed.

Phase 0B' experiments (hashline-anchored editing) showed the L1→L2 lift is +29pp and comes entirely from giving the agent the correct location. Pushing L3 diagnostics into the error type makes external feedback coaching redundant. The design principle: **diagnostics in the error object, not in the coach.**

## Goals / Non-Goals

**Goals:**
- Every LLM-facing error carries enough structured data for in-loop recovery without re-reading the document.
- Preserve `DocxEditError` as the single base class so catch-all handlers keep working.
- Error messages remain human-readable and self-describing — structured fields are additive, not a replacement.
- Consistent API: all structured errors expose their fields as instance attributes with stable names.

**Non-Goals:**
- Reworking non-LLM-facing errors (`DocumentNotFoundError`, `WorkspaceSyncError`, `XMLError`, etc.). They're raised at system boundaries, not in the edit loop.
- Adding a machine-readable error code enum. The exception class IS the code.
- Backward-compatible shims for removed `IndexError` / `ValueError` paths. Callers should catch `DocxEditError` or the specific subclass.
- Changing any behaviour other than the exceptions raised — no new validation, no new recovery paths, no retry logic.

## Decisions

**Decision: Each structured error is a plain `DocxEditError` subclass, not a dataclass.**
Rationale: `HashMismatchError` is already a regular subclass that sets instance attributes in `__init__`. Consistency beats novelty. Dataclasses would need `__init__` overrides anyway to compose the message string. Alternatives considered: `@dataclass(eq=False)` inheriting from `Exception` (works but offers nothing extra for this surface size).

**Decision: `TextNotFoundError` gains `paragraph_ref: str | None`, not a split class hierarchy.**
Rationale: The unscoped search path (no `paragraph=` arg) still raises `TextNotFoundError`. A `None` ref is honest and the message adapts. Alternatives considered: Separate `TextNotFoundInParagraphError` subclass — rejected because existing `except TextNotFoundError` handlers would need updating and the scoping information belongs in the data, not the type.

**Decision: `ParagraphIndexError` replaces stdlib `IndexError` on `_resolve_paragraph` — breaking change, no shim.**
Rationale: The old `IndexError` was raised from library code, not pass-through from list indexing. Callers using `except IndexError` were already relying on an implementation detail. The new class still inherits from `DocxEditError`, so `except DocxEditError` catches both old and new. Alternatives considered: Multiple inheritance (`class ParagraphIndexError(DocxEditError, IndexError)`) to keep backward compat — rejected because it muddles the class hierarchy and LLM consumers don't need `IndexError`.

**Decision: `BatchOperationError` replaces `ValueError` only on batch validation paths, not all `ValueError`s.**
Rationale: `ValueError` is also raised for non-batch invalid arguments (e.g., malformed ref strings). Those are developer errors, not LLM-in-loop recoverables. Scoping the breaking change to batch paths keeps the blast radius tight. Alternatives considered: Raise `BatchOperationError` from every `ValueError` site — rejected because it would reclassify unrelated errors.

**Decision: `TextNotFoundError` carries optional `occurrence: int | None` and `total_occurrences: int | None` fields.**
Rationale: The `_get_nth_match` path raises `TextNotFoundError` with a message like `"Only 3 occurrence(s)... but occurrence=5 requested"`. Without structured fields, the LLM must parse the count out of the string. Adding `occurrence` and `total_occurrences` lets the agent programmatically decide to retry with a lower occurrence number. Both are `None` when the error is not occurrence-related, keeping the API clean for the common case.

**Decision: `paragraph_preview` on `TextNotFoundError` MUST be truncated to 80 characters with `"..."` suffix.**
Rationale: `HashMismatchError` already caps `paragraph_preview` at 80 chars (`track_changes.py:93`). Applying the same rule to `TextNotFoundError` keeps the preview fields consistent across error types and avoids bloating error messages with multi-kilobyte paragraph text.

**Decision: `batch_edit` wraps `_apply_single_edit` calls with `try/except ValueError` → `BatchOperationError(i, str(e))`.**
Rationale: `_apply_single_edit` raises `ValueError` at 4 sites (missing required fields like `find`, `replace_with`, `text`, `content`) that fire exclusively from the batch path. The upfront validation in `batch_edit` (line 127) only catches structural issues; the per-operation field checks happen inside `_apply_single_edit`. Without the wrapper, these errors escape as bare `ValueError` with no `operation_index`. The wrapper converts them to `BatchOperationError` so the LLM knows which operation in the batch failed. Alternatives considered: modifying `_apply_single_edit` to raise `BatchOperationError` directly — rejected because `_apply_single_edit` shouldn't know about batching; the index is the caller's concern.

**Decision: Error messages include enough context to retry, even for callers that ignore the structured fields.**
Rationale: The LLM may not know about the structured fields yet. A good message like `"Paragraph index 999 out of range. Document has 12 paragraphs (1-indexed, valid: P1-P12)."` is itself L3 recovery info. Structured fields are the fast path; the message is the fallback.

**Decision: Documentation lives in `skills/docx/SKILL.md` as a structured-errors reference table.**
Rationale: That skill file is what Claude Code reads when using docx-editor. A table of error class → fields → recovery pattern is the shortest path to making agents use the structured fields.

**Decision: `BatchOperationError` carries only `operation_index` + `reason`, not the full `EditOperation` object.**
Rationale: The operation is reconstructible from the caller's input list by index. Embedding the full object would couple the error type to the operation schema and bloat serialized error messages. Revisit only if recovery tests prove the caller genuinely needs the full op.

**Decision: `ParagraphRef.parse()` keeps `ValueError` for malformed syntax; only `_resolve_paragraph` raises `ParagraphIndexError` for out-of-range indices.**
Rationale: Malformed syntax (e.g. `"not-a-ref"`) is a developer error, not an LLM-in-loop recoverable. The out-of-range check is the one the agent can act on (it knows the document size from the error). Keeping the distinction avoids reclassifying unrelated `ValueError`s.

## Risks / Trade-offs

- **Breaking `except IndexError` / `except ValueError` consumer code** → Mitigation: both new classes inherit from `DocxEditError`; call out in CHANGELOG/PR; bump minor version.
- **New error class proliferation if the pattern is applied too aggressively** → Mitigation: this change explicitly scopes to LLM-facing errors in the edit loop. Non-goal covers the rest.
- **Message format changes could break tests that match on strings** → Mitigation: keep existing substrings (`"Text not found"`, `"out of range"`) so loose `match=` patterns still work; tighten tests where we add new assertions.
- **`paragraph_preview` on `TextNotFoundError` could be truncated for very large paragraphs** → Mitigation: truncate to 80 characters with `"..."` suffix, matching `HashMismatchError`'s existing cap (see `track_changes.py:93`).

## Migration Plan

1. Implement Task 1 (TextNotFoundError fields) in isolation — purely additive, no breaking change yet. Ship.
2. Implement Tasks 2–3 (ParagraphIndexError, BatchOperationError) together — both breaking, one version bump.
3. Update `skills/docx/SKILL.md` and `docx_editor/__init__.py` exports.
4. Rollback: revert the commit range. Structured fields are additive on `TextNotFoundError`; for Tasks 2–3 the old `IndexError` / `ValueError` raise sites are recoverable from git history.

## Open Questions

None remaining — see Decisions above.
