# Atomic Save + Open-Document Guard Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. Use superpowers:test-driven-development within each task — every behavior change gets a failing test first.

**Goal:** Make `Document.save()` safe inside cloud-synced folders (Dropbox/OneDrive/SharePoint): the destination file is never observable in a half-written state, `validate=True` can no longer destroy the original document, and saving over a document that Word has open locally raises a structured error instead of setting up a silent-overwrite data loss.

**Architecture:** Two independent changes meeting in `save()`. (1) `pack_document()` stops writing the zip directly onto the destination; it writes to a temp file in the destination's parent directory, optionally validates *that*, and promotes it with `os.replace()` — atomic on the same volume on both POSIX and Windows. (2) A new owner-file check (`~$` stub detection) runs where the final destination is known, raising `DocumentOpenError` unless `force=True`. No new dependencies; no behavior change to unpacking, anchoring, or editing.

**Tech Stack:** Python ≥3.10, stdlib only (`zipfile`, `os`, `tempfile`, `pathlib`), pytest, uv, ruff.

**Background (why — from the 2026-07-04 cloud-sync research, verified against Microsoft docs and MS-FSSHTTP):**

- **The current write is not atomic.** `docx_editor/ooxml/pack.py:47` opens `zipfile.ZipFile(output_file, "w")` directly on the destination and streams members into it. A sync client, Word, or any pipeline sampling the file mid-write sees a truncated, invalid zip — and sync clients can propagate that state to other machines.
- **`validate=True` can destroy the document.** `pack.py:55` unlinks the freshly written output when LibreOffice validation fails — but in the default save-in-place case the output *is* the original, already overwritten. Validation failure currently leaves the user with **no file at all**. Validating the temp file before `os.replace()` fixes this as a side effect of atomicity.
- **Writing while Word has the document open loses data silently.** Word keeps its own in-memory copy and does not watch the disk; its next save overwrites whatever an external tool wrote. The `~$` owner file next to the document is the standard local signal, and docx-editor currently ignores it.
- **Known limitation (document, don't solve):** remote co-authoring (OneDrive/SharePoint/Word-for-the-web) leaves **no local filesystem signal** — no `~$` stub is created on this machine. Detecting it requires server APIs (Graph checkout / SharePoint `lockedByUser`), which is explicitly out of scope. The guard is best-effort for the *local* case and the docs must say exactly that.
- Sync-client behavior that motivates atomic replace: OneDrive and Dropbox both sync at block/file granularity (fixed blocks; the zip-container byte-shift defeats deltas anyway), so the destination should transition in a single rename, not a stream of partial writes.

## Global Constraints

- **Out of scope:** relocating the `.docx/<stem>/` unpack workspace (separate plan/PR); anything under the open MCP draft PR #7; server-side lock detection (Graph/SharePoint).
- **Coordination note:** the persistent-session plan (`2026-07-04-persistent-session-mode.md`, Task modifying `document.py` / `workspace.py`) also threads a `force` parameter through `Document.save()` to skip its *staleness* check. If both land, `force=True` must mean "skip all save-time guards" (staleness **and** open-document) — one parameter, unified semantics, documented as such. Whichever plan lands second reconciles.
- `requires-python = ">=3.10,<4.0"`; core `dependencies = ["defusedxml>=0.7.1"]` must NOT gain new entries.
- Line length 120 (ruff); run `uv run ruff check .` and `uv run ruff format .` before every commit.
- Run commands with `uv run ...` (`uv sync --dev` to install).
- Commits: conventional-commit style, never mention Claude/AI, never `git add -A` — always add specific files.
- Tests live in `tests/`, plain pytest style. `tests/conftest.py` provides `test_data_dir`, `simple_docx`, `temp_docx` (copy of simple.docx in a temp dir), `temp_dir`.
- The existing suite (~626 tests) is the oracle for zip output compatibility — it must stay green untouched.

---

## File Structure

- Modify: `docx_editor/ooxml/pack.py` — atomic emission + validate-before-replace (`pack_document`, lines 13–58)
- Modify: `docx_editor/workspace.py:168` — owner-file guard in `Workspace.save()` (destination is resolved here)
- Modify: `docx_editor/document.py:407` — `force: bool = False` parameter on `Document.save()`, passed through
- Modify: `docx_editor/exceptions.py` — new `DocumentOpenError(DocxEditError)`
- Create: `tests/test_atomic_save.py`
- Create: `tests/test_open_document_guard.py`
- Modify: `README.md`, `skills/docx/SKILL.md` — guard + `force` + remote-co-authoring caveat

---

### Task 1: Atomic pack + validate-before-replace

**Files:**
- Modify: `docx_editor/ooxml/pack.py`
- Create: `tests/test_atomic_save.py`

**Interfaces:**
- `pack_document(input_dir, output_file, validate=False) -> bool` — signature unchanged; behavior contract strengthened: at every instant, `output_file` contains either the complete previous document or the complete new one, never a partial zip; on any failure (exception or validation) the previous file is untouched and no temp file is left behind.

- [ ] **Step 1 (RED): failure mid-pack preserves the original.** Test: copy `simple_docx` to a temp dir, record its bytes; monkeypatch `zipfile.ZipFile.write` to raise `OSError` after the first member; call `pack_document` against that destination; assert the destination bytes are unchanged and no `*.tmp` (or chosen temp pattern) file remains in the directory.
- [ ] **Step 2 (RED): validation failure preserves the original.** Test: monkeypatch `validate_document` to return `False`; `pack_document(..., validate=True)` returns `False`; destination bytes unchanged; no temp litter. (This is the regression test for the current data-loss bug at `pack.py:55`.)
- [ ] **Step 3 (RED): success is a single-rename promotion.** Test: pack to a destination and assert the result is a valid zip (`zipfile.ZipFile(dest).testzip() is None`) and that the temp file is gone. Optionally assert rename semantics by monkeypatching `os.replace` to record its call (temp path in same parent dir as destination).
- [ ] **Step 4 (GREEN):** implement in `pack_document`: create the zip at `output_file.parent / f".{output_file.name}.<random>.tmp"` (same directory ⇒ same volume ⇒ atomic `os.replace`; leading dot + `.tmp` suffix keeps sync-adjacent tools from ingesting it); run `condense_xml` and member writes exactly as today; if `validate` and validation fails → unlink temp, return `False`; else `os.replace(temp, output_file)`. Wrap in `try/finally` so every failure path unlinks the temp.
- [ ] **Step 5:** run the FULL suite — zip output must be byte-compatible with whatever the existing tests assert (do not change compression, member order, or XML condensing).

---

### Task 2: Open-document guard (`~$` stub) with `force` override

**Files:**
- Modify: `docx_editor/exceptions.py`, `docx_editor/workspace.py:168`, `docx_editor/document.py:407`
- Create: `tests/test_open_document_guard.py`

**Interfaces:**
- New exception: `DocumentOpenError(DocxEditError)` — message names the stub path found and the recovery options ("close the document in Word, or pass force=True if this is a stale lock from a crashed session").
- `Document.save(path=None, validate=False, force=False)` — `force=True` skips the guard. Guard applies to the **destination** path (saving a copy to a fresh path while the source is open in Word is safe and must not be blocked).
- Helper (workspace-level or module function): `owner_file_candidates(path: Path) -> list[Path]`.

- [ ] **Step 1: verify Word's owner-file naming rule** against the Microsoft support doc ("The document is locked for editing by another user") before coding. Known shape: same directory, prefix `~$`, and for longer filenames the first characters of the stem are dropped (`Report.docx` → `~$port.docx`); short names keep the full stem (`ab.docx` → `~$ab.docx`). Implement `owner_file_candidates()` returning both `~$<name>` and `~$<name[2:]>` forms; existence check only (stubs are hidden system files on Windows — `Path.exists()` sees them regardless).
- [ ] **Step 2 (RED): guard blocks the save.** Test: `temp_docx` plus a fabricated `~$` stub (both naming variants, parametrized); `doc.replace(...)` then `doc.save()` raises `DocumentOpenError`; the destination file is unchanged.
- [ ] **Step 3 (RED): `force=True` saves.** Same setup; `doc.save(force=True)` succeeds and the edit is present on reopen.
- [ ] **Step 4 (RED): no false positives.** A stub for a *different* document in the same folder does not block; saving to a **fresh destination** path while the source has a stub does not block.
- [ ] **Step 5 (GREEN):** implement the check in `Workspace.save()` (where `destination` is resolved, line 168) before calling `pack_document`; thread `force` from `Document.save()`.
- [ ] **Step 6 (GREEN, Windows mapping):** catch `PermissionError` from the final `os.replace` in the pack/save path and re-raise as `DocumentOpenError` (Word genuinely holding the destination is exactly this condition; on POSIX this branch is effectively dead but harmless). Unit-test via monkeypatched `os.replace` raising `PermissionError`.
- [ ] **Step 7:** full suite + ruff green.

---

### Task 3: Documentation

**Files:**
- Modify: `README.md`, `skills/docx/SKILL.md`

- [ ] **Step 1:** README: short "Saving into synced folders" section — atomic save guarantee, the guard, `force=True` for stale stubs, and the explicit limitation: *remote co-authoring (OneDrive/SharePoint/Word-for-the-web) creates no local lock stub and cannot be detected from the filesystem; prefer editing when the document is not being actively collaborated on, and rely on cloud version history as the backstop.*
- [ ] **Step 2:** `skills/docx/SKILL.md`: teach the model the guard exists, that `DocumentOpenError` means "someone has this open locally — stop and tell the user," and that `force=True` is only for confirmed-stale stubs (crashed Word).

---

## Verification checklist (before PR)

- [ ] Full suite green (`uv run pytest`), ruff check + format clean.
- [ ] Manual probe: start a save, kill the process mid-pack (or simulate via the monkeypatch test) → original intact, no temp litter.
- [ ] Manual probe with real Word/LibreOffice if available: open a doc, attempt `save()` → `DocumentOpenError`; close app → save succeeds.
- [ ] PR description leads with the `validate=True` data-loss fix, then atomicity, then the guard; states the remote-co-authoring limitation; no AI attribution anywhere.
