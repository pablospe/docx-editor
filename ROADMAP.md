# Roadmap

docx-editor lets an LLM agent propose edits to a Word document as tracked changes (redlines) and comments, then lets a human adjudicate them in Word. This file is the single source of truth for what is next, what is deliberately not planned, and what shipped.

Item numbers (`#N`) are stable identifiers, not GitHub issue numbers. PR titles cite them as `type: what changed (ROADMAP.md #N)`. Current release: **0.8.1** (2026-08-30).

## Where the library stands

- **Editing**: `replace`/`delete`/`insert_*`/`rewrite_paragraph` on hash-anchored paragraph refs (`P3#a7b2`), batched atomically with dry-run validation; `\n` in edit text is a tracked paragraph split; a `note=` on any edit anchors a rationale comment on its revisions.
- **Revisions, three tiers**: revision < group (one edit) < changeset (one batch call). Groups and changesets survive save/reopen by reconstruction from author + collision-bumped `w:date`. There is no fourth tier.
- **Honest on foreign documents**: editing inside another author's insertion preserves their proposal; `list_revisions` and every resolve verb handle `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`/`w:pPrChange`, and never claim success while other revision types remain (`list_unhandled_revisions`, `ResolveResult.unhandled`).
- **Text model**: the paragraph text map covers `w:t` and `w:tab` (one atomic `\t`); text-box content is excluded; only `word/document.xml` is edited.
- **Verification**: ~1,800 tests; a 77-file, 15-producer corpus (24 of them Word-authored redlines) runs weekly in CI through open → edit → save → reopen, a LibreOffice re-save survival check, and a PDF render.
- **Interfaces**: Python API, `docx-session` CLI (persistent Jupyter kernel with a JSON `eval`), and `skills/docx/SKILL.md` as the agent-facing contract.

## Design commitments (constraints every future item inherits)

- Refs are session-scoped: a paragraph's hash may change between sessions and after edits to it; consumers re-find, they never persist refs.
- `\n` means "tracked paragraph split". Any future visible element mapped into the text map must therefore use a character that is not `\n` (see #6).
- Replacing text inside your own pending insertion amends it in place (no new revision, `group_id=None`); undo is rejecting the amended insertion's group.
- The collision-bumped `w:date` per changeset is the join key for return-leg reconciliation (#69). Do not change the stamping scheme.
- Resolve verbs (`accept_all`, `reject_all`, …) report what they could not resolve rather than returning a clean count.
- Write only what LibreOffice and Word both keep on re-save; the corpus gate exists because an unknown element is dropped silently.

## Next

Ordered by value. Each is one board task and one PR unless noted.

### 73. Decompose `track_changes.py`  [refactor, no behaviour change]

`docx_editor/track_changes.py` is 5,215 lines; `RevisionManager` alone is 107 methods over 3,739 lines. Every review round still finds real bugs in it, which says the file — not the reviewers — is the problem: each edit site re-derives the same run/insertion/boundary facts, and a fix in one site does not reach its siblings. The method list already clusters cleanly:

| cluster | representative methods | ~lines |
|---|---|---|
| public dataclasses | `EditOperation`, `EditResult`, `SearchResult`, `Revision`, `ResolveResult`, `UnhandledRevision` | 700 |
| group/changeset registry | `_reconstruct_groups`, `_grouped`, `_changeset`, `group_spans`, `groups_are_dead` | 500 |
| locate/search | `find_text`, `find_all`, `count_matches`, `_locate_*`, `_find_across_boundaries*` | 300 |
| batch + rewrite | `batch_edit`, `validate_batch`, `batch_rewrite`, `rewrite_paragraph`, `_rewrite_*` | 550 |
| replace sites | `_replace_same_context`, `_replace_within_own_ins`, `_replace_mixed_state`, `_replace_across_nodes` | 450 |
| delete sites | `_delete_*`, `_remove_from_insertion`, `_split_foreign_ins_at`, `_split_ins_after_child` | 550 |
| insert sites + paragraph split/rejoin | `insert_text_*`, `_insert_*`, `_split_*`, `_collect_tail_nodes`, `_rejoin_paragraph` | 550 |
| listing/parsing | `list_revisions`, `get_markup_text`, `_parse_revision`, `_revision_element_index` | 300 |
| resolution | `accept_*`/`reject_*`, `_resolve_*`, `_restore_*`, `_sweep_*` | 600 |

Plan: turn the module into a package `docx_editor/track_changes/` with one module per cluster; `RevisionManager` stays as the façade so the public API and `from docx_editor.track_changes import …` do not change. Rules: one cluster per PR, each PR a pure move (`git` must show renames, zero logic edits), gates unchanged — the test suite, the walk-count pins, and the corpus harness are the safety net that makes this mechanical. Do this **before** any further edit-site work (#6, #63c); it is what lets a fix land once instead of six times.

Progress and decisions (2026-08-30):
- **Steps 1–5 merged** (PRs #83, #84 registry, #85 locate, #87 batch/rewrite, #89 replace sites; each ~1 h approval-to-merge). Step 1 (PR #83): `track_changes.py` → package `track_changes/` with `models.py` (dataclasses, tag constants, validators), `dom.py` (minidom helpers), `diff.py` (tokenize/hunks/affix trimming), `manager.py` (`RevisionManager`, unchanged); `__init__.py` re-exports the public API. `scripts/check_pure_move.py --base main` is the gate for every step: line-multiset equality minus import plumbing, AST identity of every method, and `base.py` may hold only verbatim copies.
- **Mixins and `ty`**: a method in one mixin that touches an attribute or method defined elsewhere fails `ty`'s `unresolved-attribute`. Decision: a typed `_RevisionManagerBase` in `base.py` declaring the instance attributes and one verbatim-copied stub per cross-cluster callee (`raise NotImplementedError`) — not a package-wide `[[tool.ty.overrides]]`, which would turn every `self.*` into `Unknown`.
- **Remaining steps, one PR each, in order**: registry → locate → batch → replace → delete → insert (+ paragraph split/rejoin) → listing → resolution. Each moves its methods byte-identically into `<cluster>.py` as `class _<Cluster>Mixin(_RevisionManagerBase)` and adds the mixin to `RevisionManager`'s bases. Gate: pure-move script, ruff/format/ty, full suite (`-n 4`, 16 GB cap), CI; self-review only.

### 6. `w:br` in the paragraph text map  [second half of #6; `w:tab` shipped in 0.8.1]

A run-level line break is still invisible to search, hashes and edit placement. It cannot map to `\n` (that means paragraph split — see commitments); candidate is a distinct atomic character (e.g. `\u2028`, LINE SEPARATOR) with the same boundary-only edit rules as tabs. Refs of affected paragraphs change, as they did for tabs.

### 63c. Refuse edits that land inside a field result

Only `w:instrText` is known to the library today. A `replace()` targeting text inside a TOC entry, `REF` or page-number field applies, and Word silently discards it on the next field update — the same silent-loss family already eliminated elsewhere. Minimum: detect that a match falls inside `fldSimple`/`fldChar` run ranges and refuse with a teaching error; better: a read-side field inventory. Board task `f812f113` exists in the backlog.

### 72. Timing-sensitive session tests are flaky on CI

Two `tests/test_session.py` tests failed on single CI jobs in one night and passed on rerun, on commits that did not touch `session.py`: `test_main_full_lifecycle` (`main(["start", …])` returned 1 on Python 3.12, run 33277764141) and `test_stop_session_is_prompt` (`stop_session` took 5.02 s against a `< 3.0 s` assertion on Python 3.14, run 33298581269 — 5 s is exactly the `_kernel_alive` probe timeout, so the shutdown ack was missed under load). Every flake costs a 20–30 min rerun. Same family: `tests/test_session.py:138` waits with `time.sleep(1.0)` instead of an `execute_input` handshake. Worst case so far (run 33321498835, Python 3.14 shard 2, PR #87): after earlier session tests passed, every later one in the job — 41 tests — failed with `Kernel did not become ready within 30.0s`; the rerun passed. Something in that job left kernel start-up broken for the rest of the process (leaked kernel, port, or the conftest reaper); the fix needs a diagnostic in `start_session` that prints the kernel's stderr and the connection file state on timeout. Fix the tests' dependence on wall-clock under xdist load (a start budget that scales, a stop assertion on *behaviour* — ack received — not on seconds), and make a failed `start` print *why*.

### 75. Author-filtered bulk resolution must not touch other authors' revisions  [data loss; after #73 step 9]

`accept_all(author=A)` / `reject_all(author=A)` list revisions by author but resolve by `w:id`, and the id lookup has no author check. With B's `<w:ins w:id="7">` before A's `<w:del w:id="7">` (duplicate ids across authors occur in real files — the corpus has them from LibreOffice), `reject_all(author="A")` deletes B's inserted text and `accept_all(author="A")` makes it permanent, both reporting 2. The `accept_all` docstring documents this and `tests/test_track_changes.py::…:1143` pins it as expected — it is not acceptable for a data-loss path. Fix: resolve the exact elements selected by the listing (element identity), not their ids; flip the pinning test. Lives in the resolution cluster, so it lands after #73 moves it (found by CodeRabbit on PR #75; verified 2026-08-30).

### 79. Inserted text loses formatting after an empty split segment  [after #73 step 7]

`_apply_paragraph_splits` resets `fallback_rPr = ""` whenever the current paragraph has no runs; a split at a paragraph end yields an empty tail, so the *next* segment drops the surrounding `rPr`: on a bold paragraph `insert_after(…, "A\nC")` keeps bold on `C`, `"A\n\nC"` loses it. The comment above the loop already claims the propagation works — PR #72 merged the comment but not the one-line fix. No test covers it. Lives in the insert/split cluster, so it lands after #73 moves it.

### 81. Workspace creation bypasses the unpack symlink guard

#77 made `unpack_document` refuse a symlinked nearest-existing ancestor of `output_dir`, but `Document.open(path, workspace_dir=<symlink>)` still extracts through the link: `workspace.py` creates the workspace directory before calling unpack, so the guard sees an existing real directory. Applying the same rule at workspace creation would refuse `workspace_dir="/tmp/<new>"` on macOS (`/tmp → private/tmp`), so it needs the same nearest-existing-ancestor rule *and* a decision on whether a user-chosen `workspace_dir` under a system symlink is legitimate (it is on macOS; document or resolve it). Found by the multi-review of PR #88. Also: `openspec/specs/structured-errors/spec.md:49` still describes an unscoped `replace()` call that the runtime rejects.

### 69. Return-leg reconciliation: which changesets survived?  [design]

The product loop is agent proposes → human adjudicates in Word → document comes back; accepted revisions become plain text, rejected ones vanish, and there is no way to ask "which of my changesets survived?". A sent-vs-returned compare keyed on changeset `w:date` + content. Design first; large.

## Later (demand-gated — revisit when a named consumer asks)

- **74. Author moves.** The library never writes `w:moveFrom`/`w:moveTo`; a move through the API is a delete plus an insert (one changeset when done in one batch call), which Word shows as two revisions rather than one green move pair. Emitting a real move pair when a batch deletes text and inserts the identical text elsewhere is now feasible on #68's machinery (range marks, pairing by `w:name`). No consumer has asked.
- **29. Computed list numbers** ("7.2(a)") from `numbering.xml` — a real subsystem; only once list-context demand is proven.
- **30. Container parts** (headers/footers/footnotes/endnotes): multi-part refs (`H1:P2#hash`), per-part editors, tracked-change routing. Stage (a) read-only enumeration, (b) locations, (c) tracked edits. Zero requests across four dogfood rounds; the failure mode is a visible miss, not silent corruption.
- **MCP server** (PR #7, parked): persistent document caching behind an MCP interface; the `docx-session` kernel covers the same need today.

## Out of scope — deliberate refusals

- **Redaction.** Actively dangerous here: a tracked-changes library preserves history by design, so pseudo-redaction that leaves text in `w:del`, prior versions, or metadata is a liability generator. Permanent refusal.
- **Word automation suite**: table/image/style/section authoring, TOC and field authoring, content controls, mail merge, COM/live-Word integration. Cell *text* edits already work, which is the redline case; the rest enlarges the corruption surface of files shared with external parties.
- **Regex find/replace.** The consumer is an LLM that reads text and emits exact strings; regex adds span-vs-revision-context semantics and a footgun surface for zero demonstrated demand.
- **Two-arbitrary-document diff.** The domain version already exists — `list_revisions` markup view, `get_text`/`get_original_text`, rewrite diffing. The valuable cousin is #69.
- **Read-only schema validators.** The real contract is "opens in Word with zero repair prompts"; a schema check cannot predict that. The corpus gate is the honest version.
- **Text boxes as an editing surface** — only the exclusion fix (#65) shipped; box text is invisible, not half-editable.

## How this roadmap is set

CodeRabbit reviews every PR; its inline Critical/Major/Potential-issue comments are triaged before merge (an audit on 2026-08-30 of 261 unread comments produced #75–#80). Batches come from evidence, not ideas: four dogfooding rounds (consumer-persona agents driving the library from `SKILL.md` alone, adversarial error QA, token economics, scale), the corpus revision census (which types real producers emit), and a clean-room comparison against KitchenSink4Word (test *names* only — its code is PolyForm-NC and unusable). A finding becomes an item here, an item becomes a PR citing it, and a release closes a theme.

## Shipped

Newest first. Details live in each PR and in the release notes.

- **unreleased, on main**: #76 comment reply marker order, #77 unpack symlink ancestors, #78 session start leak, #80 docs/spec drift ([#88](https://github.com/pablospe/docx-editor/pull/88)); #73 steps 1–5 ([#83](https://github.com/pablospe/docx-editor/pull/83), [#84](https://github.com/pablospe/docx-editor/pull/84), [#85](https://github.com/pablospe/docx-editor/pull/85), [#87](https://github.com/pablospe/docx-editor/pull/87), [#89](https://github.com/pablospe/docx-editor/pull/89)); CI sharded, ~8½ min wall ([#86](https://github.com/pablospe/docx-editor/pull/86))
- **0.8.1** (2026-08-30): #68 Resolve moves and `w:pPrChange` as revisions ([#82](https://github.com/pablospe/docx-editor/pull/82)); #6a `w:tab` in the paragraph text map ([#81](https://github.com/pablospe/docx-editor/pull/81)); #71 LibreOffice opens-clean gate + real Word redline fixtures ([#80](https://github.com/pablospe/docx-editor/pull/80))
- **0.8.0** (2026-08-29): #65 Exclude w:txbxContent from the host paragraph text map ([#78](https://github.com/pablospe/docx-editor/pull/78)); #67 Rationale channel ([#79](https://github.com/pablospe/docx-editor/pull/79)); #66 + #70 settings.xml pair + anti-scope statement ([#77](https://github.com/pablospe/docx-editor/pull/77)); #64 Foreign-revision census + accept_all honesty floor ([#76](https://github.com/pablospe/docx-editor/pull/76))
- **0.7.2** (2026-08-28): #56+#62 Perf follow-ups + test hygiene ([#75](https://github.com/pablospe/docx-editor/pull/75)); #39 (also filed as #31) Site D own-insertion replace ordering ([#74](https://github.com/pablospe/docx-editor/pull/74)); #52+#60 Ergonomics grab-bag ([#73](https://github.com/pablospe/docx-editor/pull/73))
- **0.7.1** (2026-07-24): #59 SKILL round-4 notes ([#70](https://github.com/pablospe/docx-editor/pull/70)); #57 Accept-path performance ([#71](https://github.com/pablospe/docx-editor/pull/71)); #58+61 `\n` = tracked paragraph split ([#72](https://github.com/pablospe/docx-editor/pull/72))
- **0.7.0** (2026-07-23): #50 Session eval JSON + structured errors ([#62](https://github.com/pablospe/docx-editor/pull/62)); #48 Error-contract round 3 ([#63](https://github.com/pablospe/docx-editor/pull/63)); #46 Revision groups survive reopen ([#61](https://github.com/pablospe/docx-editor/pull/61)); #49 SKILL/docs round-3 sync ([#64](https://github.com/pablospe/docx-editor/pull/64)); #53 Collision-bumped w:date stamping ([#65](https://github.com/pablospe/docx-editor/pull/65)); #47 Formatting preservation + span trimming ([#66](https://github.com/pablospe/docx-editor/pull/66)); #51 batch_edit apply performance ([#67](https://github.com/pablospe/docx-editor/pull/67)); #54 Changeset tier ([#68](https://github.com/pablospe/docx-editor/pull/68)); #55 Docs/skills full audit + sync ([#69](https://github.com/pablospe/docx-editor/pull/69))
- **0.6.1** (2026-07-22): #44 SKILL.md round-2 sync + [create] extra ([#56](https://github.com/pablospe/docx-editor/pull/56)); #41 Open-path robustness ([#55](https://github.com/pablospe/docx-editor/pull/55)); #45 Session CLI ergonomics ([#58](https://github.com/pablospe/docx-editor/pull/58)); #40 Input-validation bugs ([#57](https://github.com/pablospe/docx-editor/pull/57)); #42 Error contract completion ([#59](https://github.com/pablospe/docx-editor/pull/59)); #43 LLM token ergonomics ([#60](https://github.com/pablospe/docx-editor/pull/60))
- **0.6.0** (2026-07-15): #31 rPr leak + render_wt dedup ([#47](https://github.com/pablospe/docx-editor/pull/47)); #36 Style-chain numbering ([#51](https://github.com/pablospe/docx-editor/pull/51)); #32 SKILL.md sync ([#48](https://github.com/pablospe/docx-editor/pull/48)); #38 Corpus harness ([#50](https://github.com/pablospe/docx-editor/pull/50)); #33 Revision location ([#52](https://github.com/pablospe/docx-editor/pull/52)); #35 Exception contract ([#49](https://github.com/pablospe/docx-editor/pull/49)); #34 find_all + occurrence ergonomics ([#53](https://github.com/pablospe/docx-editor/pull/53)); #37 Revision grouping ([#54](https://github.com/pablospe/docx-editor/pull/54)); #31 rPr correctness + render_wt dedup ([#47](https://github.com/pablospe/docx-editor/pull/47))
- **0.5.0** (2026-07-14): #18 UTF-8 unpack ([#38](https://github.com/pablospe/docx-editor/pull/38)); #26 List info on ParagraphLocation ([#41](https://github.com/pablospe/docx-editor/pull/41)); #27 Style/outline context on ParagraphLocation ([#42](https://github.com/pablospe/docx-editor/pull/42)); #22-24 Workspace hardening ([#44](https://github.com/pablospe/docx-editor/pull/44)); #28 Section index on ParagraphLocation ([#45](https://github.com/pablospe/docx-editor/pull/45)); #19 Author-aware in-place edits ([#43](https://github.com/pablospe/docx-editor/pull/43)); #20 Run-order preservation ([#46](https://github.com/pablospe/docx-editor/pull/46)); #25 + #7 API hygiene ([#40](https://github.com/pablospe/docx-editor/pull/40)); #21 Occurrence search unification ([#39](https://github.com/pablospe/docx-editor/pull/39)); #21 Occurrence drift in document-wide search ([#39](https://github.com/pablospe/docx-editor/pull/39)); #25 Typed EditOperation constructors ([#40](https://github.com/pablospe/docx-editor/pull/40)); #7 Internal types exported as public API ([#40](https://github.com/pablospe/docx-editor/pull/40)); #22 meta.json written non-atomically ([#44](https://github.com/pablospe/docx-editor/pull/44)); #23 mark_dirty() write-ahead contract violations ([#44](https://github.com/pablospe/docx-editor/pull/44)); #24 No cross-process protection ([#44](https://github.com/pablospe/docx-editor/pull/44)); #19 Editing inside another reviewer's insertion destroys their proposal ([#43](https://github.com/pablospe/docx-editor/pull/43))
- **0.4.0** (2026-07-14): #12 Workspace → cache directory ([#29](https://github.com/pablospe/docx-editor/pull/29)); #13 Foreign-revision test fixtures ([#32](https://github.com/pablospe/docx-editor/pull/32)); #14 Original text view ([#35](https://github.com/pablospe/docx-editor/pull/35)); #15 Fixed-point accept_all/reject_all ([#36](https://github.com/pablospe/docx-editor/pull/36)); #16 w:delText → w:t fallback ([#34](https://github.com/pablospe/docx-editor/pull/34)); #17 Stale workspace divergence ([#37](https://github.com/pablospe/docx-editor/pull/37))
- **0.3.2** (2026-07-02): #8 Structured paragraph output ([#26](https://github.com/pablospe/docx-editor/pull/26)); #9 Single-paragraph lookup ([#27](https://github.com/pablospe/docx-editor/pull/27)); #10 Dry-run validation ([#28](https://github.com/pablospe/docx-editor/pull/28))
- **0.3.1** (2026-06-08): #2 batch_edit() rollback ([#19](https://github.com/pablospe/docx-editor/pull/19)); #3 ZIP path traversal / symlink ([#20](https://github.com/pablospe/docx-editor/pull/20))
- **0.2.4** (2026-05-18): #5 Comment anchoring bug ([#21](https://github.com/pablospe/docx-editor/pull/21))
- **earlier**: #1 Stale workspace handling
