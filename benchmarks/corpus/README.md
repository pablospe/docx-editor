# Corpus round-trip harness

A robustness harness that runs docx-editor over a 56-file corpus of real-world
`.docx` files from 14 producer flavors (Word, LibreOffice, pandoc, ONLYOFFICE,
and the test suites of python-docx, mammoth, Apache POI, pandoc, and
LibreOffice core).

Each file goes through the stages:

```
input_validate → open → read → edit → save1 → reopen → save2 → pdf
```

`edit` performs a tracked replace of the first word and asserts it added
revisions (at least a del/ins pair; text spanning several runs legally yields
one `w:del` per run — see ISSUES.md #37); `reopen` asserts the edit marker
survived, that the revision count is unchanged by the save/reopen round-trip,
and that `accept_all()` accepts and keeps the edit. `pdf` converts the final
output with LibreOffice as an external can-other-tools-read-it check.

## Running

```bash
make corpus-check                                    # assemble + full run (incl. PDF stage)
uv run python benchmarks/corpus/build_corpus.py      # assemble corpus into files/
uv run python benchmarks/corpus/corpus_harness.py --no-pdf   # skip the PDF stage
uv run python benchmarks/corpus/corpus_harness.py --only mammoth  # filter by substring
uv run python benchmarks/corpus/corpus_harness.py --census   # revision census only
```

`--census` reads and parses the zips in-process — no subprocesses, no library
round-trip, no `soffice` — so the census below is reproducible in seconds.

Each file runs in an isolated subprocess with a hard timeout; one hang or crash
cannot kill the run. Results are written to `results.json` and a summary table
is printed. Row marks: `.` pass, `F` fail, `s` skip, `r` rejected, `-` not run.

A weekly GitHub Actions workflow (`.github/workflows/corpus.yml`) runs the full
corpus — including the PDF stage — with LibreOffice and pandoc installed;
trigger it manually with `workflow_dispatch` after changes.

## Revision census

Every run also counts revision-bearing elements by tag across each file's
`word/*.xml` parts (recorded as `rec["census"]` in `results.json`). It is
informational — never a stage, never a failure — and exists to answer which
revision types real-world producers actually emit. Observed 2026-08-28 over the
56-file corpus:

```text
tag                               elements  files  producers
 w:ins                                   9      3  LibreOffice (docx re-save), LibreOffice core ooxmlexport fix
 w:del                                   8      2  LibreOffice (docx re-save), docx-editor test fixture
*w:moveFrom                              8      1  LibreOffice core ooxmlexport fixture (Word/LO mixed)
*w:moveTo                                8      1  LibreOffice core ooxmlexport fixture (Word/LO mixed)
*w:pPrChange                             2      1  LibreOffice core ooxmlexport fixture (Word/LO mixed)
*w:moveFromRangeEnd                      1      1  LibreOffice core ooxmlexport fixture (Word/LO mixed)
*w:moveFromRangeStart                    1      1  LibreOffice core ooxmlexport fixture (Word/LO mixed)
*w:moveToRangeEnd                        1      1  LibreOffice core ooxmlexport fixture (Word/LO mixed)
*w:moveToRangeStart                      1      1  LibreOffice core ooxmlexport fixture (Word/LO mixed)

5/56 files carry at least one revision element
* = not resolved by accept_all/reject_all (22 element(s), ISSUES.md #68)

w:ins/w:del by parent element (structural markers vs content revisions):
  w:p                               16
  w:rPr                              1  <- paragraph-mark ins/del, or a change record's rPr

1 XML part(s) across 1 file(s) could not be censused:
  - poi_ExternalEntityInText.docx [word/document.xml]: EntitiesForbidden: ...
```

Read as evidence for ISSUES.md #68:

- **Moves are the largest unhandled family**, and they are real, not synthetic:
  `locore_TC-table-DnD-move.docx` (a LibreOffice-core ooxmlexport fixture of a
  Word drag-and-drop move) carries all 20 move marks. `accept_all()` on it
  resolves 0 and leaves the whole redline pending.
- **Property changes occur too**: `locore_UnknownStyleInRedline.docx` carries
  2 `w:pPrChange` and likewise resolves to 0.
- **No corpus file uses** `w:rPrChange`, `w:sectPrChange`, `w:numberingChange`,
  any table-structure revision (`w:cellIns`/`w:cellDel`/`w:cellMerge`, the
  `*PrChange` family) or any custom-XML range mark. Reported as a gap rather
  than padded: the provenance policy requires corpus files to represent real
  producers, so hand-authored XML for those types lives in
  `tests/test_unhandled_revisions.py` instead.
- **The structural `w:ins`/`w:del` cases are rare but present**: one
  paragraph-mark marker (`w:pPr/w:rPr/w:ins`, in `locore_cell-sdt-redline.docx`
  — checked individually, since a change record's recorded `w:rPrChange/w:rPr`
  would share the `w:rPr` row) against 16 ordinary content revisions, and no
  `w:trPr` row markers at all.
  These resolve *approximately* today — the marker is dropped without merging
  the paragraph or removing the row — which is why the context breakdown is
  tracked separately from the unhandled count.

The `*`-marked tags are exactly `UNHANDLED_REVISION_TAGS`
(`docx_editor/track_changes.py`), which is also what `accept_all()` /
`reject_all()` report as `.unhandled` and what `list_unhandled_revisions()`
lists.

## Failure semantics

- An invalid input that fails `input_validate` and is then refused by
  `Document.open` is **rejected** (`r`), not a failure — refusing a broken
  document is correct library behavior (e.g. `poi_ExternalEntityInText.docx`,
  which contains external XML entities).
- A manifest entry can set `"must_reject": true`: the library **must** refuse
  that file, and a run where it is accepted fails with `MustRejectViolation`.
  This pins the rejection of `poi_ExternalEntityInText.docx`, so a parser
  regression that starts expanding entities cannot slip through as "one more
  passing file".
- The harness exits nonzero if any file has a real failure (failed stage or
  harness error), and if the corpus directory is empty or a `--only` filter
  matches nothing (a run that tested nothing must not look green).
  Baseline: 55 clean + 1 rejected → exit 0.

## Provenance policy

- **No `.docx` file is ever committed to this repo** (upstream licensing + repo
  size). `files/`, `out/`, `work/`, and `results.json` are gitignored.
- Corpus files are fetched to the developer's machine or CI runner at build
  time and never redistributed.
- `manifest.json` is the single source of truth and provenance record. Every
  entry records its `kind`, producer, source, size, and truncated sha256:
  - `local` — copied from this repo's `tests/test_data/` fixtures.
  - `download` — fetched from a URL pinned to a full upstream commit SHA and
    verified against the recorded sha256 at fetch time (mismatch = failure).
  - `generated` — produced locally by LibreOffice (`soffice`) or pandoc from
    the text sources in `srcgen/` or from local fixtures (recipes live in
    `build_corpus.py`; `srcgen/plain.odt` is an uncommitted intermediate
    generated from `plain.txt`). Sizes/hashes are informational — output bytes
    vary by tool version. If a tool is missing the entries are skipped with a
    notice; the corpus still works, just smaller.

## Adding a file

1. Add a manifest entry: `kind`, `producer`, source (for downloads: a
   raw.githubusercontent.com URL pinned to a full commit SHA), `size`, and the
   truncated sha256 of the content — or add a generation recipe in
   `build_corpus.py` plus a source file in `srcgen/`. Add `"must_reject": true`
   for an intentionally invalid file the library must refuse.
2. Keep files ≤ 2MB.
3. Re-run `uv run python benchmarks/corpus/build_corpus.py`.
