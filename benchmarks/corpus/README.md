# Corpus round-trip harness

A robustness harness that runs docx-editor over a 77-file corpus of real-world
`.docx` files from 15 producer flavors (Word, LibreOffice, pandoc, ONLYOFFICE,
and the test suites of python-docx, mammoth, Apache POI, pandoc, LibreOffice
core, and Open-Xml-PowerTools).

Each file goes through the stages:

```
input_validate → open → read → edit → save1 → reopen → save2 → lo_roundtrip → pdf
```

`edit` performs a tracked replace of the first word and asserts it added
revisions (at least a del/ins pair; text spanning several runs legally yields
one `w:del` per run — see ISSUES.md #37); `reopen` asserts the edit marker
survived, that the revision count is unchanged by the save/reopen round-trip,
and that `accept_all()` accepts and keeps the edit. The last two stages hand
the outputs to LibreOffice, a real renderer — see below.

## LibreOffice stages

The contract this library cares about is "opens in Word with zero repair
prompts". A schema validator cannot predict that (deliberately out of scope);
a real renderer can approximate it, and LibreOffice is the one that runs on a
CI box. Two facts about `soffice --headless --convert-to` shape both stages,
both verified on LibreOffice 24.2 with `SAL_LOG=+WARN`:

- **Its exit code is 0 even when it refuses a file.** A truncated zip or a
  malformed `document.xml` prints `Error: source file could not be loaded`
  and writes nothing — the exit code says success. So each stage fails on any
  `Error:`/`Warning:` line soffice prints (any line mentioning `javaldx` is
  allow-listed noise: `oosplash` prints one of two such warnings on every run
  of a machine without a JRE) and on a missing output file. A nonzero exit
  code is also a failure — a crash after writing good output is worth
  knowing — but it is never the signal relied on.
- **It prints nothing for an element it does not recognize** — it drops it
  on re-save, silently. That is how ISSUES.md #66's first cut of the
  track-changes switch was caught (PR #77): a hand round-trip through
  LibreOffice, not any message.

So the stages are:

- `lo_roundtrip` re-saves the **edited** output (`out/<name>_edited.docx`:
  our pending redline plus the `w:trackRevisions` flag) as docx into
  `out/lo/`, validates the result as a zip/XML, and then reopens it with the
  library and asserts that what we wrote survived: the flag is still on when
  it was on before, the edit marker is still visible, and there is still an
  insertion and a deletion by our author (existence, not counts — LibreOffice
  may legally merge a deletion that spanned several runs). The survival part
  is skipped when `edit` was skipped, and the record's `flag` field says
  whether the flag check applied: a producer that wrote
  `<w:trackRevisions w:val="false"/>` (the library preserves it), or a
  document with no settings part at all (the library saves it without adding
  one), leaves nothing for LibreOffice to drop. The run summary prints how
  many files were actually checked, skipped, or waived. LibreOffice also
  rewrites or drops some *foreign* revision types on re-save; that is its
  behavior, not ours, and is not asserted on.
- `pdf` renders the **final** output (after `accept_all`) as PDF, the
  can-other-tools-read-it check.

Both stages skip together with `--no-soffice` or when `soffice` is not on
`PATH`. A conversion runs in its own session and is killed as a tree on
timeout: `soffice` is a wrapper that forks `soffice.bin`, and an orphaned
`soffice.bin` holds the profile lock and stalls every later conversion.
The pure helpers (`soffice_messages`, `track_revisions_on`,
`survival_check`, the stage function) are unit-tested in
`tests/test_corpus_harness.py` without LibreOffice — a fake `soffice`
stands in for the process-level tests — plus one real-`soffice` test that
skips where it is not installed.

## Running

```bash
make corpus-check                                    # assemble + full run (incl. LibreOffice stages)
uv run python benchmarks/corpus/build_corpus.py      # assemble corpus into files/
uv run python benchmarks/corpus/corpus_harness.py --no-soffice   # skip the LibreOffice stages (--no-pdf is an alias)
uv run python benchmarks/corpus/corpus_harness.py --only mammoth  # filter by substring
uv run python benchmarks/corpus/corpus_harness.py --census   # revision census only
```

`--census` reads and parses the zips in-process — no subprocesses, no library
round-trip, no `soffice` — so the census below is reproducible in seconds.

Each file runs in an isolated subprocess with a hard timeout; one hang or crash
cannot kill the run. Results are written to `results.json` and a summary table
is printed. Row marks: `.` pass, `F` fail, `s` skip, `r` rejected, `-` not run.

A weekly GitHub Actions workflow (`.github/workflows/corpus.yml`) runs the full
corpus — including the LibreOffice stages — with LibreOffice and pandoc installed;
trigger it manually with `workflow_dispatch` after changes.

## Revision census

Every run also counts revision-bearing elements by tag across each file's
`word/*.xml` parts (recorded as `rec["census"]` in `results.json`). It is
informational — never a stage, never a failure — and exists to answer which
revision types real-world producers actually emit. The `save2` stage adds the
one hard check derived from it: after `accept_all()`, the saved
`word/document.xml` may hold only what `accept_all()` itself reported as
unhandled — a `w:ins`, `w:moveFrom`, `w:pPrChange` or move range mark left
behind unreported fails the file with `AssertResolvedTypesRemain`. Observed
2026-08-30 over the 77-file corpus:

```text
tag                               elements  files  producers
 w:ins                                 162     10  LibreOffice (docx re-save), LibreOffice core ooxmlexport fix
 w:del                                 136     11  LibreOffice (docx re-save), Open-Xml-PowerTools test fixture
*w:tcPrChange                           36      6  Open-Xml-PowerTools test fixture (Word)
 w:moveFrom                             27      4  LibreOffice core ooxmlexport fixture (Word/LO mixed), Open-X
 w:moveTo                               27      4  LibreOffice core ooxmlexport fixture (Word/LO mixed), Open-X
 w:pPrChange                            11      5  LibreOffice core ooxmlexport fixture (Word/LO mixed), Open-X
*w:cellMerge                             6      1  Open-Xml-PowerTools test fixture (Word)
*w:tblGridChange                         6      6  Open-Xml-PowerTools test fixture (Word)
*w:cellDel                               4      1  Open-Xml-PowerTools test fixture (Word)
*w:cellIns                               4      1  Open-Xml-PowerTools test fixture (Word)
 w:moveFromRangeEnd                      4      4  LibreOffice core ooxmlexport fixture (Word/LO mixed), Open-X
 w:moveFromRangeStart                    4      4  LibreOffice core ooxmlexport fixture (Word/LO mixed), Open-X
 w:moveToRangeEnd                        4      4  LibreOffice core ooxmlexport fixture (Word/LO mixed), Open-X
 w:moveToRangeStart                      4      4  LibreOffice core ooxmlexport fixture (Word/LO mixed), Open-X
*w:tblPrChange                           4      4  Open-Xml-PowerTools test fixture (Word)
*w:rPrChange                             3      2  Open-Xml-PowerTools test fixture (Word)
*w:tblPrExChange                         3      2  Open-Xml-PowerTools test fixture (Word)
*w:trPrChange                            3      2  Open-Xml-PowerTools test fixture (Word)
*w:customXmlDelRangeEnd                  2      1  Open-Xml-PowerTools test fixture (Word)
*w:customXmlDelRangeStart                2      1  Open-Xml-PowerTools test fixture (Word)
*w:customXmlInsRangeEnd                  2      1  Open-Xml-PowerTools test fixture (Word)
*w:customXmlInsRangeStart                2      1  Open-Xml-PowerTools test fixture (Word)
*w:customXmlMoveFromRangeEnd             2      1  Open-Xml-PowerTools test fixture (Word)
*w:customXmlMoveFromRangeStart           2      1  Open-Xml-PowerTools test fixture (Word)
*w:customXmlMoveToRangeEnd               2      1  Open-Xml-PowerTools test fixture (Word)
*w:customXmlMoveToRangeStart             2      1  Open-Xml-PowerTools test fixture (Word)
*w:numberingChange                       2      1  Open-Xml-PowerTools test fixture (Word)
*w:sectPrChange                          1      1  Open-Xml-PowerTools test fixture (Word)

26/77 files carry at least one revision element
* = not resolved by accept_all/reject_all (88 element(s))

w:ins/w:del by parent element (structural markers vs content revisions):
  w:rPr                            172  <- paragraph-mark ins/del, or a change record's rPr
  w:p                               73
  w:trPr                            22  <- table-row ins/del
  w:moveFrom                        15
  w:moveTo                          15
  m:r                                1

1 XML part(s) across 1 file(s) could not be censused:
  - poi_ExternalEntityInText.docx [word/document.xml]: EntitiesForbidden: ...
```

Read as evidence:

- **Every unhandled type has genuine Word output behind it.** The 21
  `oxpt_RP*` files were authored in Word (see Provenance) and between them
  emit all 19 tags in `UNHANDLED_REVISION_TAGS` as well as the moves and
  `w:pPrChange` that ISSUES.md #68 took out of it. Before they were added,
  the five revision-bearing files were almost all LibreOffice-produced and
  only moves and `w:pPrChange` had any real-producer evidence; the rest was
  hand-authored XML in `tests/test_unhandled_revisions.py` (still there, for
  the edge shapes).
- **Moves resolve, on every producer that emits them** (ISSUES.md #68):
  `locore_TC-table-DnD-move.docx` (a Word drag-and-drop move of a whole table,
  re-exported by LibreOffice core) carries 20 move marks — `accept_all()`
  lists and resolves its 16 `w:moveFrom`/`w:moveTo` elements as one inferred
  changeset and sweeps the 4 range marks with them, leaving the table at its
  destination exactly once; `reject_all()` puts it back. The Word-native
  `oxpt_RP015-MoveFrom-MoveTo.docx` resolves clean (6 elements);
  `oxpt_RP018-MoveFrom-MoveTo-CC.docx` resolves its 7 move elements and leaves
  the 8 `customXmlMove*` marks around its content control pending. Half the
  DnD file's move elements are paragraph-mark markers
  (`w:pPr/w:rPr/w:moveFrom`, one per moved cell); like a deleted paragraph
  mark these resolve *approximately* — the marker is dropped without merging
  or splitting paragraphs — so the moved text ends up at its destination
  exactly once, but the source table survives as a 2×2 table of empty cells
  (the four leading empty lines in the fixture's expected visible text),
  where Word would remove it.
  Both shapes are mirrored by hand-authored fixtures in
  `tests/test_move_and_ppr_change_resolution.py` and run against the real
  files by `tests/test_corpus_fixtures.py` when the corpus is built.
- **Paragraph-property changes resolve too**: `w:pPrChange` in five files
  (Word: `oxpt_RP022`, `RP025`, `RP037`, the kitchen-sink `RP001`, and
  `locore_UnknownStyleInRedline.docx`, whose two records share `w:id="0"`, one
  self-closing, one naming a style `styles.xml` does not define — accept drops
  the records, reject restores the recorded properties verbatim). `RP037`
  also carries two `w:pPrChange` inside `word/styles.xml` style definitions,
  outside the part the library reads (ISSUES.md #30); they show in the
  per-part census and are not a `save2` failure, the same way
  `oxpt_RP050-Deleted-Footnote.docx`'s deletions in `word/footnotes.xml` are
  not.
- **The rest of the property family is still unhandled**: `w:rPrChange` on a
  run (`oxpt_RP037`) and on a paragraph mark (`oxpt_RP024`), `w:sectPrChange`
  (`oxpt_RP027`), `w:numberingChange` (`oxpt_RP026`), and the whole table
  family — `w:tblPrChange`, `w:tblPrExChange`, `w:trPrChange`,
  `w:tcPrChange`, `w:tblGridChange` (`oxpt_RP028`, `RP033`, `RP001`). These
  stay behind the honesty floor (`ResolveResult.unhandled`, ISSUES.md #64).
- **Table-structure revisions exist in the wild**: `w:cellIns` (`oxpt_RP035`),
  `w:cellDel` (`oxpt_RP034`), `w:cellMerge` (`oxpt_RP036`), each alongside the
  `w:ins`/`w:del` Word writes for the cells' content.
- **The structural `w:ins`/`w:del` contexts are the majority row**: 172
  paragraph-mark markers (`w:pPr/w:rPr/w:ins|del`; the kitchen-sink file alone
  has 160) and 22 table-row markers (`w:trPr`, from `oxpt_RP009`/`RP010` and
  `RP001`), against 73 ordinary content revisions. These resolve
  *approximately* today — the marker is dropped without merging the paragraph
  or removing the row — which is why the context breakdown is tracked
  separately from the unhandled count. The `m:r` row is a deletion inside a
  math run (`oxpt_RP013`), and the `w:moveFrom`/`w:moveTo` rows are Word's
  own paragraph-mark markers inside the kitchen-sink file's moved paragraphs.

The `*`-marked tags are `UNHANDLED_REVISION_TAGS`
(`docx_editor/track_changes.py`), which is what `accept_all()` /
`reject_all()` report as `.unhandled` and what `list_unhandled_revisions()`
lists — plus, in a document a nonconforming producer wrote, any handled-type
mark with no numeric `w:id` (nothing id-keyed can resolve it; none in this
corpus).

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
- A manifest entry can set `"survival_waiver": "<reason>"` when LibreOffice's
  own document model cannot hold our redline in that file. The
  `lo_roundtrip` survival assertion is then reported as a **skip** carrying
  the reason (never as a pass). Only `AssertOwnRevisionsDropped` — our text
  is there but no longer a revision — can be waived; a refused load, an
  `Error:` line, a nonzero exit, an unparseable output, a dropped
  `w:trackRevisions` flag, or a vanished edit marker still fails. A waiver
  that turns out to be unnecessary fails the file with
  `StaleSurvivalWaiver`, so the manifest cannot quietly outlive the behavior
  it documents. One file carries one today, verified by dumping LibreOffice's
  import as flat ODT:
  - `poi_FieldCodes.docx` — the first paragraph is an `AUTHOR` field result;
    Writer fields carry no redlines, so LibreOffice flattens our del/ins into
    the result text (`ANTONIANTONI-EDITED`).
  `locore_TC-table-DnD-move.docx` carried one until ISSUES.md #68: its first
  paragraphs sit inside a foreign `w:moveFrom` (moved-away text), which the
  library now treats as invisible, so the harness edits the first paragraph
  that is actually visible and the redline survives LibreOffice unaided.
- The harness exits nonzero if any file has a real failure (failed stage or
  harness error), and if the corpus directory is empty or a `--only` filter
  matches nothing (a run that tested nothing must not look green).
  Baseline: 76 clean + 1 rejected → exit 0. Of the 76, the `lo_roundtrip`
  survival assertion ran on 69 (1 of them with the flag check not
  applicable, `onlyoffice_sample.docx`), was skipped on 6 with no editable
  paragraph, and was waived on 1.

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
- The Word-authored redline fixtures (`oxpt_*`) are the RevisionProcessor
  test files of [Open-Xml-PowerTools](https://github.com/OpenXmlDev/Open-Xml-PowerTools)
  (MIT), pinned to commit `3891c2e5`. They were authored in Word 2013/2016
  (`docProps/app.xml` says so, and the revisions carry a human author), which
  is what makes them evidence about Word rather than about a converter. Two
  files from the same directory whose redlines were generated by
  WmlComparer (`w:author="Open-Xml-PowerTools"`) were left out for that
  reason. KitchenSink4Word is PolyForm-Noncommercial and is deliberately not
  used.

## Adding a file

1. Add a manifest entry: `kind`, `producer`, source (for downloads: a
   raw.githubusercontent.com URL pinned to a full commit SHA), `size`, and the
   truncated sha256 of the content — or add a generation recipe in
   `build_corpus.py` plus a source file in `srcgen/`. Add `"must_reject": true`
   for an intentionally invalid file the library must refuse, and
   `"survival_waiver": "<reason>"` if LibreOffice's own model cannot hold the
   redline `lo_roundtrip` writes into it (see Failure semantics).
2. Keep files ≤ 2MB.
3. Re-run `uv run python benchmarks/corpus/build_corpus.py`.
