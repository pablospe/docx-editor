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
```

Each file runs in an isolated subprocess with a hard timeout; one hang or crash
cannot kill the run. Results are written to `results.json` and a summary table
is printed. Row marks: `.` pass, `F` fail, `s` skip, `r` rejected, `-` not run.

A weekly GitHub Actions workflow (`.github/workflows/corpus.yml`) runs the full
corpus with LibreOffice and pandoc installed; trigger it manually with
`workflow_dispatch` after changes.

## Failure semantics

- An invalid input that fails `input_validate` and is then refused by
  `Document.open` is **rejected** (`r`), not a failure — refusing a broken
  document is correct library behavior (e.g. `poi_ExternalEntityInText.docx`,
  which contains external XML entities).
- The harness exits nonzero if any file has a real failure (failed stage or
  harness error). Baseline: 55 clean + 1 rejected → exit 0.

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
   `build_corpus.py` plus a source file in `srcgen/`.
2. Keep files ≤ 2MB.
3. Re-run `uv run python benchmarks/corpus/build_corpus.py`.
