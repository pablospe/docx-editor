#!/usr/bin/env python3
"""Round-trip robustness harness for docx-editor over a real-world .docx corpus.

Modes:
  corpus_harness.py                 run all corpus files (parent mode)
  corpus_harness.py --only NAME     run a single file by name (parent mode, filtered)
  corpus_harness.py --no-soffice    skip the LibreOffice stages (lo_roundtrip, pdf);
                                    --no-pdf is accepted as an alias
  corpus_harness.py --census        print the revision census only (no library round-trip)
  corpus_harness.py --single PATH   internal: run stages for one file, JSON to stdout

Parent mode isolates each file in a subprocess with a hard timeout, so one
hang/crash cannot kill the run. Results land in results.json next to this
script, a summary table is printed, and the exit code is 1 if any file has
a real failure — or if no corpus files were found at all (CI signal).

Stages per file:
  input_validate  informational: does the *input* zip/XML parse cleanly?
  open            Document.open with a dedicated workspace_dir
  read            list_paragraphs + get_visible_text + list_revisions
  edit            tracked-replace of the first word of the first non-empty paragraph
  save1           save-as to out/<name>_edited.docx + zip/XML validation
  reopen          reopen saved file, assert the edit survived and the revision
                  count is coherent, then accept_all
  save2           save to out/<name>_final.docx + zip/XML validation, then
                  assert the saved word/document.xml holds only what
                  accept_all itself reported as unhandled: an ins/del/move/
                  pPrChange element or move range mark left behind unreported
                  is a failure
  lo_roundtrip    soffice --headless --convert-to docx on the *edited* output,
                  then assert that what we wrote survived the re-save
  pdf             soffice --headless --convert-to pdf on the final output

The two LibreOffice stages are the closest thing to "opens in Word without a
repair prompt" that runs on a CI box. Two facts about soffice shape them:
its exit code is 0 even when it refuses a file (the signal is an "Error:"
line and a missing output file), and it prints nothing at all for an element
it does not recognize — it silently drops it on re-save (ISSUES.md #66, PR #77).
So each stage fails on any Error:/Warning: line soffice prints and on a
missing output (a nonzero exit after a good output is also a failure, though
never the signal relied on), and lo_roundtrip additionally reopens the
re-saved file and checks that our w:trackRevisions flag and our own
insertion/deletion are still there. Where
LibreOffice's own model cannot hold our redline (inside a field result),
the manifest records a ``survival_waiver`` with the
reason; see apply_manifest_expectations.

An input that fails input_validate and is then refused by Document.open is
recorded as "rejected" (mark: r), not as a failure — refusing an invalid
document is correct library behavior.

Every file also gets a *census*: a count of revision-bearing elements by tag
across its ``word/*.xml`` parts, recorded as ``rec["census"]``. It is
informational and can never fail a run, so it is not a stage — it exists to
say which revision types real-world producers actually emit, which is the
evidence base for handling the ones this library does not resolve.
``--census`` prints just that table, in seconds.
"""

import argparse
import json
import os
import shutil
import signal
import subprocess
import sys
import time
import traceback
import zipfile
from dataclasses import dataclass
from pathlib import Path

HERE = Path(__file__).resolve().parent
FILES_DIR = HERE / "files"
OUT_DIR = HERE / "out"
WORK_DIR = HERE / "work"
PDF_DIR = OUT_DIR / "pdf"
LO_DIR = OUT_DIR / "lo"
LO_PROFILE = HERE / "loprofile"
RESULTS_PATH = HERE / "results.json"
PER_FILE_TIMEOUT = 60  # seconds, library stages
PDF_TIMEOUT = 60  # seconds, soffice pdf conversion
LO_TIMEOUT = 60  # seconds, soffice docx re-save
DRAIN_TIMEOUT = 10  # seconds, draining soffice's pipes after a timeout kill

# Lines soffice prints that say nothing about the document. oosplash prints
# one of two javaldx warnings on every run of a machine without a JRE (a
# plain CI runner); the token covers both.
SOFFICE_NOISE = ("javaldx",)

STAGES = ["input_validate", "open", "read", "edit", "save1", "reopen", "save2", "lo_roundtrip", "pdf"]


# --------------------------------------------------------------------------
# Validation helpers
# --------------------------------------------------------------------------


def validate_docx(path: Path) -> dict:
    """Validate that a .docx is a sound zip and every XML part parses.

    Uses defusedxml so entity/DTD tricks are surfaced rather than expanded.
    Returns {"ok": True} or {"ok": False, "error_type": ..., "error": ...}.
    """
    from defusedxml import minidom as safe_minidom

    try:
        with zipfile.ZipFile(path) as z:
            bad = z.testzip()
            if bad is not None:
                return {"ok": False, "error_type": "BadZipCRC", "error": f"CRC error in {bad}"}
            names = z.namelist()
            if "word/document.xml" not in names:
                return {
                    "ok": False,
                    "error_type": "MissingDocumentXml",
                    "error": "word/document.xml not in package",
                }
            for n in names:
                if n.endswith((".xml", ".rels")):
                    try:
                        safe_minidom.parseString(z.read(n))
                    except Exception as e:
                        return {
                            "ok": False,
                            "error_type": type(e).__name__,
                            "error": f"part {n}: {e}",
                        }
    except Exception as e:
        return {"ok": False, "error_type": type(e).__name__, "error": str(e)[:300]}
    return {"ok": True}


def err_record(e: Exception) -> dict:
    tb = traceback.extract_tb(e.__traceback__)
    frame = ""
    for fr in reversed(tb):
        if "docx_editor" in fr.filename:
            frame = f"{Path(fr.filename).name}:{fr.lineno} in {fr.name}"
            break
    return {
        "status": "fail",
        "error_type": type(e).__name__,
        "error": str(e)[:400],
        "lib_frame": frame,
    }


def assert_fail(error_type: str, error: str) -> dict:
    return {"status": "fail", "error_type": error_type, "error": error}


# --------------------------------------------------------------------------
# Revision census
# --------------------------------------------------------------------------


def census_file(path: Path) -> dict:
    """Count revision-bearing elements by tag across a .docx's word/*.xml parts.

    Informational only — a file this cannot read contributes
    ``{"error": ...}`` and nothing else; the census never fails a run.

    Returns ``{"by_tag": {...}, "ins_del_contexts": {...}, "parts": {name: ...}}``,
    where the top-level maps are the sum over parts and ``parts`` keeps the
    per-part breakdown (so a redline living only in a header is visible as
    such — the library reads word/document.xml alone, ISSUES.md #30).
    """
    from defusedxml import minidom as safe_minidom

    from docx_editor.track_changes import count_revision_elements

    by_tag: dict[str, int] = {}
    contexts: dict[str, int] = {}
    parts: dict[str, dict] = {}
    try:
        with zipfile.ZipFile(path) as z:
            names = [n for n in z.namelist() if n.startswith("word/") and n.endswith(".xml")]
            for n in sorted(names):
                try:
                    dom = safe_minidom.parseString(z.read(n))
                except Exception as e:
                    parts[n] = {"error": f"{type(e).__name__}: {e}"[:200]}
                    continue
                c = count_revision_elements(dom)
                if not c.by_tag:
                    continue
                parts[n] = {"by_tag": c.by_tag, "ins_del_contexts": c.ins_del_contexts}
                for tag, n_elems in c.by_tag.items():
                    by_tag[tag] = by_tag.get(tag, 0) + n_elems
                for parent, n_elems in c.ins_del_contexts.items():
                    contexts[parent] = contexts.get(parent, 0) + n_elems
    except Exception as e:
        return {"error": f"{type(e).__name__}: {e}"[:200]}
    return {"by_tag": by_tag, "ins_del_contexts": contexts, "parts": parts}


def print_census(results: list[dict]) -> None:
    """Print the corpus-wide revision census: which tags real producers emit."""
    from docx_editor.track_changes import UNHANDLED_REVISION_TAGS

    totals: dict[str, int] = {}
    files_with: dict[str, set[str]] = {}
    producers: dict[str, set[str]] = {}
    contexts: dict[str, int] = {}
    errors: list[str] = []
    for r in results:
        census = r.get("census")
        if not census:
            # No census key at all: the file's subprocess crashed or timed out
            # (see run_all). An uncounted file is not a file with no revisions,
            # which is the distinction this whole block exists to keep.
            errors.append(f"{r['file']}: census not run (harness stage did not complete)")
            continue
        if "error" in census:
            errors.append(f"{r['file']}: {census['error']}")
            continue
        # A part that will not parse (the must_reject entity file) is reported,
        # not skipped in silence — an uncounted part is not the same as a part
        # with no revisions, and this table's whole job is that distinction.
        for part, rec in census["parts"].items():
            if "error" in rec:
                errors.append(f"{r['file']} [{part}]: {rec['error']}")
        producer = (r.get("provenance") or {}).get("producer", "?")
        for tag, n in census["by_tag"].items():
            totals[tag] = totals.get(tag, 0) + n
            files_with.setdefault(tag, set()).add(r["file"])
            producers.setdefault(tag, set()).add(producer)
        for parent, n in census["ins_del_contexts"].items():
            contexts[parent] = contexts.get(parent, 0) + n

    n_files = len(results)
    carrying = sum(1 for r in results if (r.get("census") or {}).get("by_tag"))
    print("\nRevision census (all word/*.xml parts)")
    print(f"{'tag':<32}{'elements':>10}{'files':>7}  producers")
    for tag, n in sorted(totals.items(), key=lambda kv: (-kv[1], kv[0])):
        mark = "*" if tag in UNHANDLED_REVISION_TAGS else " "
        prods = ", ".join(sorted(producers[tag]))
        print(f"{mark}{tag:<31}{n:>10}{len(files_with[tag]):>7}  {prods[:60]}")
    if not totals:
        print("  (no revision elements found in any corpus file)")
    print(f"\n{carrying}/{n_files} files carry at least one revision element")
    unhandled_total = sum(n for tag, n in totals.items() if tag in UNHANDLED_REVISION_TAGS)
    print(f"* = not resolved by accept_all/reject_all ({unhandled_total} element(s))")
    if contexts:
        print("\nw:ins/w:del by parent element (structural markers vs content revisions):")
        for parent, n in sorted(contexts.items(), key=lambda kv: (-kv[1], kv[0])):
            note = {
                "w:rPr": "  <- paragraph-mark ins/del, or a change record's rPr",
                "w:trPr": "  <- table-row ins/del",
            }.get(parent, "")
            print(f"  {parent:<28}{n:>8}{note}")
    if errors:
        n_files = len({e.split(" [", 1)[0].rstrip(":") for e in errors})
        print(f"\n{len(errors)} XML part(s) across {n_files} file(s) could not be censused:")
        for e in errors[:6]:
            print(f"  - {e}")
        if len(errors) > 6:
            print(f"  ... and {len(errors) - 6} more")


def run_census_only(only: str | None) -> int:
    """``--census``: census every corpus file in-process, print the table.

    No subprocesses, no library round-trip, no soffice — reading and parsing
    the zips is all it takes, so the report is reproducible in seconds.
    """
    manifest: dict[str, dict] = json.loads((HERE / "manifest.json").read_text())
    files = sorted(FILES_DIR.glob("*.docx"))
    if only:
        files = [f for f in files if only in f.name]
    if not files:
        if only:
            print(f"no corpus files match --only {only!r}", file=sys.stderr)
        else:
            print(f"no corpus files in {FILES_DIR} — run build_corpus.py first", file=sys.stderr)
        return 1
    results = [
        {"file": f.name, "census": census_file(f), "provenance": manifest.get(f.name, {}), "stages": {}} for f in files
    ]
    print_census(results)
    return 0


# --------------------------------------------------------------------------
# LibreOffice stages
# --------------------------------------------------------------------------

AUTHOR = "CorpusHarness"
EDIT_MARKER = "-EDITED"
CHECKED = "checked"  # the lo_roundtrip record's survival/flag value when the assertion ran
NO_EDIT = "skipped (no edit)"  # its survival value when the edit stage was skipped

# The survival assertions a manifest ``survival_waiver`` may cover. Only the
# one where LibreOffice's model kept our text but not our revision: a dropped
# w:trackRevisions flag or a vanished edit marker is a loss, never a waiver.
WAIVABLE_ASSERTIONS = frozenset({"AssertOwnRevisionsDropped"})


@dataclass
class SofficeRun:
    """One ``soffice --convert-to`` invocation, read the only way it can be.

    ``returncode`` is recorded but is not the load signal: soffice exits 0
    after refusing a file. ``output`` is the file it actually wrote (None if
    none), ``messages`` the ``Error:``/``Warning:`` lines it printed minus
    ``SOFFICE_NOISE``, ``raw`` the combined output minus that noise, trimmed,
    for diagnostics. ``timed_out`` means the process tree was killed at the
    stage timeout; ``returncode`` is then the kill's and the rest is empty.
    """

    output: Path | None
    messages: list[str]
    returncode: int
    timed_out: bool
    raw: str


def soffice_noise(line: str) -> bool:
    return any(noise in line for noise in SOFFICE_NOISE)


def soffice_messages(text: str) -> list[str]:
    """The ``Error:``/``Warning:`` lines in soffice's output, minus known noise."""
    return [line for line in text.splitlines() if line.startswith(("Error:", "Warning:")) and not soffice_noise(line)]


def soffice_convert(soffice: str, src: Path, convert_to: str, outdir: Path, timeout: int) -> SofficeRun:
    """Convert ``src`` with LibreOffice into ``outdir``; ``convert_to`` is a soffice filter spec.

    ``soffice`` is a wrapper that forks ``soffice.bin``, so the conversion
    runs in its own session and a timeout kills that whole tree: an orphaned
    ``soffice.bin`` keeps the profile lock and stalls every later conversion.
    """
    ext = convert_to.split(":", 1)[0]
    output = outdir / f"{src.stem}.{ext}"
    output.unlink(missing_ok=True)  # a stale file from an earlier run must not pass
    cmd = [
        soffice,
        "--headless",
        # A real file URL: a space in the path hangs soffice on a bare file://{path}.
        f"-env:UserInstallation={LO_PROFILE.as_uri()}",
        "--convert-to",
        convert_to,
        "--outdir",
        str(outdir),
        str(src),
    ]
    # errors="replace": soffice's output is diagnostics, and a stray byte in
    # it must not raise past the timeout handling below.
    with subprocess.Popen(
        cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, errors="replace", start_new_session=True
    ) as proc:
        try:
            stdout, stderr = proc.communicate(timeout=timeout)
        except subprocess.TimeoutExpired:
            try:
                os.killpg(proc.pid, signal.SIGKILL)  # the session leader's pgid is its pid
            except OSError:
                proc.kill()  # no group to kill; the direct child at least must not be waited on alive
            try:
                # Drain the pipes; a descendant that left the session could still hold them.
                proc.communicate(timeout=DRAIN_TIMEOUT)
            except subprocess.TimeoutExpired:
                pass
            proc.wait()  # the direct child is SIGKILLed: reaping it cannot block
            return SofficeRun(output=None, messages=[], returncode=proc.returncode, timed_out=True, raw="")
    produced = output if output.exists() and output.stat().st_size > 0 else None
    combined = stdout + "\n" + stderr
    return SofficeRun(
        output=produced,
        messages=soffice_messages(combined),
        returncode=proc.returncode,
        timed_out=False,
        raw="\n".join(line for line in combined.splitlines() if line.strip() and not soffice_noise(line))[:400],
    )


def local_name(qname: str) -> str:
    return qname.rsplit(":", 1)[-1]


def track_revisions_on(path: Path) -> bool | None:
    """Is ``w:trackRevisions`` on in a .docx's settings.xml? None if the part is absent.

    A bare ``<w:trackRevisions/>`` is on (that is how Word writes it); an
    explicit ``val`` is read as ST_OnOff. Element and attribute are matched
    by local name, whatever prefix the producer used, as the library does
    when it writes the flag.
    """
    from defusedxml import minidom as safe_minidom

    with zipfile.ZipFile(path) as z:
        if "word/settings.xml" not in z.namelist():
            return None
        dom = safe_minidom.parseString(z.read("word/settings.xml"))
    for node in dom.documentElement.childNodes:
        if node.nodeType == node.ELEMENT_NODE and local_name(node.tagName) == "trackRevisions":
            val = next((v for k, v in node.attributes.items() if local_name(k) == "val"), None)
            return val is None or val.strip().lower() in ("true", "1", "on")
    return False


def survival_check(out1: Path, roundtrip: Path, work: Path, author: str) -> dict:
    """Did what the library wrote into ``out1`` survive a LibreOffice re-save?

    LibreOffice drops an element it does not recognize without a word — the
    ISSUES.md #66 defect (an unknown settings element, found in PR #77) was
    caught only by a hand round-trip. This is that round-trip, automated:
    the flag we wrote must still be on, and our own insertion + deletion
    must still be read back as revisions with our author. Existence checks
    only — LibreOffice may legally merge a deletion that spanned several
    runs, so counts are not compared. Returns a failure record, or a pass
    record whose ``flag`` says whether the flag check applied: when the
    edited file's flag is not on (a producer wrote ``w:val="false"`` and the
    library preserved it), or the file has no settings part at all (the
    library saves such a document without adding one), there is nothing for
    LibreOffice to drop.
    """
    flag_on = track_revisions_on(out1)
    if flag_on:
        flag = CHECKED
        if not track_revisions_on(roundtrip):
            return assert_fail(
                "AssertTrackRevisionsDropped",
                "w:trackRevisions was on in the edited file but is absent or off after the LibreOffice re-save",
            )
    elif flag_on is None:
        flag = "not applicable (no settings part in the edited file)"
    else:
        flag = "not applicable (not on in the edited file)"
    from docx_editor import Document

    doc = Document.open(roundtrip, author=author, workspace_dir=work, force_recreate=True)
    try:
        if EDIT_MARKER not in doc.get_visible_text():
            return assert_fail(
                "AssertEditMarkerLostInRoundtrip",
                "edit marker missing from visible text after the LibreOffice re-save",
            )
        kinds = {rev.type for rev in doc.list_revisions(author=author)}
        missing = [kind for kind in ("insertion", "deletion") if kind not in kinds]
        if missing:
            return assert_fail(
                "AssertOwnRevisionsDropped",
                f"no {' or '.join(missing)} by {author} after the LibreOffice re-save",
            )
    finally:
        doc.close(cleanup=True)
    return {"status": "pass", "survival": CHECKED, "flag": flag}


# Error types: ``Lo*`` names are LibreOffice-side failures — shared by both
# stages (LoMessage, LoNonzeroExit, SofficeSpawnFailed) or, for LoLoadRefused
# and LoOutputInvalid, raised only by lo_roundtrip; ``LoRoundtrip*`` and
# ``Pdf*`` names carry their stage explicitly.


def run_lo_roundtrip(out1: Path, edited: bool, soffice: str | None, work: Path) -> dict:
    """Stage lo_roundtrip: re-save the edited output through LibreOffice and check survival."""
    if soffice is None:
        return {"status": "skip", "reason": "soffice not found"}
    try:
        run = soffice_convert(soffice, out1, "docx:MS Word 2007 XML", LO_DIR, LO_TIMEOUT)
    except OSError as e:  # a soffice on PATH that cannot be started (half-removed install)
        return assert_fail("SofficeSpawnFailed", str(e)[:400])
    if run.timed_out:
        return assert_fail("LoRoundtripTimeout", "")
    if run.messages:
        return assert_fail("LoMessage", "; ".join(run.messages)[:400])
    if run.output is None:
        return assert_fail("LoLoadRefused", run.raw or "no docx produced")
    if run.returncode != 0:
        return assert_fail("LoNonzeroExit", nonzero_exit_message(run, "docx"))
    v = validate_docx(run.output)
    if not v["ok"]:
        # LibreOffice's output, not ours: named apart from save1/save2's OutputValidation.
        return assert_fail("LoOutputInvalid:" + v["error_type"], v["error"][:400])
    if not edited:
        return {"status": "pass", "survival": NO_EDIT}
    try:
        return survival_check(out1, run.output, work, AUTHOR)
    except Exception as e:
        return err_record(e)


def run_pdf(out2: Path, soffice: str | None) -> dict:
    """Stage pdf: render the final output with LibreOffice."""
    if soffice is None:
        return {"status": "skip", "reason": "soffice not found"}
    try:
        run = soffice_convert(soffice, out2, "pdf", PDF_DIR, PDF_TIMEOUT)
    except OSError as e:
        return assert_fail("SofficeSpawnFailed", str(e)[:400])
    if run.timed_out:
        return assert_fail("PdfConversionTimeout", "")
    if run.messages:
        return assert_fail("LoMessage", "; ".join(run.messages)[:400])
    if run.output is None:
        return assert_fail("PdfConversionFailed", run.raw or "no pdf produced")
    if run.returncode != 0:
        return assert_fail("LoNonzeroExit", nonzero_exit_message(run, "pdf"))
    return {"status": "pass"}


def nonzero_exit_message(run: SofficeRun, ext: str) -> str:
    """The exit code is this failure's only signal, so it is always in the message."""
    return (f"exit {run.returncode} after writing the {ext}" + (f": {run.raw}" if run.raw else ""))[:400]


# --------------------------------------------------------------------------
# Single-file mode (runs inside the timeout subprocess)
# --------------------------------------------------------------------------


def run_single(path: Path, do_soffice: bool) -> dict:
    name = path.name
    stages: dict[str, dict] = {}
    result = {"file": name, "stages": stages}

    def fail_rest(from_stage: str):
        idx = STAGES.index(from_stage)
        for s in STAGES[idx + 1 :]:
            stages[s] = {"status": "not_run"}

    # Stage 0: input validation (informational, never blocks)
    stages["input_validate"] = validate_docx(path)
    stages["input_validate"]["status"] = "pass" if stages["input_validate"]["ok"] else "fail"
    input_ok = stages["input_validate"]["ok"]

    # Census: not a stage (it has no pass/fail semantics), and taken before
    # open so a file the library refuses still contributes its tag counts.
    result["census"] = census_file(path)

    work = WORK_DIR / path.stem
    out1 = OUT_DIR / f"{path.stem}_edited.docx"
    out2 = OUT_DIR / f"{path.stem}_final.docx"

    from docx_editor import Document

    # Stage 1: open. An invalid input that the library refuses is "rejected",
    # not a failure — that is the correct behavior for such a document.
    doc = None
    try:
        doc = Document.open(path, author=AUTHOR, workspace_dir=work, force_recreate=True)
        stages["open"] = {"status": "pass"}
    except Exception as e:
        rec = err_record(e)
        if not input_ok:
            rec["status"] = "rejected"
        stages["open"] = rec
        fail_rest("open")
        return result

    try:
        # Stage 2: read
        try:
            paras = doc.list_paragraphs_structured(limit=None)
            visible = doc.get_visible_text()
            revs = doc.list_revisions()
            stages["read"] = {
                "status": "pass",
                "paragraphs": len(paras),
                "visible_chars": len(visible),
                "revisions": len(revs),
            }
        except Exception as e:
            stages["read"] = err_record(e)
            fail_rest("read")
            return result

        # Stage 3: edit (tracked replace of first word of first non-empty paragraph)
        target = next((p for p in paras if p.text.strip()), None)
        if target is None:
            stages["edit"] = {"status": "skip", "reason": "no non-empty paragraph"}
        else:
            word = target.text.split()[0]
            try:
                doc.replace(word, word + EDIT_MARKER, paragraph=target.ref, occurrence=0)
                # One logical replace yields >= 2 element-level revisions: one
                # w:del per source run the text spans, plus the w:ins
                # (grouping them is ISSUES.md #37). The exact count is recorded
                # so reopen can assert it survives the save/reopen round-trip.
                revisions_after_edit = len(doc.list_revisions())
                if revisions_after_edit < stages["read"]["revisions"] + 2:
                    stages["edit"] = assert_fail(
                        "AssertEditRevisionsMissing",
                        f"expected at least {stages['read']['revisions']} original + 2 "
                        f"revisions after tracked replace, got {revisions_after_edit}",
                    )
                    fail_rest("edit")
                    return result
                stages["edit"] = {
                    "status": "pass",
                    "word": word[:40],
                    "ref": target.ref,
                    "revisions_after_edit": revisions_after_edit,
                }
            except Exception as e:
                stages["edit"] = err_record(e)
                fail_rest("edit")
                return result

        # Stage 4: save-as + validate
        try:
            doc.save(out1)
            v = validate_docx(out1)
            if v["ok"]:
                stages["save1"] = {"status": "pass"}
            else:
                stages["save1"] = {
                    "status": "fail",
                    "error_type": "OutputValidation:" + v["error_type"],
                    "error": v["error"][:400],
                }
                fail_rest("save1")
                return result
        except Exception as e:
            stages["save1"] = err_record(e)
            fail_rest("save1")
            return result
    finally:
        try:
            doc.close(cleanup=True)
        except Exception:
            pass

    # Stage 5: reopen + deep assertions + accept_all
    edited = stages["edit"]["status"] == "pass"
    doc2 = None
    try:
        try:
            doc2 = Document.open(out1, author=AUTHOR, workspace_dir=work, force_recreate=True)
            if edited:
                if EDIT_MARKER not in doc2.get_visible_text():
                    stages["reopen"] = assert_fail(
                        "AssertEditMarkerLost",
                        "edit marker missing from visible text after reopen",
                    )
                    fail_rest("reopen")
                    return result
                reopen_revisions = len(doc2.list_revisions())
                if reopen_revisions != stages["edit"]["revisions_after_edit"]:
                    stages["reopen"] = assert_fail(
                        "AssertRevisionCountMismatch",
                        f"{stages['edit']['revisions_after_edit']} revisions before save, "
                        f"{reopen_revisions} after reopen",
                    )
                    fail_rest("reopen")
                    return result
            accepted = doc2.accept_all()
            if edited and accepted <= 0:
                stages["reopen"] = assert_fail(
                    "AssertNoAcceptedRevisions", "accept_all() accepted 0 revisions after an edit"
                )
                fail_rest("reopen")
                return result
            if edited and EDIT_MARKER not in doc2.get_visible_text():
                stages["reopen"] = assert_fail(
                    "AssertEditMarkerLostAfterAccept",
                    "edit marker missing from visible text after accept_all",
                )
                fail_rest("reopen")
                return result
            stages["reopen"] = {
                "status": "pass",
                "accepted": accepted,
                "unhandled": dict(accepted.unhandled_types),
            }
        except Exception as e:
            stages["reopen"] = err_record(e)
            fail_rest("reopen")
            return result

        # Stage 6: save2 + validate + post-condition: after accept_all, the
        # saved file may hold only revision types the library never resolves.
        try:
            from docx_editor.track_changes import MOVE_RANGE_TAGS, UNHANDLED_REVISION_TAGS

            doc2.save(out2)
            v = validate_docx(out2)
            if not v["ok"]:
                stages["save2"] = {
                    "status": "fail",
                    "error_type": "OutputValidation:" + v["error_type"],
                    "error": v["error"][:400],
                }
                fail_rest("save2")
                return result
            # word/document.xml only: the part accept_all reads (ISSUES.md #30).
            # A redline in styles.xml or footnotes.xml is visible in the
            # per-part census but is not something accept_all claimed to do.
            # Measured against accept_all's own report, not a static tag set:
            # a handled-type mark it could not reach (no numeric w:id) is
            # reported in unhandled_types and may stay; only what it left
            # behind *unreported* is a failure. Range marks are scaffolding
            # swept with their move, so they may remain only while a move
            # element of their family is still reported as unhandled.
            final_census = census_file(out2)
            body_tags = final_census.get("parts", {}).get("word/document.xml", {}).get("by_tag", {})
            reported = accepted.unhandled_types
            leftover = {
                tag: n - reported.get(tag, 0)
                for tag, n in body_tags.items()
                if tag not in UNHANDLED_REVISION_TAGS and tag not in MOVE_RANGE_TAGS and n > reported.get(tag, 0)
            }
            for tag, n in body_tags.items():
                if tag in MOVE_RANGE_TAGS and not reported.get("w:" + tag[len("w:") : tag.index("Range")], 0):
                    leftover[tag] = n
            if leftover:
                stages["save2"] = assert_fail(
                    "AssertResolvedTypesRemain",
                    f"accept_all() left unreported revision elements in the saved file: {leftover}",
                )
                fail_rest("save2")
                return result
            stages["save2"] = {"status": "pass", "census": final_census.get("by_tag", {})}
        except Exception as e:
            stages["save2"] = err_record(e)
            fail_rest("save2")
            return result
    finally:
        if doc2 is not None:
            try:
                doc2.close(cleanup=True)
            except Exception:
                pass

    # Stages 7-8: LibreOffice. lo_roundtrip re-saves the *edited* output (our
    # pending redline + the w:trackRevisions flag) and checks it survived;
    # pdf renders the *final* output. Both skip together.
    if not do_soffice:
        stages["lo_roundtrip"] = {"status": "skip", "reason": "--no-soffice"}
        stages["pdf"] = {"status": "skip", "reason": "--no-soffice"}
        return result
    soffice = shutil.which("soffice")
    # Independent artifacts: a dropped element in the edited file says
    # nothing about whether the final file renders, so pdf runs regardless.
    stages["lo_roundtrip"] = run_lo_roundtrip(out1, edited, soffice, WORK_DIR / f"{path.stem}_lo")
    stages["pdf"] = run_pdf(out2, soffice)
    return result


# --------------------------------------------------------------------------
# Parent mode
# --------------------------------------------------------------------------


def apply_manifest_expectations(rec: dict) -> None:
    """Enforce the per-file expectations the manifest records, in place.

    ``must_reject``: the library must refuse the file (e.g. external XML
    entities); accepting it is a failure. A genuine open failure already
    counts as failed and keeps its own diagnostics.

    ``survival_waiver``: a documented reason why LibreOffice cannot keep our
    redline in *this* file (a redline inside a field result, ...). The
    survival assertion is then reported as a skip with that
    reason — never as a pass — and a waiver that turns out to be unnecessary
    (everything survived) is itself a failure, so the manifest cannot quietly
    outlive the LibreOffice behavior it describes. Only ``WAIVABLE_ASSERTIONS``
    are waived: a refused load, an Error: line, a nonzero exit, a dropped
    w:trackRevisions flag, or a vanished edit marker still fails.
    """
    prov = rec["provenance"]
    stages = rec["stages"]
    if prov.get("must_reject") and stages.get("open", {}).get("status") == "pass":
        stages["open"] = {
            "status": "fail",
            "error_type": "MustRejectViolation",
            "error": "manifest marks this file must_reject, but Document.open accepted it",
        }
    waiver = prov.get("survival_waiver")
    lo = stages.get("lo_roundtrip", {})
    if waiver and lo.get("status") == "fail" and lo.get("error_type") in WAIVABLE_ASSERTIONS:
        stages["lo_roundtrip"] = {
            "status": "skip",
            "reason": f"survival waived: {waiver}",
            "dropped": lo["error_type"],
        }
    elif waiver and lo.get("status") == "pass" and lo.get("survival") == CHECKED:
        stages["lo_roundtrip"] = {
            "status": "fail",
            "error_type": "StaleSurvivalWaiver",
            "error": "manifest waives survival for this file, but everything survived the "
            "LibreOffice re-save — remove the waiver",
        }


def file_failed(rec: dict) -> bool:
    """A real failure: any harness error or failed stage other than input_validate.

    "rejected" (invalid input refused by the library) is not a failure.
    """
    return (
        any(st.get("status") == "fail" for s, st in rec["stages"].items() if s != "input_validate")
        or "harness" in rec["stages"]
    )


def make_dirs() -> None:
    for d in (OUT_DIR, PDF_DIR, LO_DIR, WORK_DIR):
        d.mkdir(exist_ok=True)


def run_all(only: str | None, do_soffice: bool) -> int:
    make_dirs()

    # Loaded unconditionally: must_reject enforcement lives in the manifest,
    # and a missing manifest must not silently disable it.
    manifest: dict[str, dict] = json.loads((HERE / "manifest.json").read_text())

    files = sorted(FILES_DIR.glob("*.docx"))
    if only:
        files = [f for f in files if only in f.name]
    if not files:
        if only:
            print(f"no corpus files match --only {only!r}", file=sys.stderr)
        else:
            print(f"no corpus files in {FILES_DIR} — run build_corpus.py first", file=sys.stderr)
        return 1

    # Budget every stage timeout plus the drain after each soffice kill, with
    # slack, so a child whose last stage timed out still reports that stage
    # instead of being killed mid-report.
    soffice_budget = LO_TIMEOUT + PDF_TIMEOUT + 2 * DRAIN_TIMEOUT if do_soffice else 0
    timeout = PER_FILE_TIMEOUT + soffice_budget + 10
    results = []
    for i, f in enumerate(files, 1):
        t0 = time.time()
        cmd = [sys.executable, str(Path(__file__).resolve()), "--single", str(f)]
        if not do_soffice:
            cmd.append("--no-soffice")
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
            if proc.returncode == 0 and proc.stdout.strip():
                rec = json.loads(proc.stdout)
            else:
                rec = {
                    "file": f.name,
                    "stages": {
                        "harness": {
                            "status": "fail",
                            "error_type": "SubprocessCrash",
                            "error": (proc.stderr or "")[-400:],
                        }
                    },
                }
        except subprocess.TimeoutExpired:
            rec = {
                "file": f.name,
                "stages": {
                    "harness": {
                        "status": "fail",
                        "error_type": "Timeout",
                        "error": f">{timeout}s",
                    }
                },
            }
        except Exception as e:
            rec = {
                "file": f.name,
                "stages": {"harness": {"status": "fail", "error_type": type(e).__name__, "error": str(e)}},
            }
        rec["duration_s"] = round(time.time() - t0, 1)
        rec["provenance"] = manifest.get(f.name, {})
        apply_manifest_expectations(rec)
        results.append(rec)
        # Rewrite results after every file so a killed run keeps partial
        # diagnostics; write-then-replace so a kill mid-write can't truncate it.
        tmp_results = RESULTS_PATH.with_suffix(".json.tmp")
        tmp_results.write_text(json.dumps(results, indent=2))
        tmp_results.replace(RESULTS_PATH)
        statuses = summarize_row(rec)
        print(f"[{i:2d}/{len(files)}] {f.name:50s} {statuses}", flush=True)

    print(f"\nresults written to {RESULTS_PATH}\n")
    print_summary(results)
    return sum(1 for r in results if file_failed(r))


def summarize_row(rec: dict) -> str:
    marks = {"pass": ".", "fail": "F", "skip": "s", "not_run": "-", "rejected": "r"}
    if "harness" in rec["stages"]:
        return "HARNESS-FAIL " + rec["stages"]["harness"].get("error_type", "")
    return " ".join(f"{s}:{marks.get(rec['stages'].get(s, {}).get('status', '?'), '?')}" for s in STAGES)


def print_summary(results: list[dict]) -> None:
    print(f"{'stage':<16}{'pass':>6}{'fail':>6}{'skip':>6}{'rejected':>10}{'not_run':>9}")
    for s in STAGES:
        counts = {"pass": 0, "fail": 0, "skip": 0, "rejected": 0, "not_run": 0}
        for r in results:
            st = r["stages"].get(s, {}).get("status")
            if st in counts:
                counts[st] += 1
        print(
            f"{s:<16}{counts['pass']:>6}{counts['fail']:>6}{counts['skip']:>6}"
            f"{counts['rejected']:>10}{counts['not_run']:>9}"
        )
    fails = {}
    for r in results:
        for s, rec in r["stages"].items():
            if rec.get("status") == "fail" and s != "input_validate":
                sig = f"{s}/{rec.get('error_type', '?')}"
                fails.setdefault(sig, []).append(r["file"])
    if fails:
        print("\nFailure signatures:")
        for sig, names in sorted(fails.items()):
            print(f"  {sig}: {len(names)} file(s)")
            for n in names[:6]:
                print(f"    - {n}")
    clean = sum(
        1
        for r in results
        if all(r["stages"].get(s, {}).get("status") in ("pass", "skip") for s in STAGES if s != "input_validate")
        and "harness" not in r["stages"]
    )
    rejected = sum(1 for r in results if r["stages"].get("open", {}).get("status") == "rejected")
    print(f"\n{clean}/{len(results)} files fully clean (all stages pass/skip)")
    if rejected:
        print(f"{rejected} rejected (invalid input refused by the library — not a failure)")
    print_survival_summary(results)
    print_census(results)


def print_survival_summary(results: list[dict]) -> None:
    """How many files the lo_roundtrip survival assertion actually ran on."""
    lo = [r["stages"].get("lo_roundtrip", {}) for r in results]
    checked = [s for s in lo if s.get("survival") == CHECKED]
    flag_na = sum(1 for s in checked if s.get("flag") != CHECKED)
    no_edit = sum(1 for s in lo if s.get("survival") == NO_EDIT)
    waived = sum(1 for s in lo if s.get("status") == "skip" and "dropped" in s)
    if checked or no_edit or waived:
        print(
            f"lo_roundtrip survival: {len(checked)} checked ({flag_na} with the flag not applicable), "
            f"{no_edit} skipped (no edit), {waived} waived"
        )


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--single", type=Path, help="internal: run one file, JSON to stdout")
    ap.add_argument("--only", help="filter corpus files by substring")
    ap.add_argument(
        "--no-soffice",
        "--no-pdf",
        dest="no_soffice",
        action="store_true",
        help="skip the LibreOffice stages (lo_roundtrip, pdf); --no-pdf is an alias",
    )
    ap.add_argument("--census", action="store_true", help="print the revision census only")
    args = ap.parse_args()

    if args.census:
        sys.exit(run_census_only(args.only))

    if args.single:
        make_dirs()
        print(json.dumps(run_single(args.single, do_soffice=not args.no_soffice)))
        return
    failures = run_all(args.only, do_soffice=not args.no_soffice)
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
