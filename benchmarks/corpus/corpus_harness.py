#!/usr/bin/env python3
"""Round-trip robustness harness for docx-editor over a real-world .docx corpus.

Modes:
  corpus_harness.py                 run all corpus files (parent mode)
  corpus_harness.py --only NAME     run a single file by name (parent mode, filtered)
  corpus_harness.py --no-pdf        skip the LibreOffice PDF-conversion stage
  corpus_harness.py --single PATH   internal: run stages for one file, JSON to stdout

Parent mode isolates each file in a subprocess with a hard timeout, so one
hang/crash cannot kill the run. Results land in results.json next to this
script, a summary table is printed, and the exit code is 1 if any file has
a real failure (CI signal).

Stages per file:
  input_validate  informational: does the *input* zip/XML parse cleanly?
  open            Document.open with a dedicated workspace_dir
  read            list_paragraphs + get_visible_text + list_revisions
  edit            tracked-replace of the first word of the first non-empty paragraph
  save1           save-as to out/<name>_edited.docx + zip/XML validation
  reopen          reopen saved file, assert the edit survived and the revision
                  count is coherent, then accept_all
  save2           save to out/<name>_final.docx + zip/XML validation
  pdf             soffice --headless --convert-to pdf on the final output

An input that fails input_validate and is then refused by Document.open is
recorded as "rejected" (mark: r), not as a failure — refusing an invalid
document is correct library behavior.
"""

import argparse
import json
import shutil
import subprocess
import sys
import time
import traceback
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
FILES_DIR = HERE / "files"
OUT_DIR = HERE / "out"
WORK_DIR = HERE / "work"
PDF_DIR = OUT_DIR / "pdf"
RESULTS_PATH = HERE / "results.json"
PER_FILE_TIMEOUT = 60  # seconds, library stages
PDF_TIMEOUT = 60  # seconds, soffice conversion

STAGES = ["input_validate", "open", "read", "edit", "save1", "reopen", "save2", "pdf"]


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
                    except Exception as e:  # noqa: BLE001
                        return {
                            "ok": False,
                            "error_type": type(e).__name__,
                            "error": f"part {n}: {e}",
                        }
    except Exception as e:  # noqa: BLE001
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
# Single-file mode (runs inside the timeout subprocess)
# --------------------------------------------------------------------------


def run_single(path: Path, do_pdf: bool) -> dict:
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

    work = WORK_DIR / path.stem
    out1 = OUT_DIR / f"{path.stem}_edited.docx"
    out2 = OUT_DIR / f"{path.stem}_final.docx"

    from docx_editor import Document

    # Stage 1: open. An invalid input that the library refuses is "rejected",
    # not a failure — that is the correct behavior for such a document.
    doc = None
    try:
        doc = Document.open(path, author="CorpusHarness", workspace_dir=work, force_recreate=True)
        stages["open"] = {"status": "pass"}
    except Exception as e:  # noqa: BLE001
        rec = err_record(e)
        if not input_ok:
            rec["status"] = "rejected"
        stages["open"] = rec
        fail_rest("open")
        return result

    try:
        # Stage 2: read
        try:
            paras = doc.list_paragraphs_structured()
            visible = doc.get_visible_text()
            revs = doc.list_revisions()
            stages["read"] = {
                "status": "pass",
                "paragraphs": len(paras),
                "visible_chars": len(visible),
                "revisions": len(revs),
            }
        except Exception as e:  # noqa: BLE001
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
                doc.replace(word, word + "-EDITED", paragraph=target.ref, occurrence=0)
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
            except Exception as e:  # noqa: BLE001
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
        except Exception as e:  # noqa: BLE001
            stages["save1"] = err_record(e)
            fail_rest("save1")
            return result
    finally:
        try:
            doc.close(cleanup=True)
        except Exception:  # noqa: BLE001
            pass

    # Stage 5: reopen + deep assertions + accept_all
    edited = stages["edit"]["status"] == "pass"
    doc2 = None
    try:
        try:
            doc2 = Document.open(out1, author="CorpusHarness", workspace_dir=work, force_recreate=True)
            if edited:
                if "-EDITED" not in doc2.get_visible_text():
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
            if edited and "-EDITED" not in doc2.get_visible_text():
                stages["reopen"] = assert_fail(
                    "AssertEditMarkerLostAfterAccept",
                    "edit marker missing from visible text after accept_all",
                )
                fail_rest("reopen")
                return result
            stages["reopen"] = {"status": "pass", "accepted": accepted}
        except Exception as e:  # noqa: BLE001
            stages["reopen"] = err_record(e)
            fail_rest("reopen")
            return result

        # Stage 6: save2 + validate
        try:
            doc2.save(out2)
            v = validate_docx(out2)
            if v["ok"]:
                stages["save2"] = {"status": "pass"}
            else:
                stages["save2"] = {
                    "status": "fail",
                    "error_type": "OutputValidation:" + v["error_type"],
                    "error": v["error"][:400],
                }
                fail_rest("save2")
                return result
        except Exception as e:  # noqa: BLE001
            stages["save2"] = err_record(e)
            fail_rest("save2")
            return result
    finally:
        if doc2 is not None:
            try:
                doc2.close(cleanup=True)
            except Exception:  # noqa: BLE001
                pass

    # Stage 7: PDF conversion of the final output
    if not do_pdf:
        stages["pdf"] = {"status": "skip", "reason": "--no-pdf"}
        return result
    soffice = shutil.which("soffice")
    if soffice is None:
        stages["pdf"] = {"status": "skip", "reason": "soffice not found"}
        return result
    profile = HERE / "loprofile"
    pdf_path = PDF_DIR / f"{out2.stem}.pdf"
    pdf_path.unlink(missing_ok=True)
    try:
        proc = subprocess.run(
            [
                soffice,
                "--headless",
                f"-env:UserInstallation=file://{profile}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(PDF_DIR),
                str(out2),
            ],
            capture_output=True,
            text=True,
            timeout=PDF_TIMEOUT,
        )
        if proc.returncode == 0 and pdf_path.exists() and pdf_path.stat().st_size > 0:
            stages["pdf"] = {"status": "pass"}
        else:
            stages["pdf"] = {
                "status": "fail",
                "error_type": "PdfConversionFailed",
                "error": (proc.stderr or proc.stdout or "no pdf produced")[:400],
            }
    except subprocess.TimeoutExpired:
        stages["pdf"] = {"status": "fail", "error_type": "PdfConversionTimeout", "error": ""}
    return result


# --------------------------------------------------------------------------
# Parent mode
# --------------------------------------------------------------------------


def file_failed(rec: dict) -> bool:
    """A real failure: any harness error or failed stage other than input_validate.

    "rejected" (invalid input refused by the library) is not a failure.
    """
    return (
        any(st.get("status") == "fail" for s, st in rec["stages"].items() if s != "input_validate")
        or "harness" in rec["stages"]
    )


def run_all(only: str | None, do_pdf: bool) -> int:
    OUT_DIR.mkdir(exist_ok=True)
    PDF_DIR.mkdir(exist_ok=True)
    WORK_DIR.mkdir(exist_ok=True)

    manifest = {}
    manifest_path = HERE / "manifest.json"
    if manifest_path.exists():
        manifest = json.loads(manifest_path.read_text())

    files = sorted(FILES_DIR.glob("*.docx"))
    if only:
        files = [f for f in files if only in f.name]

    results = []
    for i, f in enumerate(files, 1):
        t0 = time.time()
        cmd = [sys.executable, str(Path(__file__).resolve()), "--single", str(f)]
        if not do_pdf:
            cmd.append("--no-pdf")
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=PER_FILE_TIMEOUT + PDF_TIMEOUT)
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
                        "error": f">{PER_FILE_TIMEOUT + PDF_TIMEOUT}s",
                    }
                },
            }
        except Exception as e:  # noqa: BLE001
            rec = {
                "file": f.name,
                "stages": {"harness": {"status": "fail", "error_type": type(e).__name__, "error": str(e)}},
            }
        rec["duration_s"] = round(time.time() - t0, 1)
        rec["provenance"] = manifest.get(f.name, {})
        results.append(rec)
        statuses = summarize_row(rec)
        print(f"[{i:2d}/{len(files)}] {f.name:50s} {statuses}", flush=True)

    RESULTS_PATH.write_text(json.dumps(results, indent=2))
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


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--single", type=Path, help="internal: run one file, JSON to stdout")
    ap.add_argument("--only", help="filter corpus files by substring")
    ap.add_argument("--no-pdf", action="store_true", help="skip soffice PDF stage")
    args = ap.parse_args()

    if args.single:
        OUT_DIR.mkdir(exist_ok=True)
        PDF_DIR.mkdir(exist_ok=True)
        WORK_DIR.mkdir(exist_ok=True)
        print(json.dumps(run_single(args.single, do_pdf=not args.no_pdf)))
        return
    failures = run_all(args.only, do_pdf=not args.no_pdf)
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
