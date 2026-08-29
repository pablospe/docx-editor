"""Tests for the LibreOffice gate of the corpus harness (benchmarks/corpus/corpus_harness.py).

The harness is a script, not a package, so it is loaded by path. Everything
here runs without LibreOffice — a fake ``soffice`` script stands in for the
process-level tests — except the last test, which drives the real ``soffice``
when one is installed and is skipped otherwise.
"""

import importlib.util
import os
import re
import shutil
import sys
import time
import zipfile
from pathlib import Path

import pytest
from conftest import replace_docx_parts

from docx_editor import Document

REPO = Path(__file__).resolve().parents[1]
HARNESS_PATH = REPO / "benchmarks" / "corpus" / "corpus_harness.py"
SIMPLE = REPO / "tests" / "test_data" / "simple.docx"

JAVALDX = "Warning: failed to launch javaldx - java may not function correctly"


def _load_harness():
    spec = importlib.util.spec_from_file_location("corpus_harness", HARNESS_PATH)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


harness = _load_harness()


def rewrite_part(src: Path, dst: Path, part: str, transform) -> Path:
    """Copy ``src`` to ``dst`` with one zip part rewritten (``transform(text) -> text``)."""
    with zipfile.ZipFile(src) as z:
        text = z.read(part).decode("utf-8")
    replace_docx_parts(src, dst, {part: transform(text)})
    return dst


FAKE_SOFFICE = f"""#!{sys.executable}
# Stands in for soffice: records argv, prints the javaldx noise, then either
# converts (copies src to outdir with the target extension), refuses, or
# hangs with a grandchild that heartbeats — like oosplash forking soffice.bin.
import os, pathlib, shutil, subprocess, sys
pathlib.Path(os.environ["FAKE_SOFFICE_ARGV"]).write_text("\\n".join(sys.argv[1:]))
print("Warning: failed to launch javaldx - java may not function correctly", file=sys.stderr, flush=True)
mode = os.environ.get("FAKE_SOFFICE_MODE", "convert")
if mode == "hang":
    beat = os.environ["FAKE_SOFFICE_HEARTBEAT"]
    code = (
        "import time, pathlib\\n"
        "for _ in range(1200):\\n"  # ~60s bound: a leaked grandchild cannot spin forever if the group kill regresses
        f"    pathlib.Path({{beat!r}}).write_text(str(time.time()))\\n"
        "    time.sleep(0.05)"
    )
    subprocess.Popen([sys.executable, "-c", code]).wait()
elif mode == "refuse":
    print("Error: source file could not be loaded", file=sys.stderr)
else:
    args = sys.argv[1:]
    src = pathlib.Path(args[-1])
    outdir = pathlib.Path(args[args.index("--outdir") + 1])
    ext = args[args.index("--convert-to") + 1].split(":")[0]
    shutil.copy(src, outdir / f"{{src.stem}}.{{ext}}")
    print(f"convert {{src}} -> {{outdir}} using filter : fake")
"""


@pytest.fixture
def fake_soffice(tmp_path: Path, monkeypatch) -> Path:
    if os.name != "posix":
        pytest.skip("the fake is executed via a shebang and killed as a process group")
    script = tmp_path / "soffice"
    script.write_text(FAKE_SOFFICE)
    script.chmod(0o755)
    monkeypatch.setenv("FAKE_SOFFICE_ARGV", str(tmp_path / "argv.txt"))
    monkeypatch.setenv("FAKE_SOFFICE_HEARTBEAT", str(tmp_path / "heartbeat"))
    monkeypatch.setattr(harness, "LO_PROFILE", tmp_path / "lo profile")  # the space is the point
    return script


@pytest.fixture
def edited(tmp_path: Path) -> Path:
    """simple.docx after the harness's edit: first word replaced, tracked, saved."""
    out1 = tmp_path / "simple_edited.docx"
    doc = Document.open(SIMPLE, author=harness.AUTHOR, workspace_dir=tmp_path / "work_edit", force_recreate=True)
    try:
        target = next(p for p in doc.list_paragraphs_structured(limit=None) if p.text.strip())
        word = target.text.split()[0]
        doc.replace(word, word + harness.EDIT_MARKER, paragraph=target.ref, occurrence=0)
        doc.save(out1)
    finally:
        doc.close(cleanup=True)
    return out1


# --------------------------------------------------------------------------
# soffice output scan
# --------------------------------------------------------------------------


def test_soffice_messages_keeps_errors_and_drops_javaldx_noise():
    text = (
        "convert /a.docx -> /b.pdf using filter : writer_pdf_Export\n"
        f"{JAVALDX}\n"
        "Error: source file could not be loaded\n"
        "Warning: something about the document\n"
    )
    assert harness.soffice_messages(text) == [
        "Error: source file could not be loaded",
        "Warning: something about the document",
    ]


def test_soffice_messages_is_empty_for_a_clean_run():
    assert harness.soffice_messages(f"convert /a.docx -> /b.docx using filter : MS Word 2007 XML\n{JAVALDX}\n") == []


# --------------------------------------------------------------------------
# w:trackRevisions detection
# --------------------------------------------------------------------------


@pytest.mark.parametrize(
    "element, expected",
    [
        ("<w:trackRevisions/>", True),
        ('<w:trackRevisions w:val="true"/>', True),
        ('<w:trackRevisions w:val="1"/>', True),
        ('<w:trackRevisions w:val="on"/>', True),
        ('<w:trackRevisions w:val="TRUE"/>', True),
        ('<w:trackRevisions w:val="false"/>', False),
        ('<w:trackRevisions w:val="0"/>', False),
        # Another prefix for the same namespace: element and attribute are read by local name.
        (
            '<x:trackRevisions xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main" x:val="false"/>',
            False,
        ),
        ("", False),
    ],
)
def test_track_revisions_on(tmp_path: Path, element: str, expected: bool):
    def transform(xml: str) -> str:
        xml = re.sub(r"<w:trackRevisions\b[^>]*/>", "", xml)
        return re.sub(r"(<w:settings\b[^>]*>)", lambda m: m.group(1) + element, xml, count=1)

    path = rewrite_part(SIMPLE, tmp_path / "s.docx", "word/settings.xml", transform)
    assert harness.track_revisions_on(path) is expected


def test_track_revisions_on_is_none_without_settings_part(tmp_path: Path):
    path = tmp_path / "no_settings.docx"
    replace_docx_parts(SIMPLE, path, {"word/settings.xml": None})
    assert harness.track_revisions_on(path) is None


# --------------------------------------------------------------------------
# survival check (the #77 class), without LibreOffice
# --------------------------------------------------------------------------


def test_edited_file_carries_the_flag_and_a_redline(edited: Path):
    # Precondition for the survival tests: the library wrote what the check looks for.
    assert harness.track_revisions_on(edited) is True
    assert harness.census_file(edited)["by_tag"] == {"w:ins": 1, "w:del": 1}


def test_survival_check_passes_when_nothing_was_dropped(edited: Path, tmp_path: Path):
    assert harness.survival_check(edited, edited, tmp_path / "work", harness.AUTHOR) == {
        "status": "pass",
        "survival": "checked",
        "flag": "checked",
    }


def test_survival_check_says_when_the_flag_check_did_not_apply(edited: Path, tmp_path: Path):
    # A producer wrote w:val="false" and the library preserved it: LibreOffice
    # dropping that element loses nothing, and the record must not claim the
    # flag was checked.
    off = rewrite_part(
        edited,
        tmp_path / "flag_off.docx",
        "word/settings.xml",
        lambda xml: re.sub(r"<w:trackRevisions\b[^>]*/>", '<w:trackRevisions w:val="false"/>', xml),
    )
    assert harness.track_revisions_on(off) is False
    record = harness.survival_check(off, off, tmp_path / "work", harness.AUTHOR)
    assert record["status"] == "pass"
    assert record["flag"] == "not applicable (not on in the edited file)"


def test_survival_check_says_when_the_edited_file_has_no_settings_part(edited: Path, tmp_path: Path):
    # The library saves a document with no settings part without adding one.
    no_settings = tmp_path / "no_settings_edited.docx"
    replace_docx_parts(edited, no_settings, {"word/settings.xml": None})
    record = harness.survival_check(no_settings, no_settings, tmp_path / "work", harness.AUTHOR)
    assert record["status"] == "pass"
    assert record["flag"] == "not applicable (no settings part in the edited file)"


def test_survival_check_detects_a_dropped_track_revisions_flag(edited: Path, tmp_path: Path):
    roundtrip = rewrite_part(
        edited,
        tmp_path / "flag_dropped.docx",
        "word/settings.xml",
        lambda xml: re.sub(r"<w:trackRevisions\b[^>]*/>", "", xml),
    )
    record = harness.survival_check(edited, roundtrip, tmp_path / "work", harness.AUTHOR)
    assert record["status"] == "fail"
    assert record["error_type"] == "AssertTrackRevisionsDropped"


def test_survival_check_detects_a_lost_edit_marker(edited: Path, tmp_path: Path):
    # The pristine input posing as the round-trip: no marker, no revisions.
    # Its settings.xml has no flag either, so this also proves the flag check
    # fires first only when the flag was actually written and then lost.
    roundtrip = rewrite_part(
        edited,
        tmp_path / "marker_lost.docx",
        "word/document.xml",
        lambda xml: xml.replace(harness.EDIT_MARKER, ""),
    )
    record = harness.survival_check(edited, roundtrip, tmp_path / "work", harness.AUTHOR)
    assert record["error_type"] == "AssertEditMarkerLostInRoundtrip"


def test_survival_check_detects_a_dropped_own_revision(edited: Path, tmp_path: Path):
    # Unwrap the insertion (keep its runs, drop the w:ins element): the marker
    # is still visible, but it is no longer a revision by our author.
    roundtrip = rewrite_part(
        edited,
        tmp_path / "ins_unwrapped.docx",
        "word/document.xml",
        lambda xml: re.sub(r"<w:ins\b[^>/]*>(.*?)</w:ins>", r"\1", xml, flags=re.DOTALL),
    )
    record = harness.survival_check(edited, roundtrip, tmp_path / "work", harness.AUTHOR)
    assert record["error_type"] == "AssertOwnRevisionsDropped"
    assert "no insertion" in record["error"]


# --------------------------------------------------------------------------
# soffice_convert, against the fake soffice
# --------------------------------------------------------------------------


def test_soffice_convert_passes_a_real_file_url_and_reports_noise_free_diagnostics(
    fake_soffice: Path, edited: Path, tmp_path: Path
):
    out = tmp_path / "out"
    out.mkdir()
    run = harness.soffice_convert(str(fake_soffice), edited, "docx:MS Word 2007 XML", out, 10)
    assert run.output == out / edited.name and not run.timed_out and run.messages == []
    argv = (tmp_path / "argv.txt").read_text().splitlines()
    # A space in the profile path hangs soffice unless the URL is percent-encoded.
    assert argv[1] == f"-env:UserInstallation={(tmp_path / 'lo profile').as_uri()}"
    assert "%20" in argv[1]
    assert "javaldx" not in run.raw and run.raw.startswith("convert ")


def test_soffice_convert_refused_load_keeps_the_error_line_without_noise(
    fake_soffice: Path, edited: Path, tmp_path: Path, monkeypatch
):
    monkeypatch.setenv("FAKE_SOFFICE_MODE", "refuse")
    run = harness.soffice_convert(str(fake_soffice), edited, "pdf", tmp_path, 10)
    assert run.output is None and run.returncode == 0
    assert run.messages == ["Error: source file could not be loaded"]
    assert run.raw == "Error: source file could not be loaded"


def test_soffice_convert_timeout_kills_the_whole_process_tree(
    fake_soffice: Path, edited: Path, tmp_path: Path, monkeypatch
):
    # soffice forks soffice.bin; killing only the wrapper leaves an orphan
    # holding the profile lock. The grandchild's heartbeat must stop.
    monkeypatch.setenv("FAKE_SOFFICE_MODE", "hang")
    heartbeat = tmp_path / "heartbeat"
    run = harness.soffice_convert(str(fake_soffice), edited, "pdf", tmp_path, 2)
    assert run.timed_out and run.output is None and run.returncode < 0
    assert heartbeat.exists(), "the grandchild never started"
    time.sleep(0.3)
    first = os.stat(heartbeat).st_mtime_ns
    time.sleep(0.3)
    assert os.stat(heartbeat).st_mtime_ns == first, "the grandchild outlived the timeout"


# --------------------------------------------------------------------------
# the stage function
# --------------------------------------------------------------------------


def test_run_lo_roundtrip_skips_without_soffice(edited: Path, tmp_path: Path):
    assert harness.run_lo_roundtrip(edited, True, None, tmp_path / "work") == {
        "status": "skip",
        "reason": "soffice not found",
    }


@pytest.mark.parametrize(
    "run, error_type",
    [
        (harness.SofficeRun(output=None, messages=[], returncode=0, timed_out=True, raw=""), "LoRoundtripTimeout"),
        (
            harness.SofficeRun(
                output=None,
                messages=["Error: source file could not be loaded"],
                returncode=0,
                timed_out=False,
                raw="Error: source file could not be loaded",
            ),
            "LoMessage",
        ),
        # soffice exits 0 after refusing a file: the missing output is the signal.
        (harness.SofficeRun(output=None, messages=[], returncode=0, timed_out=False, raw=""), "LoLoadRefused"),
        # A crash after writing good output is still a failure, named apart from a refused load.
        (
            harness.SofficeRun(output=SIMPLE, messages=[], returncode=1, timed_out=False, raw="convert x"),
            "LoNonzeroExit",
        ),
    ],
)
def test_run_lo_roundtrip_reads_soffice_the_only_way_it_can_be_read(
    edited: Path, tmp_path: Path, monkeypatch, run, error_type
):
    monkeypatch.setattr(harness, "soffice_convert", lambda *args, **kwargs: run)
    record = harness.run_lo_roundtrip(edited, True, "soffice", tmp_path / "work")
    assert record["status"] == "fail"
    assert record["error_type"] == error_type
    if error_type == "LoNonzeroExit":
        assert record["error"] == "exit 1 after writing the docx: convert x"


def test_run_lo_roundtrip_names_a_soffice_that_cannot_be_started(edited: Path, tmp_path: Path, monkeypatch):
    # shutil.which() happily returns a half-removed install; that is a stage
    # failure with a name, not a crash that discards the file's other stages.
    monkeypatch.setattr(harness, "LO_DIR", tmp_path)
    broken = tmp_path / "soffice"
    broken.write_text("#!/nonexistent/interpreter\n")
    broken.chmod(0o755)
    record = harness.run_lo_roundtrip(edited, True, str(broken), tmp_path / "work")
    assert record["status"] == "fail"
    assert record["error_type"] == "SofficeSpawnFailed"


def test_run_lo_roundtrip_fails_on_an_unparseable_output(edited: Path, tmp_path: Path, monkeypatch):
    bad = rewrite_part(edited, tmp_path / "bad.docx", "word/document.xml", lambda xml: xml + "<w:p>")
    run = harness.SofficeRun(output=bad, messages=[], returncode=0, timed_out=False, raw="")
    monkeypatch.setattr(harness, "soffice_convert", lambda *args, **kwargs: run)
    record = harness.run_lo_roundtrip(edited, True, "soffice", tmp_path / "work")
    assert record["status"] == "fail"
    assert record["error_type"].startswith("LoOutputInvalid:")


def test_run_lo_roundtrip_skips_survival_when_nothing_was_edited(edited: Path, tmp_path: Path, monkeypatch):
    run = harness.SofficeRun(output=edited, messages=[], returncode=0, timed_out=False, raw="")
    monkeypatch.setattr(harness, "soffice_convert", lambda *args, **kwargs: run)
    assert harness.run_lo_roundtrip(edited, False, "soffice", tmp_path / "work") == {
        "status": "pass",
        "survival": "skipped (no edit)",
    }


# --------------------------------------------------------------------------
# manifest expectations (applied by the parent, like must_reject)
# --------------------------------------------------------------------------


def _record(lo_stage: dict, provenance: dict) -> dict:
    stages = {"open": {"status": "pass"}, "lo_roundtrip": lo_stage}
    return {"file": "x.docx", "stages": stages, "provenance": provenance}


def test_survival_waiver_turns_a_dropped_redline_into_a_documented_skip():
    rec = _record(
        {"status": "fail", "error_type": "AssertOwnRevisionsDropped", "error": "no deletion by CorpusHarness ..."},
        {"survival_waiver": "field result"},
    )
    harness.apply_manifest_expectations(rec)
    assert rec["stages"]["lo_roundtrip"] == {
        "status": "skip",
        "reason": "survival waived: field result",
        "dropped": "AssertOwnRevisionsDropped",
    }
    assert not harness.file_failed(rec)


def test_survival_waiver_does_not_cover_a_refused_load():
    rec = _record(
        {"status": "fail", "error_type": "LoLoadRefused", "error": "Error: source file could not be loaded"},
        {"survival_waiver": "field result"},
    )
    harness.apply_manifest_expectations(rec)
    assert rec["stages"]["lo_roundtrip"]["error_type"] == "LoLoadRefused"
    assert harness.file_failed(rec)


@pytest.mark.parametrize("error_type", ["AssertTrackRevisionsDropped", "AssertEditMarkerLostInRoundtrip"])
def test_survival_waiver_covers_only_a_dropped_own_revision(error_type: str):
    # A waiver says LibreOffice kept our text but not our revision. A dropped
    # flag (the #66/#77 class) or a vanished marker is a loss on any file.
    rec = _record({"status": "fail", "error_type": error_type, "error": "..."}, {"survival_waiver": "field result"})
    harness.apply_manifest_expectations(rec)
    assert rec["stages"]["lo_roundtrip"]["error_type"] == error_type
    assert harness.file_failed(rec)


def test_survival_waiver_that_is_no_longer_needed_fails():
    rec = _record({"status": "pass", "survival": "checked"}, {"survival_waiver": "field result"})
    harness.apply_manifest_expectations(rec)
    assert rec["stages"]["lo_roundtrip"]["status"] == "fail"
    assert rec["stages"]["lo_roundtrip"]["error_type"] == "StaleSurvivalWaiver"


def test_survival_waiver_leaves_an_unedited_file_alone():
    rec = _record({"status": "pass", "survival": "skipped (no edit)"}, {"survival_waiver": "field result"})
    harness.apply_manifest_expectations(rec)
    assert rec["stages"]["lo_roundtrip"] == {"status": "pass", "survival": "skipped (no edit)"}


def test_without_a_waiver_a_dropped_redline_stays_a_failure():
    rec = _record({"status": "fail", "error_type": "AssertTrackRevisionsDropped", "error": "..."}, {})
    harness.apply_manifest_expectations(rec)
    assert rec["stages"]["lo_roundtrip"]["error_type"] == "AssertTrackRevisionsDropped"
    assert harness.file_failed(rec)


def test_must_reject_violation_is_still_enforced():
    rec = _record({"status": "pass", "survival": "checked"}, {"must_reject": True})
    harness.apply_manifest_expectations(rec)
    assert rec["stages"]["open"]["error_type"] == "MustRejectViolation"


@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice not installed")
def test_run_lo_roundtrip_against_real_libreoffice(edited: Path, tmp_path: Path, monkeypatch):
    lo_dir = tmp_path / "lo"
    lo_dir.mkdir()
    monkeypatch.setattr(harness, "LO_DIR", lo_dir)
    monkeypatch.setattr(harness, "LO_PROFILE", tmp_path / "profile")
    record = harness.run_lo_roundtrip(edited, True, shutil.which("soffice"), tmp_path / "work")
    assert record == {"status": "pass", "survival": "checked", "flag": "checked"}
    assert (lo_dir / edited.name).exists()
