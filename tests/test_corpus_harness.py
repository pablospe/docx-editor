"""Tests for the LibreOffice gate of the corpus harness (benchmarks/corpus/corpus_harness.py).

The harness is a script, not a package, so it is loaded by path. Everything
here runs without LibreOffice except the last test, which drives the real
``soffice`` when one is installed and is skipped otherwise.
"""

import importlib.util
import re
import shutil
import zipfile
from pathlib import Path

import pytest

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
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == part:
                data = transform(data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
    return dst


def drop_part(src: Path, dst: Path, part: str) -> Path:
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename != part:
                zout.writestr(item, zin.read(item.filename))
    return dst


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
        ('<w:trackRevisions w:val="false"/>', False),
        ('<w:trackRevisions w:val="0"/>', False),
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
    path = drop_part(SIMPLE, tmp_path / "no_settings.docx", "word/settings.xml")
    assert harness.track_revisions_on(path) is None


# --------------------------------------------------------------------------
# survival check (the #77 class), without LibreOffice
# --------------------------------------------------------------------------


def test_edited_file_carries_the_flag_and_a_redline(edited: Path):
    # Precondition for the survival tests: the library wrote what the check looks for.
    assert harness.track_revisions_on(edited) is True
    assert harness.census_file(edited)["by_tag"] == {"w:ins": 1, "w:del": 1}


def test_survival_check_passes_when_nothing_was_dropped(edited: Path, tmp_path: Path):
    assert harness.survival_check(edited, edited, tmp_path / "work", harness.AUTHOR) is None


def test_survival_check_detects_a_dropped_track_revisions_flag(edited: Path, tmp_path: Path):
    roundtrip = rewrite_part(
        edited,
        tmp_path / "flag_dropped.docx",
        "word/settings.xml",
        lambda xml: re.sub(r"<w:trackRevisions\b[^>]*/>", "", xml),
    )
    record = harness.survival_check(edited, roundtrip, tmp_path / "work", harness.AUTHOR)
    assert record is not None
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
    assert record is not None
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
    assert record is not None
    assert record["error_type"] == "AssertOwnRevisionsDropped"
    assert "no insertion" in record["error"]


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
    ],
)
def test_run_lo_roundtrip_reads_soffice_the_only_way_it_can_be_read(
    edited: Path, tmp_path: Path, monkeypatch, run, error_type
):
    monkeypatch.setattr(harness, "soffice_convert", lambda *args, **kwargs: run)
    record = harness.run_lo_roundtrip(edited, True, "soffice", tmp_path / "work")
    assert record["status"] == "fail"
    assert record["error_type"] == error_type


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
    assert record == {"status": "pass", "survival": "checked"}
    assert (lo_dir / edited.name).exists()
