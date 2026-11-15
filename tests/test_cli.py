"""The docx-edit one-shot CLI (ROADMAP.md #104).

Each test drives ``main([...])`` in-process and checks the three things an
agent relies on: the exit status, the receipt (or the refusal with its retry
line), and the file on disk — written only when the command succeeded.
"""

import json
import shlex
import shutil
import subprocess
import sys
from pathlib import Path

import pytest
from test_settings_xml import with_settings

from docx_editor import Document
from docx_editor.cli import AGENT_GUIDE, _parser, _retry_for, main
from docx_editor.exceptions import RevisionError

AUTHOR = "Me"


@pytest.fixture
def ws(tmp_path, monkeypatch) -> Path:
    """An isolated workspace base, so a test can check nothing is left behind."""
    base = tmp_path / "ws"
    monkeypatch.setenv("DOCX_EDITOR_WORKSPACE_DIR", str(base))
    return base


@pytest.fixture
def src(simple_docx, tmp_path, ws) -> Path:
    # simple.docx paragraphs: P2 "The quick brown fox jumps over the lazy dog."
    # P3 "This is a sample document for testing the editing features."
    path = tmp_path / "f.docx"
    shutil.copy(simple_docx, path)
    return path


def run(capsys, *argv: object) -> tuple[int, str, str]:
    code = main([str(a) for a in argv])
    captured = capsys.readouterr()
    return code, captured.out, captured.err


def retry_argv(err: str, **fill: str) -> list[str]:
    """The printed ``retry:`` command as argv for ``main``, placeholders filled in."""
    line = next(line for line in err.splitlines() if line.startswith("retry: "))
    argv = shlex.split(line.removeprefix("retry: "))
    assert argv[0] == "docx-edit"
    return [fill.get(a, a) for a in argv[1:]]


def revisions(path: Path) -> list[tuple[str, str]]:
    with Document.open(path) as doc:
        return [(r.type, r.text) for r in doc.list_revisions()]


def ref_of(path: Path, index: int) -> str:
    with Document.open(path) as doc:
        return doc.list_paragraphs()[index - 1].split("|")[0]


class TestEditReceipts:
    def test_replace_prints_the_ref_the_paragraph_and_the_saved_path(self, capsys, src, tmp_path, ws):
        out_path = tmp_path / "out.docx"
        before = src.read_bytes()
        code, out, err = run(capsys, "replace", src, "quick", "swift", "--author", AUTHOR, "-o", out_path)
        assert (code, err) == (0, "")
        lines = out.splitlines()
        assert lines[0].startswith("replace P2#") and "revisions " in lines[0]
        ref = lines[0].split()[1]
        assert lines[1] == f"  {ref}| The [del#0:Me]quick[/del][ins#1:Me]swift[/ins] brown fox jumps over the lazy dog."
        assert lines[-1] == f"saved {out_path}"
        assert revisions(out_path) == [("deletion", "quick"), ("insertion", "swift")]
        assert ref_of(out_path, 2) == ref  # the printed ref is live in the saved file
        assert src.read_bytes() == before
        assert not ws.exists() or not any(ws.iterdir())

    def test_in_place_overwrites_the_source(self, capsys, src):
        code, out, _ = run(capsys, "delete", src, "lazy ", "--author", AUTHOR, "--in-place")
        assert code == 0 and out.splitlines()[-1] == f"saved {src}"
        assert revisions(src) == [("deletion", "lazy ")]

    @pytest.mark.parametrize(
        ("verb", "anchor", "text", "markup"),
        [
            ("insert-after", "brown", " big", "brown[ins#0:Me] big[/ins] fox"),
            ("insert-before", "fox", "red ", "brown [ins#0:Me]red [/ins]fox"),
        ],
    )
    def test_inserts(self, capsys, src, verb, anchor, text, markup):
        code, out, _ = run(capsys, verb, src, anchor, text, "--author", AUTHOR, "--in-place")
        assert code == 0 and markup in out

    def test_note_is_reported_as_a_comment(self, capsys, src):
        code, out, _ = run(capsys, "replace", src, "quick", "swift", "--note", "tone", "--author", AUTHOR, "--in-place")
        assert code == 0 and "comment 0" in out.splitlines()[0]
        with Document.open(src) as doc:
            assert [c.text for c in doc.list_comments()] == ["tone"]

    def test_comment_names_its_paragraph(self, capsys, src):
        code, out, _ = run(capsys, "comment", src, "lazy dog", "why a dog?", "--author", AUTHOR, "--in-place")
        assert code == 0
        assert out.splitlines()[0] == f"comment 0 on {ref_of(src, 2)}"

    def test_expect_replaces_every_match(self, capsys, src):
        code, out, _ = run(capsys, "replace", src, "the", "THE", "--expect", 2, "--author", AUTHOR, "--in-place")
        assert code == 0
        assert [line.split()[0] for line in out.splitlines()[:2]] == ["replace", "replace"]
        assert sorted(t for kind, t in revisions(src) if kind == "insertion") == ["THE", "THE"]

    def test_expect_zero_says_nothing_changed(self, capsys, src):
        code, out, _ = run(capsys, "replace", src, "zebra", "x", "--expect", 0, "--author", AUTHOR, "--in-place")
        assert code == 0 and out.splitlines()[0] == "replace: no match, nothing changed"

    def test_library_warning_is_reported(self, capsys, src):
        # A note on an edit that only amends your own pending insertion creates
        # no revision to anchor to: the dropped note must not be silent.
        run(capsys, "insert-after", src, "brown", " big", "--author", AUTHOR, "--in-place")
        code, out, _ = run(capsys, "replace", src, "big", "huge", "--note", "n", "--author", AUTHOR, "--in-place")
        assert code == 0
        assert any(line.startswith("warning: note='n' was not recorded") for line in out.splitlines())
        assert "ResourceWarning" not in out and "unclosed file" not in out


class TestRefusals:
    def test_ambiguous_prints_the_pick_list_and_a_runnable_retry(self, capsys, src, tmp_path, ws):
        out_path = tmp_path / "out.docx"
        argv = ["replace", str(src), "the", "THE", "--author", AUTHOR, "-o", str(out_path)]
        code, out, err = run(capsys, *argv)
        assert (code, out) == (1, "")
        assert "occurrence=0  P2#" in err and "occurrence=1  P3#" in err
        assert "note: N: an occurrence= from the list above" in err
        assert not out_path.exists()
        assert not ws.exists() or not any(ws.iterdir())
        # The retry, with N filled in, is the one command that finishes the job.
        code, out, _ = run(capsys, *retry_argv(err, N="1"))
        assert code == 0 and "testing [del#0:Me]the[/del][ins#1:Me]THE[/ins] editing" in out

    def test_retry_survives_double_dash_and_dash_leading_text(self, capsys, src):
        # Text starting with "-" needs "--"; the retry must keep it text.
        code, _, err = run(capsys, "replace", "--author", AUTHOR, "--in-place", "--", src, "the", "- THE")
        assert code == 1
        code, out, _ = run(capsys, *retry_argv(err, N="0"))
        assert code == 0 and "[ins#1:Me]- THE[/ins]" in out

    def test_repeated_flag_retry_uses_the_value_argparse_used(self, capsys, src):
        stale = ref_of(src, 2)
        run(capsys, "replace", src, "quick", "swift", "--author", AUTHOR, "--in-place")
        code, _, err = run(
            capsys,
            "replace",
            src,
            "brown",
            "red",
            "--paragraph",
            "P9#0000",
            "--paragraph",
            stale,
            "--author",
            AUTHOR,
            "--in-place",
        )
        assert code == 1
        code, out, _ = run(capsys, *retry_argv(err))
        assert code == 0 and "[ins#3:Me]red[/ins]" in out

    def test_occurrence_out_of_range_retry_is_runnable(self, capsys, src):
        code, _, err = run(capsys, "replace", src, "the", "THE", "--occurrence", 5, "--author", AUTHOR, "--in-place")
        assert code == 1 and "note: N: 0..1" in err
        code, _, _ = run(capsys, *retry_argv(err, N="1"))
        assert code == 0

    def test_not_found_points_at_list(self, capsys, src):
        code, _, err = run(capsys, "replace", src, "zebra", "x", "--author", AUTHOR, "--in-place")
        assert code == 1
        assert err.splitlines()[-1] == f"retry: docx-edit list -- {shlex.quote(str(src))}"

    def test_not_found_in_a_paragraph_lists_that_paragraph(self, capsys, src):
        ref = ref_of(src, 3)
        code, _, err = run(capsys, "delete", src, "quick", "--paragraph", ref, "--author", AUTHOR, "--in-place")
        assert code == 1
        assert retry_argv(err) == ["list", "--start", "3", "--limit", "1", "--", str(src)]

    def test_stale_ref_retry_carries_the_fresh_ref(self, capsys, src):
        stale = ref_of(src, 2)
        run(capsys, "replace", src, "quick", "swift", "--author", AUTHOR, "--in-place")
        before = src.read_bytes()
        argv = ["replace", src, "brown", "red", "--paragraph", stale, "--author", AUTHOR, "--in-place"]
        code, _, err = run(capsys, *argv)
        assert code == 1 and src.read_bytes() == before
        fresh = ref_of(src, 2)
        retry = retry_argv(err)
        assert retry[retry.index("--paragraph") + 1] == fresh
        code, out, _ = run(capsys, *retry)
        assert code == 0 and "[ins#3:Me]red[/ins]" in out

    def test_expect_mismatch_writes_nothing_and_offers_the_actual_count(self, capsys, src):
        before = src.read_bytes()
        code, _, err = run(capsys, "replace", src, "the", "THE", "--expect", 3, "--author", AUTHOR, "--in-place")
        assert code == 1 and src.read_bytes() == before
        assert "matches 2 times in the document, expected 3" in err
        assert "--expect 2" in err.splitlines()[-1]
        assert "note: only if every match listed above should change" in err

    def test_expect_with_nothing_matching_does_not_offer_expect_0(self, capsys, src):
        code, _, err = run(capsys, "replace", src, "zebra", "x", "--expect", 2, "--author", AUTHOR, "--in-place")
        assert code == 1 and retry_argv(err) == ["list", "--", str(src)]

    def test_find_count_mismatch_has_no_retry(self, capsys, src):
        code, _, err = run(capsys, "find", src, "the", "--expect", 3)
        assert code == 1 and "retry:" not in err

    def test_unknown_revision_id_saves_nothing(self, capsys, src):
        run(capsys, "replace", src, "quick", "swift", "--author", AUTHOR, "--in-place")
        before = src.read_bytes()
        code, _, err = run(capsys, "accept", src, 0, 99, "--author", AUTHOR, "--in-place")
        assert code == 1 and src.read_bytes() == before
        assert "No pending revision 99" in err
        assert err.splitlines()[-1] == f"retry: docx-edit revisions -- {shlex.quote(str(src))}"

    def test_shared_revision_id_retry_names_by_not_author(self, src):
        args = _parser().parse_args(["accept", str(src), "7", "--author", AUTHOR, "--in-place"])
        retry = _retry_for(RevisionError("two revisions share id 7", next="retry with author=<name>"), args)
        assert retry is not None and retry.argv[retry.argv.index("--by") + 1] == "NAME"
        assert "--author" in retry.argv  # who is resolving, unchanged

    def test_protected_document_is_readable_and_editable_on_retry(self, capsys, src, tmp_path):
        protected = with_settings(
            src, tmp_path / "p.docx", '<w:zoom/><w:documentProtection w:edit="readOnly" w:enforcement="1"/>'
        )
        code, out, _ = run(capsys, "list", protected)
        assert code == 0 and out.startswith("P1#")
        code, _, err = run(capsys, "replace", protected, "quick", "swift", "--author", AUTHOR, "--in-place")
        assert code == 1 and "--allow-protected" in err.splitlines()[-1]
        code, _, _ = run(capsys, *retry_argv(err))
        assert code == 0
        with Document.open(protected, allow_protected=True) as doc:
            assert [(r.type, r.text) for r in doc.list_revisions()] == [("deletion", "quick"), ("insertion", "swift")]

    def test_leftover_workspace_retry_recreates_it(self, capsys, src, tmp_path):
        doc = Document.open(src, author=AUTHOR)
        doc.replace("quick", "swift")
        doc.save(tmp_path / "elsewhere.docx")
        doc.close(cleanup=False)  # a session that never finished: unsaved edits stay behind
        code, _, err = run(capsys, "list", src)
        assert code == 1 and "discards the leftover workspace" in err
        code, out, _ = run(capsys, *retry_argv(err))
        assert code == 0 and "quick brown" in out

    def test_locked_document_is_refused(self, capsys, src):
        with Document.open(src):
            code, _, err = run(capsys, "replace", src, "quick", "swift", "--author", AUTHOR, "--in-place")
        assert code == 1 and err.startswith("error:")

    def test_json_refusal_carries_the_structured_error(self, capsys, src):
        code, out, err = run(capsys, "replace", src, "the", "THE", "--author", AUTHOR, "--in-place", "--json")
        assert (code, err) == (1, "")
        payload = json.loads(out)
        assert payload["ok"] is False
        assert payload["error"]["type"] == "AmbiguousTextError"
        assert payload["error"]["total_occurrences"] == 2
        assert [m["occurrence"] for m in payload["error"]["matches"]] == [0, 1]
        assert payload["error"]["next"].startswith("retry with occurrence=")
        argv = payload["retry_argv"]
        assert argv[argv.index("--occurrence") + 1] == "N" and "--json" in argv
        assert payload["retry"] == shlex.join(argv)
        assert payload["retry_note"] == "N: an occurrence= from the list above"

    def test_json_cli_refusal_has_a_next_key(self, capsys, src):
        run(capsys, "replace", src, "quick", "swift", "--author", AUTHOR, "--in-place")
        code, out, _ = run(capsys, "accept", src, 99, "--author", AUTHOR, "--in-place", "--json")
        error = json.loads(out)["error"]
        assert code == 1 and error == {"type": "Refusal", "message": error["message"], "next": None}


class TestUsage:
    @pytest.mark.parametrize(
        "argv",
        [
            ["replace", "{src}", "quick", "swift", "-o", "x.docx"],  # no --author
            ["replace", "{src}", "quick", "swift", "--author", AUTHOR],  # no output
            ["replace", "{src}", "quick", "swift", "--author", AUTHOR, "-o", "x.docx", "--in-place"],
            [
                "replace",
                "{src}",
                "the",
                "x",
                "--expect",
                "2",
                "--paragraph",
                "P2#0000",
                "--author",
                AUTHOR,
                "--in-place",
            ],
            ["accept", "{src}", "--author", AUTHOR, "--in-place"],  # neither ids nor --all
            ["accept", "{src}", "1", "--all", "--author", AUTHOR, "--in-place"],
        ],
    )
    def test_bad_command_lines_exit_2_and_touch_nothing(self, capsys, src, argv):
        before = src.read_bytes()
        with pytest.raises(SystemExit) as exc:
            main([a.format(src=src) for a in argv])
        assert exc.value.code == 2
        assert src.read_bytes() == before

    def test_verb_help_is_short_with_one_example(self, capsys):
        with pytest.raises(SystemExit):
            main(["replace", "--help"])
        out = capsys.readouterr().out
        assert out.count("example:") == 1 and "docx-edit replace F.docx" in out
        assert len(out.splitlines()) < 30

    def test_skill_prints_the_agent_guide(self, capsys):
        code, out, _ = run(capsys, "skill")
        assert code == 0 and out == AGENT_GUIDE
        for verb in ("list", "find", "replace", "delete", "insert-after", "comment", "accept", "reject"):
            assert f"docx-edit {verb} " in out


class TestReadVerbs:
    def test_find_lists_matches_with_edit_flags(self, capsys, src):
        code, out, _ = run(capsys, "find", src, "the")
        lines = out.splitlines()
        assert code == 0 and lines[0] == "2 match(es):"
        p2 = ref_of(src, 2)
        assert lines[1].startswith(f"  occurrence=0  {p2} at char 31")
        assert lines[1].endswith(f"(--paragraph {p2} --occurrence 0)")

    def test_find_expect_and_json(self, capsys, src):
        code, out, _ = run(capsys, "find", src, "the", "--expect", 2, "--json")
        payload = json.loads(out)
        assert code == 0 and payload["ok"] is True
        assert [m["paragraph_ref"].split("#")[0] for m in payload["matches"]] == ["P2", "P3"]
        code, _, err = run(capsys, "find", src, "the", "--expect", 1)
        assert code == 1 and "expected 1" in err

    def test_list_and_revisions(self, capsys, src):
        code, out, _ = run(capsys, "list", src, "--start", 2, "--limit", 1)
        assert code == 0 and out.startswith("P2#")
        notice = out.splitlines()[1]
        assert notice == f"... 2 more paragraphs; next page: docx-edit list --start 3 -- {src} (--limit 0 lists all)"
        with Document.open(src) as doc:
            total = doc.paragraph_count()
        code, out, _ = run(capsys, "list", src, "--limit", 0)
        assert code == 0 and len(out.splitlines()) == total
        run(capsys, "replace", src, "quick", "swift", "--author", "Bot", "--in-place")
        code, out, _ = run(capsys, "revisions", src, "--by", "Bot")
        assert code == 0 and out.splitlines()[0].startswith("#0 deletion by Bot in P2#") and '"quick"' in out
        code, out, _ = run(capsys, "revisions", src, "--by", "Nobody", "--json")
        assert json.loads(out)["revisions"] == []


class TestResolve:
    def test_accept_ids_and_reject_all_by_author(self, capsys, src):
        run(capsys, "replace", src, "quick", "swift", "--author", "Bot", "--in-place")
        run(capsys, "replace", src, "lazy", "sleepy", "--author", "Human", "--in-place")
        code, out, _ = run(capsys, "reject", src, "--all", "--by", "Bot", "--author", AUTHOR, "--in-place")
        assert code == 0 and out.startswith("rejected 2 revision(s); 2 pending")
        assert {t for _, t in revisions(src)} == {"lazy", "sleepy"}
        with Document.open(src) as doc:
            ids = [r.id for r in doc.list_revisions()]
        code, out, _ = run(capsys, "accept", src, *ids, "--author", AUTHOR, "--in-place")
        assert code == 0 and out.splitlines()[0] == "accepted 2 revision(s); 0 pending"
        # The receipt shows the paragraph as it now reads, under its current ref.
        assert out.splitlines()[1] == f"  {ref_of(src, 2)}| The quick brown fox jumps over the sleepy dog."
        assert revisions(src) == []
        with Document.open(src) as doc:
            assert "sleepy dog" in doc.get_visible_text()


def test_installed_entry_point(src):
    exe = Path(sys.executable).parent / "docx-edit"
    if not exe.exists():
        pytest.skip("docx-edit entry point not installed in this environment")
    done = subprocess.run([str(exe), "find", str(src), "quick"], capture_output=True, text=True, check=False)
    assert done.returncode == 0 and done.stdout.startswith("1 match(es):")


class TestJsonSuccess:
    @pytest.mark.parametrize(
        "argv",
        [
            ["replace", "{src}", "quick", "swift"],
            ["delete", "{src}", "lazy "],
            ["insert-after", "{src}", "brown", " big"],
            ["insert-before", "{src}", "fox", "red "],
        ],
    )
    def test_edit_verbs(self, capsys, src, argv):
        code, out, _ = run(capsys, *[a.format(src=src) for a in argv], "--author", AUTHOR, "--in-place", "--json")
        payload = json.loads(out)
        assert code == 0 and payload["ok"] is True and payload["verb"] == argv[0]
        assert set(payload) == {"ok", "verb", "results", "paragraphs", "saved", "warnings"}
        (result,) = payload["results"]
        assert result["revision_ids"] and result["ref"].startswith("P2#")
        (paragraph,) = payload["paragraphs"]
        assert paragraph["ref"] == result["ref"] and "[" in paragraph["markup"]
        assert payload["saved"] == str(src)

    def test_comment_and_accept(self, capsys, src):
        code, out, _ = run(capsys, "comment", src, "lazy dog", "why?", "--author", AUTHOR, "--in-place", "--json")
        payload = json.loads(out)
        assert code == 0 and payload["comment_id"] == 0 and payload["ref"] == ref_of(src, 2)
        run(capsys, "replace", src, "quick", "swift", "--author", AUTHOR, "--in-place")
        code, out, _ = run(capsys, "accept", src, "--all", "--author", AUTHOR, "--in-place", "--json")
        payload = json.loads(out)
        assert code == 0 and (payload["resolved"], payload["pending"], payload["unhandled"]) == (2, 0, 0)

    def test_warnings_ride_in_json(self, capsys, src):
        run(capsys, "insert-after", src, "brown", " big", "--author", AUTHOR, "--in-place")
        code, out, _ = run(
            capsys, "replace", src, "big", "huge", "--note", "n", "--author", AUTHOR, "--in-place", "--json"
        )
        assert code == 0 and json.loads(out)["warnings"][0].startswith("note='n' was not recorded")


def test_receipt_is_right_after_a_line_separator_in_a_paragraph(capsys, src):
    # U+2028 inside a paragraph's text must not shift later paragraphs' markup.
    run(capsys, "replace", src, "quick", "qu\u2028ick", "--author", AUTHOR, "--in-place")
    code, out, _ = run(capsys, "replace", src, "sample", "SAMPLE", "--author", AUTHOR, "--in-place")
    assert code == 0
    assert out.splitlines()[1].startswith(f"  {ref_of(src, 3)}| This is a [del#")
