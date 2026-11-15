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

from docx_editor import Document
from docx_editor.cli import AGENT_GUIDE, main

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
        retry = next(line for line in err.splitlines() if line.startswith("retry: "))
        assert retry.startswith(f"retry: docx-edit {shlex.join(argv)} --occurrence N")
        assert not out_path.exists()
        assert not ws.exists() or not any(ws.iterdir())
        # The retry, with N filled in, is the one command that finishes the job.
        code, out, _ = run(capsys, *argv, "--occurrence", 1)
        assert code == 0 and "testing [del#0:Me]the[/del][ins#1:Me]THE[/ins] editing" in out

    def test_not_found_points_at_list(self, capsys, src):
        code, _, err = run(capsys, "replace", src, "zebra", "x", "--author", AUTHOR, "--in-place")
        assert code == 1
        assert err.splitlines()[-1] == f"retry: docx-edit list {shlex.quote(str(src))}"

    def test_stale_ref_retry_carries_the_fresh_ref(self, capsys, src):
        stale = ref_of(src, 2)
        run(capsys, "replace", src, "quick", "swift", "--author", AUTHOR, "--in-place")
        before = src.read_bytes()
        argv = ["replace", src, "brown", "red", "--paragraph", stale, "--author", AUTHOR, "--in-place"]
        code, _, err = run(capsys, *argv)
        assert code == 1 and src.read_bytes() == before
        fresh = ref_of(src, 2)
        assert err.splitlines()[-1] == "retry: " + shlex.join(["docx-edit", *[str(a) for a in argv]]).replace(
            shlex.quote(stale), shlex.quote(fresh)
        )

    def test_expect_mismatch_writes_nothing_and_offers_the_actual_count(self, capsys, src):
        before = src.read_bytes()
        code, _, err = run(capsys, "replace", src, "the", "THE", "--expect", 3, "--author", AUTHOR, "--in-place")
        assert code == 1 and src.read_bytes() == before
        assert "matches 2 times in the document, expected 3" in err
        assert "--expect 2" in err.splitlines()[-1]

    def test_unknown_revision_id_saves_nothing(self, capsys, src):
        run(capsys, "replace", src, "quick", "swift", "--author", AUTHOR, "--in-place")
        before = src.read_bytes()
        code, _, err = run(capsys, "accept", src, 0, 99, "--author", AUTHOR, "--in-place")
        assert code == 1 and src.read_bytes() == before
        assert "No pending revision 99" in err
        assert err.splitlines()[-1] == f"retry: docx-edit revisions {shlex.quote(str(src))}"

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
        assert payload["retry"].endswith("--occurrence N   (N from the list above)")


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
        assert code == 0 and out.startswith("P2#") and len(out.splitlines()) >= 1
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
        assert code == 0 and out.startswith("rejected 2 revision(s)")
        assert {t for _, t in revisions(src)} == {"lazy", "sleepy"}
        with Document.open(src) as doc:
            ids = [r.id for r in doc.list_revisions()]
        code, out, _ = run(capsys, "accept", src, *ids, "--author", AUTHOR, "--in-place")
        assert code == 0 and out.startswith("accepted 2 revision(s)")
        assert revisions(src) == []
        with Document.open(src) as doc:
            assert "sleepy dog" in doc.get_visible_text()


def test_installed_entry_point(src):
    exe = Path(sys.executable).parent / "docx-edit"
    if not exe.exists():
        pytest.skip("docx-edit entry point not installed in this environment")
    done = subprocess.run([str(exe), "find", str(src), "quick"], capture_output=True, text=True, check=False)
    assert done.returncode == 0 and done.stdout.startswith("1 match(es):")
