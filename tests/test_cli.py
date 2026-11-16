"""The docx-edit one-shot CLI (ROADMAP.md #104).

Each test drives ``main([...])`` in-process and checks what an agent relies
on: the exit status, the receipt (or the refusal with its ``next:`` hint), and
the files on disk — written only when the command succeeded, and never the
input unless asked.
"""

import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

import pytest
from test_settings_xml import with_settings

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


def edited(path: Path) -> Path:
    return path.with_name(path.stem + ".edited.docx")


def revisions(path: Path) -> list[tuple[str, str]]:
    with Document.open(path) as doc:
        return [(r.type, r.text) for r in doc.list_revisions()]


def ref_of(path: Path, index: int) -> str:
    with Document.open(path) as doc:
        return doc.list_paragraphs()[index - 1].split("|")[0]


def nothing_left(ws: Path) -> bool:
    return not ws.exists() or not any(ws.iterdir())


class TestDefaultOutput:
    def test_edit_writes_stem_edited_and_leaves_the_input_alone(self, capsys, src, ws):
        before = src.read_bytes()
        code, out, err = run(capsys, "replace", src, "quick", "swift", "--author", AUTHOR)
        assert (code, err) == (0, "")
        assert out.splitlines()[-1] == f"saved {edited(src)}"
        assert src.read_bytes() == before
        assert revisions(edited(src)) == [("deletion", "quick"), ("insertion", "swift")]
        assert nothing_left(ws)

    def test_an_edited_input_is_updated_in_place_so_a_chain_keeps_one_file(self, capsys, src):
        run(capsys, "replace", src, "quick", "swift", "--author", AUTHOR)
        redline = edited(src)
        code, out, _ = run(capsys, "delete", redline, "lazy ", "--author", AUTHOR)
        assert code == 0 and out.splitlines()[-1] == f"saved {redline}"
        assert not edited(redline).exists()
        assert [t for _, t in revisions(redline)] == ["quick", "swift", "lazy "]

    def test_an_existing_edited_file_is_never_overwritten(self, capsys, src):
        run(capsys, "replace", src, "quick", "swift", "--author", AUTHOR)
        redline_before = edited(src).read_bytes()
        source_before = src.read_bytes()
        code, out, err = run(capsys, "replace", src, "lazy", "idle", "--author", AUTHOR)
        assert code == 1 and out == ""
        assert "already exists" in err and "-o OUT" in err and "--in-place" in err
        assert edited(src).read_bytes() == redline_before and src.read_bytes() == source_before

    def test_output_and_in_place_override(self, capsys, src, tmp_path):
        out_path = tmp_path / "elsewhere.docx"
        code, out, _ = run(capsys, "replace", src, "quick", "swift", "--author", AUTHOR, "-o", out_path)
        assert code == 0 and out.splitlines()[-1] == f"saved {out_path}" and not edited(src).exists()
        code, out, _ = run(capsys, "delete", src, "lazy ", "--author", AUTHOR, "--in-place")
        assert code == 0 and out.splitlines()[-1] == f"saved {src}"
        assert revisions(src) == [("deletion", "lazy ")]

    def test_the_rule_is_stated_in_help_and_guide(self, capsys):
        with pytest.raises(SystemExit):
            main(["replace", "--help"])
        assert "edited.docx" in capsys.readouterr().out
        assert "F.edited.docx" in AGENT_GUIDE and "updates it in place" in AGENT_GUIDE


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
        assert ref_of(out_path, 2) == ref  # the printed ref is live in the saved file
        assert src.read_bytes() == before
        assert nothing_left(ws)

    @pytest.mark.parametrize(
        ("verb", "anchor", "text", "markup"),
        [
            ("insert-after", "brown", " big", "brown[ins#0:Me] big[/ins] fox"),
            ("insert-before", "fox", "red ", "brown [ins#0:Me]red [/ins]fox"),
        ],
    )
    def test_inserts(self, capsys, src, verb, anchor, text, markup):
        code, out, _ = run(capsys, verb, src, anchor, text, "--author", AUTHOR)
        assert code == 0 and markup in out

    def test_note_is_reported_as_a_comment(self, capsys, src):
        code, out, _ = run(capsys, "replace", src, "quick", "swift", "--note", "tone", "--author", AUTHOR)
        assert code == 0 and "comment 0" in out.splitlines()[0]
        with Document.open(edited(src)) as doc:
            assert [c.text for c in doc.list_comments()] == ["tone"]

    def test_comment_names_its_paragraph(self, capsys, src):
        code, out, _ = run(capsys, "comment", src, "lazy dog", "why a dog?", "--author", AUTHOR)
        assert code == 0
        assert out.splitlines()[0] == f"comment 0 on {ref_of(edited(src), 2)}"

    def test_expect_replaces_every_match(self, capsys, src):
        code, out, _ = run(capsys, "replace", src, "the", "THE", "--expect", 2, "--author", AUTHOR)
        assert code == 0
        assert [line.split()[0] for line in out.splitlines()[:2]] == ["replace", "replace"]
        assert sorted(t for kind, t in revisions(edited(src)) if kind == "insertion") == ["THE", "THE"]

    def test_expect_zero_says_nothing_changed(self, capsys, src):
        code, out, _ = run(capsys, "replace", src, "zebra", "x", "--expect", 0, "--author", AUTHOR)
        assert code == 0 and out.splitlines()[0] == "replace: no match, nothing changed"

    def test_library_warning_is_reported(self, capsys, src):
        # A note on an edit that only amends your own pending insertion creates
        # no revision to anchor to: the dropped note must not be silent.
        run(capsys, "insert-after", src, "brown", " big", "--author", AUTHOR)
        code, out, _ = run(capsys, "replace", edited(src), "big", "huge", "--note", "n", "--author", AUTHOR)
        assert code == 0
        assert any(line.startswith("warning: note='n' was not recorded") for line in out.splitlines())
        assert "ResourceWarning" not in out and "unclosed file" not in out

    def test_receipt_is_right_after_a_line_separator_in_a_paragraph(self, capsys, src):
        # U+2028 inside a paragraph's text must not shift later paragraphs' markup.
        run(capsys, "replace", src, "quick", "qu ick", "--author", AUTHOR)
        redline = edited(src)
        code, out, _ = run(capsys, "replace", redline, "sample", "SAMPLE", "--author", AUTHOR)
        assert code == 0
        assert out.splitlines()[1].startswith(f"  {ref_of(redline, 3)}| This is a [del#")

    def test_newline_split_shows_both_halves(self, capsys, src):
        code, out, _ = run(capsys, "replace", src, "brown", "brown\nX", "--author", AUTHOR)
        assert code == 0 and len([line for line in out.splitlines() if line.startswith("  P")]) == 2


class TestRefusals:
    def test_ambiguous_prints_the_pick_list_and_an_occurrence_hint(self, capsys, src, ws):
        code, out, err = run(capsys, "replace", src, "the", "THE", "--author", AUTHOR)
        assert (code, out) == (1, "")
        assert "occurrence=0  P2#" in err and "occurrence=1  P3#" in err
        assert err.splitlines()[-1] == "next: add --occurrence N (0..1)"
        assert not edited(src).exists() and nothing_left(ws)
        # The hint, followed, is the one command that finishes the job.
        code, out, _ = run(capsys, "replace", src, "the", "THE", "--author", AUTHOR, "--occurrence", 1)
        assert code == 0 and "testing [del#0:Me]the[/del][ins#1:Me]THE[/ins] editing" in out

    def test_not_found_points_at_list(self, capsys, src):
        code, _, err = run(capsys, "replace", src, "zebra", "x", "--author", AUTHOR)
        assert code == 1 and err.splitlines()[-1] == f"next: docx-edit list {src}"

    def test_stale_ref_points_at_the_current_paragraph(self, capsys, src):
        stale = ref_of(src, 2)
        run(capsys, "replace", src, "quick", "swift", "--author", AUTHOR, "--in-place")
        before = src.read_bytes()
        code, _, err = run(capsys, "replace", src, "brown", "red", "--paragraph", stale, "--author", AUTHOR)
        assert code == 1 and src.read_bytes() == before and not edited(src).exists()
        assert f"Use {ref_of(src, 2)} to target current content" in err
        assert err.splitlines()[-1] == f"next: docx-edit list {src} --start 2 --limit 1"

    def test_expect_mismatch_writes_nothing_and_points_at_find(self, capsys, src):
        code, _, err = run(capsys, "replace", src, "the", "THE", "--expect", 3, "--author", AUTHOR)
        assert code == 1 and not edited(src).exists()
        assert "matches 2 times in the document, expected 3" in err
        assert err.splitlines()[-1] == f"next: docx-edit find {src} the"

    def test_unknown_revision_id_saves_nothing(self, capsys, src):
        run(capsys, "replace", src, "quick", "swift", "--author", AUTHOR, "--in-place")
        before = src.read_bytes()
        code, _, err = run(capsys, "accept", src, 0, 99, "--author", AUTHOR, "--in-place")
        assert code == 1 and src.read_bytes() == before
        assert "No pending revision 99" in err
        assert err.splitlines()[-1] == f"next: docx-edit revisions {src}"

    def test_locked_document_is_refused_with_a_way_out(self, capsys, src):
        with Document.open(src):
            code, _, err = run(capsys, "replace", src, "quick", "swift", "--author", AUTHOR)
            assert code == 1 and err.startswith("error:") and "next: close the session" in err
            # A read verb reads around the lock, in a private workspace.
            code, out, _ = run(capsys, "find", src, "quick")
            assert code == 0 and out.startswith("1 match(es):")
        assert not edited(src).exists()

    def test_protected_document_is_readable_and_editable_with_the_flag(self, capsys, src, tmp_path):
        protected = with_settings(
            src, tmp_path / "p.docx", '<w:zoom/><w:documentProtection w:edit="readOnly" w:enforcement="1"/>'
        )
        code, out, _ = run(capsys, "list", protected)
        assert code == 0 and out.startswith("P1#")
        code, _, err = run(capsys, "replace", protected, "quick", "swift", "--author", AUTHOR)
        assert code == 1 and err.splitlines()[-1].startswith("next: add --allow-protected")
        code, _, _ = run(capsys, "replace", protected, "quick", "swift", "--author", AUTHOR, "--allow-protected")
        assert code == 0
        with Document.open(edited(protected), allow_protected=True) as doc:
            assert [(r.type, r.text) for r in doc.list_revisions()] == [("deletion", "quick"), ("insertion", "swift")]

    def test_leftover_workspace_is_read_around_and_discarded_only_on_request(self, capsys, src, tmp_path):
        doc = Document.open(src, author=AUTHOR)
        doc.replace("quick", "swift")
        doc.save(tmp_path / "elsewhere.docx")
        doc.close(cleanup=False)  # a session that never finished: unsaved edits stay behind
        code, out, _ = run(capsys, "list", src)
        assert code == 0 and "quick brown" in out
        code, _, err = run(capsys, "replace", src, "lazy", "idle", "--author", AUTHOR)
        assert code == 1 and err.splitlines()[-1].startswith("next: add --force-recreate")
        code, out, _ = run(capsys, "replace", src, "lazy", "idle", "--author", AUTHOR, "--force-recreate")
        assert code == 0 and "the [del#0:Me]lazy[/del][ins#1:Me]idle[/ins] dog." in out

    def test_force_recreate_never_takes_over_a_live_session(self, capsys, src):
        with Document.open(src):
            code, _, err = run(capsys, "replace", src, "lazy", "idle", "--author", AUTHOR, "--force-recreate")
        assert code == 1 and err.startswith("error:") and not edited(src).exists()

    def test_json_refusal_carries_the_structured_error_and_the_hint(self, capsys, src):
        code, out, err = run(capsys, "replace", src, "the", "THE", "--author", AUTHOR, "--json")
        assert (code, err) == (1, "")
        payload = json.loads(out)
        assert payload["ok"] is False
        assert payload["error"]["type"] == "AmbiguousTextError"
        assert payload["error"]["total_occurrences"] == 2
        assert [m["occurrence"] for m in payload["error"]["matches"]] == [0, 1]
        assert payload["error"]["next"].startswith("retry with occurrence=")
        assert payload["next"] == "add --occurrence N (0..1)"
        assert set(payload) == {"ok", "error", "next"}

    def test_json_cli_refusal_has_a_next_key(self, capsys, src):
        run(capsys, "replace", src, "quick", "swift", "--author", AUTHOR, "--in-place")
        code, out, _ = run(capsys, "accept", src, 99, "--author", AUTHOR, "--in-place", "--json")
        error = json.loads(out)["error"]
        assert code == 1 and set(error) == {"type", "message", "next"} and error["type"] == "Refusal"


class TestUsage:
    @pytest.mark.parametrize(
        "argv",
        [
            ["replace", "{src}", "quick", "swift"],  # no --author
            ["replace", "{src}", "quick", "swift", "--author", AUTHOR, "-o", "x.docx", "--in-place"],
            ["replace", "{src}", "the", "x", "--expect", "2", "--paragraph", "P2#0000", "--author", AUTHOR],
            ["accept", "{src}", "--author", AUTHOR],  # neither ids nor --all
            ["accept", "{src}", "1", "--all", "--author", AUTHOR],
            ["replace", "{src}", "fox", "cat", "--author", AUTHOR, "--in"],  # no abbreviations
        ],
    )
    def test_bad_command_lines_exit_2_and_touch_nothing(self, capsys, src, argv):
        before = src.read_bytes()
        with pytest.raises(SystemExit) as exc:
            main([a.format(src=src) for a in argv])
        assert exc.value.code == 2
        assert src.read_bytes() == before and not edited(src).exists()

    def test_verb_help_is_short_with_one_example(self, capsys):
        with pytest.raises(SystemExit):
            main(["replace", "--help"])
        out = capsys.readouterr().out
        assert out.count("example:") == 1 and "docx-edit replace F.docx" in out
        assert len(out.splitlines()) < 35

    def test_skill_prints_the_agent_guide(self, capsys):
        code, out, _ = run(capsys, "skill")
        assert code == 0 and out == AGENT_GUIDE
        for verb in ("list", "find", "replace", "delete", "insert-after", "comment", "accept", "reject"):
            assert f"docx-edit {verb} " in out
        assert "retry" not in out


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

    def test_list_pages_and_revisions(self, capsys, src):
        code, out, _ = run(capsys, "list", src, "--start", 2, "--limit", 1)
        assert code == 0 and out.startswith("P2#")
        assert out.splitlines()[1] == "... 2 more paragraphs; next page: --start 3 (--limit 0 lists all)"
        with Document.open(src) as doc:
            total = doc.paragraph_count()
        code, out, _ = run(capsys, "list", src, "--limit", 0)
        assert code == 0 and len(out.splitlines()) == total
        code, out, _ = run(capsys, "list", src, "--start", 99)
        assert code == 0 and out.strip() == f"no paragraphs from 99 (the document has {total})"
        run(capsys, "replace", src, "quick", "swift", "--author", "Bot", "--in-place")
        code, out, _ = run(capsys, "revisions", src, "--by", "Bot")
        assert code == 0 and out.splitlines()[0].startswith("#0 deletion by Bot in P2#") and '"quick"' in out
        code, out, _ = run(capsys, "revisions", src, "--by", "Nobody", "--json")
        assert json.loads(out)["revisions"] == []

    def test_read_verbs_never_write(self, capsys, src, ws):
        before = src.read_bytes()
        for argv in (["list", src], ["find", src, "fox"], ["revisions", src]):
            assert run(capsys, *argv)[0] == 0
        assert src.read_bytes() == before and not edited(src).exists() and nothing_left(ws)


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

    def test_rejecting_a_split_by_id_reports_the_merge_not_a_neighbour(self, capsys, src):
        run(capsys, "replace", src, "brown", "brown\nX", "--author", AUTHOR, "--in-place")
        with Document.open(src) as doc:
            ids = [r.id for r in doc.list_revisions()]
        code, out, _ = run(capsys, "reject", src, *ids, "--author", AUTHOR, "--in-place")
        assert code == 0
        shown = [line for line in out.splitlines() if line.startswith("  P")]
        assert shown == [f"  {ref_of(src, 2)}| The quick brown fox jumps over the lazy dog."]


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
        code, out, _ = run(capsys, *[a.format(src=src) for a in argv], "--author", AUTHOR, "--json")
        payload = json.loads(out)
        assert code == 0 and payload["ok"] is True and payload["verb"] == argv[0]
        assert set(payload) == {"ok", "verb", "results", "paragraphs", "saved", "warnings"}
        (result,) = payload["results"]
        assert result["revision_ids"] and result["ref"].startswith("P2#")
        (paragraph,) = payload["paragraphs"]
        assert paragraph["ref"] == result["ref"] and "[" in paragraph["markup"]
        assert payload["saved"] == str(edited(src))

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


class TestHonestExitCodes:
    def test_workspace_removal_failure_after_save_is_a_warning(self, capsys, src, ws, monkeypatch):
        real_rmtree = shutil.rmtree

        def busy(path, *args, **kwargs):
            # Only the document's workspace is "held open by a scanner".
            if Path(path).is_relative_to(ws):
                raise PermissionError("file in use")
            return real_rmtree(path, *args, **kwargs)

        monkeypatch.setattr(shutil, "rmtree", busy)
        code, out, _ = run(capsys, "replace", src, "lazy", "idle", "--author", AUTHOR)
        assert code == 0
        assert any(
            line.startswith("warning: saved, but the workspace could not be removed") for line in out.splitlines()
        )
        monkeypatch.undo()
        assert revisions(edited(src)) == [("deletion", "lazy"), ("insertion", "idle")]

    def test_read_verb_survives_an_unremovable_private_workspace(self, capsys, src, monkeypatch):
        real_rmtree = shutil.rmtree

        def busy(path, *args, **kwargs):
            if kwargs.get("ignore_errors") or kwargs.get("onexc") or kwargs.get("onerror"):
                return real_rmtree(path, *args, **kwargs)
            raise PermissionError("held by scanner")

        monkeypatch.setattr(shutil, "rmtree", busy)
        code, out, _ = run(capsys, "find", src, "fox")
        assert code == 0 and out.startswith("1 match(es):")


def _entry_point() -> Path:
    exe = Path(sys.executable).parent / "docx-edit"
    if not exe.exists():
        pytest.skip("docx-edit entry point not installed in this environment")
    return exe


def test_installed_entry_point(src):
    done = subprocess.run([str(_entry_point()), "find", str(src), "quick"], capture_output=True, text=True, check=False)
    assert done.returncode == 0 and done.stdout.startswith("1 match(es):")


def test_unencodable_output_still_exits_0_after_saving(src, tmp_path):
    # A cp1252 pipe (Windows) cannot show "→"; the saved edit must not turn
    # into a traceback and exit 1, which means "nothing saved".
    out_path = tmp_path / "w.docx"
    env = {**os.environ, "PYTHONIOENCODING": "cp1252"}
    cmd = [str(_entry_point()), "replace", str(src), "lazy", "lazy → idle", "--author", AUTHOR, "-o", str(out_path)]
    done = subprocess.run(cmd, capture_output=True, env=env, check=False)
    assert done.returncode == 0 and out_path.exists()
    assert b"\\u2192" in done.stdout


@pytest.mark.parametrize("fail", ["success", "refusal"])
def test_a_closed_pipe_keeps_the_exit_status(src, tmp_path, fail):
    text = "quick" if fail == "success" else "zebra"
    cmd = [str(_entry_point()), "replace", str(src), text, "x", "--author", AUTHOR, "-o", str(tmp_path / "p.docx")]
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    assert proc.stdout is not None and proc.stderr is not None
    proc.stdout.close()
    proc.stderr.close()
    assert proc.wait() == (0 if fail == "success" else 1)
