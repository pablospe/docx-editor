"""docx-edit: one tracked-change edit per command (ROADMAP.md #104).

Each verb is one :class:`~docx_editor.Document` method, named in kebab-case,
with its keywords as flags: ``docx-edit replace F OLD NEW --paragraph P`` is
``doc.replace(OLD, NEW, paragraph=P)``. A command opens the file, makes the
one call, saves (mutating verbs only) and closes. Nothing is saved when the
call refuses.

The input is never touched unless asked: an edit writes ``<stem>.edited.docx``
beside the input, or updates the input in place when it is itself a
``*.edited.docx`` (so a chain of edits keeps one redline file); ``-o OUT`` and
``--in-place`` override.

Output is a receipt the caller can stop on: what changed, the paragraph ref,
the paragraph as it now reads (tracked-change markup), and where it was saved.
A refusal prints the library's message (with its pick-list) and a ``next:``
hint. ``--json`` prints the same as one JSON object on stdout. Exit status:
0 success, 1 refused (nothing saved), 2 usage error.
"""

import argparse
import contextlib
import json
import os
import re
import shlex
import sys
import tempfile
import warnings
from collections.abc import Sequence
from pathlib import Path

from . import Document, EditResult
from .exceptions import (
    DocumentProtectedError,
    DocxEditError,
    MatchCountError,
    UnanchoredNoteWarning,
    UnhandledRevisionWarning,
    WorkspaceLockedError,
    WorkspaceSyncError,
    _error_payload,
    _json_default,
    _match_context,
)
from .xml_editor import body_paragraphs

PROG = "docx-edit"
_EDIT_VERBS = ("replace", "delete", "insert-after", "insert-before")
_MUTATING = (*_EDIT_VERBS, "comment", "accept", "reject")
_EDITED = ".edited.docx"
_LIBRARY_WARNINGS = (UnanchoredNoteWarning, UnhandledRevisionWarning)

AGENT_GUIDE = """\
docx-edit: tracked-change edits to a .docx, one command each. Every edit is a
Word tracked change by --author; nothing is saved unless the command succeeds.

Look:
  docx-edit list F.docx                      paragraphs as "P3#a7b2| text"
  docx-edit find F.docx "30 days"            every match, with its ref
  docx-edit revisions F.docx [--by NAME]     pending tracked changes

Edit (each needs --author NAME):
  docx-edit replace F.docx "30 days" "60 days" --author Me
  docx-edit delete F.edited.docx "unless agreed" --author Me
  docx-edit insert-after F.edited.docx "Payment" " terms" --author Me
  docx-edit insert-before F.edited.docx "due" "now " --author Me
  docx-edit comment F.edited.docx "Effective Date" "Check this" --author Me
  docx-edit replace F.docx "30 days" "60 days" --expect 3 --author Me
  docx-edit accept F.edited.docx 5 6 --author Me     (or --all [--by NAME])
  docx-edit reject F.edited.docx --all --by Bot --author Me

Output file:
- An edit on F.docx writes F.edited.docx and never touches F.docx. An edit on
  a file already named *.edited.docx updates it in place, so a chain of edits
  keeps one redline file: edit F.docx once, then keep editing F.edited.docx.
- -o OUT writes elsewhere; --in-place overwrites the input. If F.edited.docx
  already exists, an edit on F.docx refuses rather than overwrite it.

Rules:
- Without --paragraph the text is searched in the whole document and must
  occur once. If it occurs more often the command refuses and lists up to 10
  matches as "occurrence=N  P3#a7b2 ..." (find lists all); rerun with
  --occurrence N. --paragraph P3#a7b2 (a ref from list/find) pins the edit to
  that paragraph as you saw it; a stale ref refuses.
- --expect N on replace replaces every match, and refuses unless there are
  exactly N. On find it checks the count without editing.
- --note "why" attaches your rationale as a comment on the change.
- A newline in the new text splits the paragraph (tracked), e.g. $'a\\nb'.
- Text starting with "-": put -- before the positional arguments, as in
  docx-edit replace --author Me -- F.docx "- item" "- entry".
- A moved paragraph is two revisions (move_from + move_to): accept or reject
  both ids together, or use --all.
- The receipt shows each changed paragraph as it now reads:
  [del#5:Me]old[/del][ins#6:Me]new[/ins]. It is the check; no follow-up
  call is needed.
- A refusal prints the reason and a "next:" hint. Exit 0 = done,
  1 = refused (nothing saved), 2 = bad command line. --json prints one JSON
  object instead of text.
"""


class _Refusal(Exception):
    """A CLI-level refusal that is not a library error (e.g. an unknown revision id)."""

    def __init__(self, message: str, next: str | None = None):
        super().__init__(message)
        self.next = next


def _next_hint(error: Exception, args: argparse.Namespace) -> str | None:
    """The error's ``next`` hint in the CLI's own spelling, where a CLI verb does the job."""
    file = shlex.quote(args.file)
    if isinstance(error, _Refusal):
        return error.next
    if isinstance(error, WorkspaceSyncError):
        return "add --force-recreate (discards the leftover workspace and any unsaved edits it holds)"
    if isinstance(error, WorkspaceLockedError):
        return "close the session holding this file (doc.close() or docx-session stop), then run this again"
    if isinstance(error, DocumentProtectedError):
        return "add --allow-protected (the protection stays in place in the saved file)"
    hint = getattr(error, "next", None)
    if not hint:
        return None
    if hint.startswith("retry with occurrence="):
        return "add --occurrence " + hint.removeprefix("retry with occurrence=")
    if hint.startswith("retry with author="):
        return "add --by NAME (one of the authors named above)"
    if hint.startswith("find_all(") and isinstance(error, MatchCountError):
        return f"{PROG} find {file} {shlex.quote(error.search_text)}"
    if match := re.match(r"get_paragraph\((\d+)\)", hint):
        return f"{PROG} list {file} --start {match[1]} --limit 1"
    if "list_paragraphs()" in hint:
        return f"{PROG} list {file}"
    if hint == "list_revisions()":
        return f"{PROG} revisions {file}"
    return f"(Python API) {hint}"


def _destination(args: argparse.Namespace) -> Path:
    """Where a mutating verb saves: the input is never touched unless asked.

    ``--in-place`` and ``-o`` decide when given. Otherwise a ``*.edited.docx``
    input is updated in place, and any other input writes
    ``<stem>.edited.docx`` beside it — refused when that file already exists,
    so a different file is never overwritten silently.
    """
    source = Path(args.file)
    if args.in_place:
        return source
    if args.output:
        return Path(args.output)
    if source.name.lower().endswith(_EDITED):
        return source
    target = source.with_name(source.stem + _EDITED)
    if target.exists():
        raise _Refusal(
            f"{target} already exists; nothing was written. To keep editing that redline, run the command "
            f"on {target} itself; to write elsewhere pass -o OUT; to edit {source} itself pass --in-place.",
            next=f"{PROG} {args.verb} {shlex.quote(str(target))} ...",
        )
    return target


def _paragraphs(doc: Document, refs: Sequence[str]) -> list[dict]:
    """``{"ref", "markup"}`` for each distinct ref: the paragraph as it now reads."""
    markup = doc._revision_manager.markup_lines()
    out = []
    for ref in dict.fromkeys(refs):
        index = int(ref[1:].split("#")[0])
        out.append({"ref": ref, "markup": markup[index - 1] if 0 < index <= len(markup) else ""})
    return out


def _run_edit(doc: Document, args: argparse.Namespace) -> dict:
    scope = {"paragraph": args.paragraph, "occurrence": args.occurrence}
    results: list[EditResult]
    if args.verb == "replace" and args.expect is not None:
        results = doc.replace_all(args.old, args.new, expected_count=args.expect, note=args.note)
    elif args.verb == "replace":
        results = [doc.replace(args.old, args.new, note=args.note, **scope)]
    elif args.verb == "delete":
        results = [doc.delete(args.text, note=args.note, **scope)]
    elif args.verb == "insert-after":
        results = [doc.insert_after(args.anchor, args.text, note=args.note, **scope)]
    else:
        results = [doc.insert_before(args.anchor, args.text, note=args.note, **scope)]
    return {
        "results": [
            {"ref": str(r), "refs": list(r.refs), "revision_ids": list(r.revision_ids), "comment_id": r.comment_id}
            for r in results
        ],
        "paragraphs": _paragraphs(doc, [ref for r in results for ref in (r.refs or (str(r),))]),
    }


def _run_comment(doc: Document, args: argparse.Namespace) -> dict:
    comment_id = doc.add_comment(args.anchor, args.text, paragraph=args.paragraph, occurrence=args.occurrence)
    ref = next((c.paragraph_ref for c in doc.list_comments() if c.id == comment_id), None)
    return {"comment_id": comment_id, "ref": ref, "paragraphs": _paragraphs(doc, [ref] if ref else [])}


def _run_resolve(doc: Document, args: argparse.Namespace) -> dict:
    if args.all:
        resolved = (doc.accept_all if args.verb == "accept" else doc.reject_all)(author=args.by)
        return {
            "resolved": int(resolved),
            "unhandled": resolved.unhandled,
            "unhandled_types": dict(resolved.unhandled_types),
            "pending": len(doc.list_revisions()),
        }
    # The touched paragraphs, as elements: resolving a paragraph mark or a
    # move merges or removes paragraphs, so an index taken before would name
    # another paragraph afterwards.
    before = body_paragraphs(doc._document_editor.dom)
    where = {r.id: r.paragraph_ref for r in doc.list_revisions(author=args.by)}
    touched = list(dict.fromkeys(before[int(ref[1:].split("#")[0]) - 1] for i in args.ids if (ref := where.get(i))))
    one = doc.accept_revision if args.verb == "accept" else doc.reject_revision
    for rev_id in args.ids:
        if not one(rev_id, author=args.by):
            by = f" by {args.by}" if args.by else ""
            raise _Refusal(
                f"No pending revision {rev_id}{by} (it may have been resolved together with an earlier id). "
                f"Nothing was saved.",
                next=f"{PROG} revisions {shlex.quote(args.file)}",
            )
    # Each touched paragraph that still exists, under its current ref
    # (resolving changed its text); one merged away is counted, not guessed.
    after = body_paragraphs(doc._document_editor.dom)
    position = {id(p): i for i, p in enumerate(after, start=1)}
    refs = {p.index: p.ref for p in doc.list_paragraphs_structured(limit=None)} if touched else {}
    current = [refs[position[id(p)]] for p in touched if id(p) in position]
    out = {
        "resolved": len(args.ids),
        "ids": list(args.ids),
        "pending": len(doc.list_revisions()),
        "paragraphs": _paragraphs(doc, current),
    }
    if gone := len(touched) - len(current):
        out["merged_paragraphs"] = gone
    return out


def _run_find(doc: Document, args: argparse.Namespace) -> dict:
    matches = doc.find_all(args.text, args.paragraph, expected_count=args.expect)
    texts = {p.index: p.text for p in doc.list_paragraphs_structured(limit=None)} if matches else {}
    return {
        "matches": [
            {
                "occurrence": i,
                "paragraph_ref": m.paragraph_ref,
                "paragraph_occurrence": m.paragraph_occurrence,
                "start": m.start,
                "context": _match_context(texts.get(m.paragraph_index, ""), m.start, m.end),
            }
            for i, m in enumerate(matches)
        ],
    }


def _run_list(doc: Document, args: argparse.Namespace) -> dict:
    kwargs: dict = {}
    if args.start is not None:
        kwargs["start"] = args.start
    if args.max_chars is not None:
        kwargs["max_chars"] = args.max_chars
    if args.limit is not None:
        kwargs["limit"] = args.limit or None
    lines = doc.list_paragraphs(**kwargs)
    # The library's truncation notice names Python keywords; name the CLI
    # flags instead. Paragraph lines start with their ref ("P3#..").
    if lines and lines[-1].startswith("... "):
        next_start = kwargs.get("start", 1) + len(lines) - 1
        remaining = doc.paragraph_count() - next_start + 1
        lines[-1] = f"... {remaining} more paragraphs; next page: --start {next_start} (--limit 0 lists all)"
    elif not lines:
        lines = [f"no paragraphs from {kwargs.get('start', 1)} (the document has {doc.paragraph_count()})"]
    return {"lines": lines}


def _run_revisions(doc: Document, args: argparse.Namespace) -> dict:
    return {"revisions": doc.list_revisions(author=args.by, paragraph=args.paragraph)}


_RUNNERS = {
    **dict.fromkeys(_EDIT_VERBS, _run_edit),
    "comment": _run_comment,
    "accept": _run_resolve,
    "reject": _run_resolve,
    "find": _run_find,
    "list": _run_list,
    "revisions": _run_revisions,
}


def _print_text(verb: str, out: dict) -> None:
    if verb == "find":
        print(f"{len(out['matches'])} match(es):")
        for m in out["matches"]:
            print(
                f'  occurrence={m["occurrence"]}  {m["paragraph_ref"]} at char {m["start"]}  "{m["context"]}"'
                f"  (--paragraph {m['paragraph_ref']} --occurrence {m['paragraph_occurrence']})"
            )
    elif verb == "list":
        print("\n".join(out["lines"]))
    elif verb == "revisions":
        if not out["revisions"]:
            print("no pending revisions")
        for r in out["revisions"]:
            print(f'#{r.id} {r.type} by {r.author} in {r.paragraph_ref}: "{r.text}"')
    elif verb in ("accept", "reject"):
        line = f"{verb}ed {out['resolved']} revision(s); {out['pending']} pending"
        if out.get("merged_paragraphs"):
            line += f"; {out['merged_paragraphs']} touched paragraph(s) merged into a neighbour"
        if out.get("unhandled"):
            kinds = ", ".join(f"{k}={v}" for k, v in out["unhandled_types"].items())
            line += f"; {out['unhandled']} left that this tool does not resolve ({kinds})"
        print(line)
    elif verb == "comment":
        print(f"comment {out['comment_id']} on {out['ref']}")
    else:
        if not out["results"]:
            print(f"{verb}: no match, nothing changed")
        for r in out["results"]:
            details = [
                "revisions " + ", ".join(map(str, r["revision_ids"]))
                if r["revision_ids"]
                else "no new revision (amended your own insertion, or nothing changed)"
            ]
            if r["comment_id"] is not None:
                details.append(f"comment {r['comment_id']}")
            print(f"{verb} {' '.join(r['refs'] or [r['ref']])}  ({'; '.join(details)})")
    for p in out.get("paragraphs", []):
        print(f"  {p['ref']}| {p['markup']}")
    for w in out.get("warnings", []):
        print(f"warning: {w}")
    if out.get("saved"):
        print(f"saved {out['saved']}")


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=PROG,
        allow_abbrev=False,
        description="One tracked-change edit to a .docx per command. `docx-edit skill` prints the agent guide.",
    )
    verbs = parser.add_subparsers(dest="verb", required=True, metavar="VERB")

    def verb(name: str, summary: str, example: str) -> argparse.ArgumentParser:
        sub = verbs.add_parser(
            name,
            allow_abbrev=False,
            help=summary,
            description=summary,
            epilog=f"example:\n  {example}",
            formatter_class=argparse.RawDescriptionHelpFormatter,
        )
        if name != "skill":
            sub.add_argument("file", help="the .docx file")
            sub.add_argument("--json", action="store_true", help="print one JSON object instead of text")
        return sub

    def scope(sub: argparse.ArgumentParser) -> None:
        sub.add_argument("--paragraph", metavar="REF", help="pin to this paragraph ref (e.g. P3#a7b2)")
        sub.add_argument("--occurrence", type=int, metavar="N", help="pick match N (0-based) when text repeats")

    def output(sub: argparse.ArgumentParser, author_help: str = "name recorded on the tracked change") -> None:
        sub.add_argument("--author", required=True, help=author_help)
        where = sub.add_mutually_exclusive_group()
        where.add_argument(
            "-o",
            "--output",
            metavar="OUT",
            help="save to OUT (default: FILE's <stem>.edited.docx; an input named *.edited.docx is updated in place)",
        )
        where.add_argument("--in-place", action="store_true", help="overwrite FILE")
        sub.add_argument("--allow-protected", action="store_true", help="edit a document under Restrict Editing")
        sub.add_argument(
            "--force-recreate", action="store_true", help="discard a workspace a crashed session left behind"
        )

    def note(sub: argparse.ArgumentParser) -> None:
        sub.add_argument("--note", help="rationale, attached as a comment on the change")

    sub = verb(
        "replace", "Replace text as a tracked change.", 'docx-edit replace F.docx "30 days" "60 days" --author Me'
    )
    sub.add_argument("old")
    sub.add_argument("new")
    scope(sub)
    sub.add_argument("--expect", type=int, metavar="N", help="replace every match; refuse unless there are exactly N")
    note(sub)
    output(sub)

    sub = verb("delete", "Delete text as a tracked change.", 'docx-edit delete F.docx "unless agreed" --author Me')
    sub.add_argument("text")
    scope(sub)
    note(sub)
    output(sub)

    for name, where in (("insert-after", "after"), ("insert-before", "before")):
        sub = verb(
            name,
            f"Insert text {where} an anchor as a tracked change.",
            f'docx-edit {name} F.docx "Payment" " terms" --author Me',
        )
        sub.add_argument("anchor")
        sub.add_argument("text")
        scope(sub)
        note(sub)
        output(sub)

    sub = verb("comment", "Comment on text.", 'docx-edit comment F.docx "Effective Date" "Check this" --author Me')
    sub.add_argument("anchor")
    sub.add_argument("text", metavar="comment")
    scope(sub)
    output(sub, "name recorded on the comment")

    sub = verb("find", "List every match of text, with refs.", 'docx-edit find F.docx "30 days"')
    sub.add_argument("text")
    sub.add_argument("--paragraph", metavar="REF", help="search only this paragraph")
    sub.add_argument("--expect", type=int, metavar="N", help="refuse unless there are exactly N matches")

    sub = verb("list", "List paragraphs with their refs.", "docx-edit list F.docx --start 40 --limit 20")
    sub.add_argument("--start", type=int, help="first paragraph (1-based; default 1)")
    sub.add_argument("--limit", type=int, help="most paragraphs to list (0 = all; default: the library's page size)")
    sub.add_argument("--max-chars", type=int, help="cut each paragraph's preview at this length")

    sub = verb("revisions", "List pending tracked changes.", "docx-edit revisions F.docx --by Bot")
    sub.add_argument("--by", metavar="NAME", help="only revisions by this author")
    sub.add_argument("--paragraph", metavar="REF", help="only revisions in this paragraph")

    for name in ("accept", "reject"):
        sub = verb(
            name,
            f"{name.capitalize()} tracked changes by id, or all of them.",
            f"docx-edit {name} F.edited.docx 5 6 --author Me",
        )
        sub.add_argument("ids", nargs="*", type=int, metavar="ID", help="revision ids from `revisions`")
        sub.add_argument("--all", action="store_true", help="every pending revision (with --by: only that author's)")
        sub.add_argument("--by", metavar="NAME", help="only revisions by this author")
        output(sub, "who is resolving")

    verb("skill", "Print the agent guide for this tool.", "docx-edit skill")
    return parser


def _check_usage(parser: argparse.ArgumentParser, args: argparse.Namespace) -> None:
    if args.verb == "replace" and args.expect is not None and (args.paragraph or args.occurrence is not None):
        parser.error("replace: --expect replaces every match in the document; drop --paragraph/--occurrence")
    if args.verb in ("accept", "reject") and bool(args.ids) == bool(args.all):
        parser.error(f"{args.verb}: give revision ids or --all, not both or neither")


def _open(args: argparse.Namespace, mutating: bool, scratch: str) -> Document:
    """Open ``args.file`` the way this verb needs.

    A read-only verb gets a private workspace under ``scratch``: it never
    saves, so it must not lock, adopt or destroy the shared one a session or
    an earlier crash left behind, and a protected file may be read. A
    mutating verb uses the shared workspace, so a live session's lock still
    refuses; ``--force-recreate`` discards a leftover workspace only when
    opening without it reports one (WorkspaceSyncError), never a live lock.
    """
    if not mutating:
        return Document.open(args.file, workspace_dir=scratch, allow_protected=True)
    kwargs = {"author": args.author, "allow_protected": args.allow_protected}
    try:
        return Document.open(args.file, **kwargs)
    except WorkspaceSyncError:
        if not args.force_recreate:
            raise
        return Document.open(args.file, force_recreate=True, **kwargs)


def _emit(text_or_json: str, *, stream=None) -> None:
    """Print, surviving a reader that closed the pipe: the exit status stays the command's."""
    stream = stream or sys.stdout
    try:
        print(text_or_json, file=stream)
        stream.flush()
    except BrokenPipeError:
        devnull = os.open(os.devnull, os.O_WRONLY)
        os.dup2(devnull, sys.stdout.fileno())
        os.dup2(devnull, sys.stderr.fileno())


def _refusal_output(error: Exception, args: argparse.Namespace) -> None:
    hint = _next_hint(error, args)
    if args.json:
        payload = (
            {"type": "Refusal", "message": str(error), "next": error.next}
            if isinstance(error, _Refusal)
            else _error_payload(error)
        )
        _emit(json.dumps({"ok": False, "error": payload, "next": hint}, default=_json_default))
        return
    _emit(f"error: {error}" + (f"\nnext: {hint}" if hint else ""), stream=sys.stderr)


def main(argv: Sequence[str] | None = None) -> int:
    """Run one docx-edit command; returns 0 (done) or 1 (refused, nothing saved).

    A bad command line exits with status 2 through argparse (SystemExit).
    """
    # An output encoding that cannot show a character must not turn a saved
    # edit into a traceback and exit 1 ("nothing saved").
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(errors="backslashreplace")
    parser = _parser()
    args = parser.parse_args(list(sys.argv[1:] if argv is None else argv))
    if args.verb == "skill":
        print(AGENT_GUIDE, end="")
        return 0
    _check_usage(parser, args)

    mutating = args.verb in _MUTATING
    cleanup_warnings: list[str] = []
    try:
        destination = _destination(args) if mutating else None
        # A read-only verb's private workspace; an edit needs none. A failure
        # to remove it is ignored: it never holds anything the caller wants.
        scratch_dir = (
            contextlib.nullcontext("") if mutating else tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        )
        with warnings.catch_warnings(record=True) as caught, scratch_dir as scratch:
            # Only the library's own warnings are the caller's business; a
            # ResourceWarning or DeprecationWarning from a dependency is not.
            for category in _LIBRARY_WARNINGS:
                warnings.simplefilter("always", category)
            doc = _open(args, mutating, scratch)
            out: dict = {}
            try:
                out = _RUNNERS[args.verb](doc, args)
                if destination is not None:
                    out["saved"] = str(doc.save(destination))
            finally:
                try:
                    # A read-only verb's private workspace goes with its
                    # temporary directory, whose removal errors are ignored.
                    doc.close(cleanup=mutating)
                except OSError as error:
                    # After a save, a workspace that cannot be removed (a file
                    # held open by a scanner) is not a refusal: the edit is done.
                    if "saved" not in out:
                        raise
                    cleanup_warnings.append(f"saved, but the workspace could not be removed: {error}")
    except (DocxEditError, ValueError, OSError, _Refusal) as error:
        _refusal_output(error, args)
        return 1

    found = (str(w.message) for w in caught if issubclass(w.category, _LIBRARY_WARNINGS))
    out["warnings"] = list(dict.fromkeys([*found, *cleanup_warnings]))
    if args.json:
        _emit(json.dumps({"ok": True, "verb": args.verb, **out}, default=_json_default))
    else:
        try:
            _print_text(args.verb, out)
            sys.stdout.flush()
        except BrokenPipeError:
            # The reader went away; the command itself succeeded (and saved).
            os.dup2(os.open(os.devnull, os.O_WRONLY), sys.stdout.fileno())
    return 0


if __name__ == "__main__":
    sys.exit(main())
