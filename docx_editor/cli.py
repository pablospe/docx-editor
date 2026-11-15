"""docx-edit: one tracked-change edit per command (ROADMAP.md #104).

Each verb is one :class:`~docx_editor.Document` method, named in kebab-case,
with its keywords as flags: ``docx-edit replace F OLD NEW --paragraph P`` is
``doc.replace(OLD, NEW, paragraph=P)``. A command opens the file, makes the
one call, saves (mutating verbs only, to ``-o OUT`` or ``--in-place``) and
closes. Nothing is saved when the call refuses.

Output is a receipt the caller can stop on: what changed, the paragraph ref,
the paragraph as it now reads (tracked-change markup), and where it was saved.
A refusal prints the library's message (with its pick-list) and, when one
command fixes it, a ``retry:`` line holding that command. ``--json`` prints
the same as one JSON object on stdout. Exit status: 0 success, 1 refused,
2 usage error.
"""

import argparse
import contextlib
import json
import os
import shlex
import sys
import tempfile
import warnings
from collections.abc import Sequence
from dataclasses import dataclass

from . import Document, EditResult
from .exceptions import (
    AmbiguousTextError,
    DocumentProtectedError,
    DocxEditError,
    HashMismatchError,
    MatchCountError,
    ParagraphIndexError,
    RevisionError,
    TextNotFoundError,
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
# Positional arguments of each verb after FILE, in order (argparse dests).
_POSITIONALS = {
    "replace": ("old", "new"),
    "delete": ("text",),
    "insert-after": ("anchor", "text"),
    "insert-before": ("anchor", "text"),
    "comment": ("anchor", "text"),
    "find": ("text",),
    "accept": ("ids",),
    "reject": ("ids",),
}
# Flags a retry carries over, in the order they are written (argparse dest, flag).
_FLAGS = (
    ("paragraph", "--paragraph"),
    ("occurrence", "--occurrence"),
    ("expect", "--expect"),
    ("note", "--note"),
    ("by", "--by"),
    ("start", "--start"),
    ("limit", "--limit"),
    ("max_chars", "--max-chars"),
    ("author", "--author"),
    ("output", "--output"),
)
_SWITCHES = (
    ("all", "--all"),
    ("in_place", "--in-place"),
    ("allow_protected", "--allow-protected"),
)

AGENT_GUIDE = """\
docx-edit: tracked-change edits to a .docx, one command each. Every edit is a
Word tracked change by --author; nothing is saved unless the command succeeds.

Look:
  docx-edit list F.docx                      paragraphs as "P3#a7b2| text"
  docx-edit find F.docx "30 days"            every match, with its ref
  docx-edit revisions F.docx [--by NAME]     pending tracked changes

Edit (each needs --author NAME and -o OUT.docx or --in-place):
  docx-edit replace F.docx "30 days" "60 days" --author Me -o out.docx
  docx-edit delete F.docx "unless agreed" --author Me --in-place
  docx-edit insert-after F.docx "Payment" " terms" --author Me --in-place
  docx-edit insert-before F.docx "due" "now " --author Me --in-place
  docx-edit comment F.docx "Effective Date" "Check this" --author Me --in-place
  docx-edit replace F.docx "30 days" "60 days" --expect 3 --author Me --in-place
  docx-edit accept F.docx 5 6 --author Me --in-place   (or --all [--by NAME])
  docx-edit reject F.docx --all --by Bot --author Me --in-place

Rules:
- Without --paragraph the text is searched in the whole document and must
  occur once. If it occurs more often the command refuses, lists up to 10
  matches as "occurrence=N  P3#a7b2 ..." (find lists all), and prints a
  retry: line; run it with N replaced. --paragraph P3#a7b2 (a ref from
  list/find) pins the edit to that paragraph as you saw it; a stale ref
  refuses and the retry carries the fresh one.
- --expect N on replace replaces every match, and refuses unless there are
  exactly N. On find it checks the count without editing.
- --note "why" attaches your rationale as a comment on the change.
- A newline in the new text splits the paragraph (tracked), e.g. $'a\\nb'.
- Text starting with "-": put -- before the positional arguments, as in
  docx-edit replace --author Me --in-place -- F.docx "- item" "- entry".
- A moved paragraph is two revisions (move_from + move_to): accept or reject
  both ids together, or use --all.
- The receipt shows each changed paragraph as it now reads:
  [del#5:Me]old[/del][ins#6:Me]new[/ins]. It is the check; no follow-up
  call is needed.
- Exit 0 = done, 1 = refused (nothing saved), 2 = bad command line.
  --json prints one JSON object instead of text.
"""


@dataclass
class _Retry:
    """The one command that resolves a refusal, and what to fill in or know first."""

    argv: list[str]
    note: str | None = None


class _Refusal(Exception):
    """A CLI-level refusal that is not a library error (e.g. an unknown revision id)."""

    next: str | None = None

    def __init__(self, message: str, retry: _Retry | None = None):
        super().__init__(message)
        self.retry = retry


def _command(args: argparse.Namespace, *switches: str, **overrides: object) -> list[str]:
    """The command ``args`` came from, with ``overrides`` applied, as an argv.

    Rebuilt from the parsed arguments, not edited from the raw argv, so a
    repeated flag or a ``--`` cannot make the retry differ from what argparse
    actually ran. Valued flags are one ``--flag=value`` token, so a value
    starting with ``-`` stays a value; positionals follow ``--``, so text
    starting with ``-`` stays text. ``switches`` (e.g. ``--force-recreate``)
    are added for this retry only and never carried over.
    """
    values = {**vars(args), **overrides}
    argv = [PROG, args.verb]
    for dest, flag in _FLAGS:
        if values.get(dest) is not None:
            argv.append(f"{flag}={values[dest]}")
    for dest, flag in _SWITCHES:
        if values.get(dest):
            argv.append(flag)
    argv += switches
    if values.get("json"):
        argv.append("--json")
    argv += ["--", values["file"]]
    for dest in _POSITIONALS.get(args.verb, ()):
        value = values[dest]
        argv += [str(v) for v in value] if isinstance(value, list) else [str(value)]
    return argv


def _other(args: argparse.Namespace, verb: str, *argv: str) -> list[str]:
    """Another verb on the same file, keeping ``--json`` when the caller used it."""
    return [PROG, verb, *argv, *(["--json"] if args.json else []), "--", args.file]


_ANSI_C_ESCAPES = {"\\": "\\\\", "'": "\\'", "\n": "\\n", "\r": "\\r", "\t": "\\t"}


def _quote(arg: str) -> str:
    """One shell word. An argument with a control character (a newline in a
    split) uses bash's ``$'...'`` form, so the retry stays on one line."""
    if arg.isprintable():
        return shlex.quote(arg)
    body = []
    for c in arg:
        if c in _ANSI_C_ESCAPES:
            body.append(_ANSI_C_ESCAPES[c])
        elif c.isprintable():
            body.append(c)
        else:
            # \x only for ASCII: bash writes \xHH as a raw byte, so U+00A0
            # needs \u00a0 to come out as UTF-8; above the BMP, \U.
            code = ord(c)
            body.append(f"\\x{code:02x}" if code < 0x80 else f"\\u{code:04x}" if code <= 0xFFFF else f"\\U{code:08x}")
    return "$'" + "".join(body) + "'"


def _render(argv: Sequence[str]) -> str:
    """``argv`` as one POSIX-shell command line (bash, zsh, Git Bash).

    Not cmd.exe or PowerShell, whose quoting differs; ``--json`` gives the
    exact ``retry_argv`` for any other caller.
    """
    return " ".join(_quote(a) for a in argv)


def _retry_for(error: Exception, args: argparse.Namespace) -> _Retry | None:
    """The command to run next after ``error``, or None when no single command fixes it."""
    if isinstance(error, _Refusal):
        return error.retry
    if isinstance(error, AmbiguousTextError):
        return _Retry(_command(args, occurrence="N"), "N: an occurrence= from the list above")
    if isinstance(error, MatchCountError):
        if args.verb == "find":
            return None
        if error.total_occurrences == 0:
            return _Retry(_other(args, "list"), "nothing matches; read the paragraphs to quote the text")
        if shared := _split_conflict(args, [m.paragraph_ref for m in error.matches]):
            return _one_at_a_time(args, shared)
        return _Retry(_command(args, expect=error.total_occurrences), "only if every match listed above should change")
    if isinstance(error, HashMismatchError) and getattr(args, "paragraph", None):
        return _Retry(_command(args, paragraph=f"P{error.paragraph_index}#{error.actual_hash}"))
    if isinstance(error, TextNotFoundError):
        if error.total_occurrences:
            return _Retry(_command(args, occurrence="N"), f"N: 0..{error.total_occurrences - 1}")
        if getattr(args, "paragraph", None):
            target = getattr(args, "old", None) or getattr(args, "anchor", None) or args.text
            return _Retry([*_other(args, "find"), target], "not in that paragraph; find lists where it is")
        return _Retry(_other(args, "list"))
    if isinstance(error, ParagraphIndexError):
        return _Retry(_other(args, "list"))
    if isinstance(error, ValueError) and str(error).startswith("Invalid paragraph reference"):
        return _Retry(_other(args, "list"), "--paragraph takes a ref like P3#a7b2, as list prints it")
    if isinstance(error, ValueError) and str(error).startswith((
        "replace_all(): replace_with contains a newline",
        "replace_all(): matches of",
    )):
        return _one_at_a_time(args)
    if isinstance(error, WorkspaceLockedError):
        return _Retry(
            _command(args), "another session holds this file: close it (doc.close() / docx-session stop), then run this"
        )
    if isinstance(error, RevisionError) and (error.next or "").startswith("retry with author="):
        return _Retry(_command(args, by="NAME"), "NAME: one of the authors named above")
    if isinstance(error, RevisionError) and error.next == "list_revisions()":
        return _Retry(_other(args, "revisions"))
    if isinstance(error, DocumentProtectedError):
        return _Retry(_command(args, allow_protected=True), "the protection stays in place in the saved file")
    if isinstance(error, WorkspaceSyncError):
        return _Retry(
            _command(args, "--force-recreate"), "discards the leftover workspace and any unsaved edits it holds"
        )
    return None


def _split_conflict(args: argparse.Namespace, refs: Sequence[str]) -> str | None:
    """The paragraph a ``--expect`` replace with a newline would split twice, if any."""
    if args.verb != "replace" or "\n" not in args.new:
        return None
    return next((ref for i, ref in enumerate(refs) if ref in refs[:i]), None)


def _one_at_a_time(args: argparse.Namespace, ref: str | None = None) -> _Retry:
    """Matches that cannot go in one call: list them, then edit each in its own command."""
    scope = [f"--paragraph={ref}"] if ref else []
    then = (
        f"; after the first, continue on {args.output} with --in-place, since each run reads its FILE"
        if args.output
        else ""
    )
    return _Retry(
        [*_other(args, "find", *scope), args.old],
        "these matches cannot all change in one command; find prints each one's --paragraph and --occurrence: "
        f"replace them one per command, without --expect{then}",
    )


def _paragraphs(doc: Document, refs: Sequence[str]) -> list[dict]:
    """``{"ref", "markup"}`` for each distinct ref: the paragraph as it now reads."""
    markup = doc._revision_manager.markup_lines()
    out = []
    for ref in dict.fromkeys(refs):
        index = int(ref[1:].split("#")[0])
        out.append({"ref": ref, "markup": markup[index - 1] if 0 < index <= len(markup) else ""})
    return out


def _current_refs(doc: Document, indexes: Sequence[int]) -> list[str]:
    """The current ref of each paragraph index (refs change when a paragraph's text does)."""
    if not indexes:
        return []
    refs = {p.index: p.ref for p in doc.list_paragraphs_structured(limit=None)}
    return [refs[i] for i in dict.fromkeys(indexes) if i in refs]


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
    elif args.verb == "insert-before":
        results = [doc.insert_before(args.anchor, args.text, note=args.note, **scope)]
    else:
        raise AssertionError(args.verb)
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
    touched = [before[int(ref[1:].split("#")[0]) - 1] for i in args.ids if (ref := where.get(i))]
    one = doc.accept_revision if args.verb == "accept" else doc.reject_revision
    for rev_id in args.ids:
        if not one(rev_id, author=args.by):
            by = f" by {args.by}" if args.by else ""
            raise _Refusal(
                f"No pending revision {rev_id}{by} (it may have been resolved together with an earlier id). "
                f"Nothing was saved.",
                _Retry(_other(args, "revisions")),
            )
    # Report each touched paragraph that still exists under its current ref
    # (resolving changed its text); one merged away is counted, not guessed.
    after = {id(p): i for i, p in enumerate(body_paragraphs(doc._document_editor.dom), start=1)}
    indexes = [after[id(p)] for p in dict.fromkeys(touched) if id(p) in after]
    out = {
        "resolved": len(args.ids),
        "ids": list(args.ids),
        "pending": len(doc.list_revisions()),
        "paragraphs": _paragraphs(doc, _current_refs(doc, indexes)),
    }
    if gone := len(dict.fromkeys(touched)) - len(indexes):
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
    # The library's truncation notice names Python keywords; say it as the
    # next command instead. Paragraph lines start with their ref ("P3#..").
    if lines and lines[-1].startswith("... "):
        next_start = kwargs.get("start", 1) + len(lines) - 1
        remaining = doc.paragraph_count() - next_start + 1
        page = _render(_command(args, start=next_start))
        lines[-1] = f"... {remaining} more paragraphs; next page: {page} (--limit=0 lists all)"
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
            sub.add_argument("--force-recreate", action="store_true", help=argparse.SUPPRESS)
        return sub

    def scope(sub: argparse.ArgumentParser) -> None:
        sub.add_argument("--paragraph", metavar="REF", help="pin to this paragraph ref (e.g. P3#a7b2)")
        sub.add_argument("--occurrence", type=int, metavar="N", help="pick match N (0-based) when text repeats")

    def output(sub: argparse.ArgumentParser, author_help: str = "name recorded on the tracked change") -> None:
        sub.add_argument("--author", required=True, help=author_help)
        where = sub.add_mutually_exclusive_group(required=True)
        where.add_argument("-o", "--output", metavar="OUT", help="save to OUT")
        where.add_argument("--in-place", action="store_true", help="overwrite FILE")
        sub.add_argument("--allow-protected", action="store_true", help=argparse.SUPPRESS)

    def note(sub: argparse.ArgumentParser) -> None:
        sub.add_argument("--note", help="rationale, attached as a comment on the change")

    sub = verb(
        "replace",
        "Replace text as a tracked change.",
        'docx-edit replace F.docx "30 days" "60 days" --author Me -o out.docx',
    )
    sub.add_argument("old")
    sub.add_argument("new")
    scope(sub)
    sub.add_argument("--expect", type=int, metavar="N", help="replace every match; refuse unless there are exactly N")
    note(sub)
    output(sub)

    sub = verb(
        "delete", "Delete text as a tracked change.", 'docx-edit delete F.docx "unless agreed" --author Me --in-place'
    )
    sub.add_argument("text")
    scope(sub)
    note(sub)
    output(sub)

    for name, where in (("insert-after", "after"), ("insert-before", "before")):
        sub = verb(
            name,
            f"Insert text {where} an anchor as a tracked change.",
            f'docx-edit {name} F.docx "Payment" " terms" --author Me --in-place',
        )
        sub.add_argument("anchor")
        sub.add_argument("text")
        scope(sub)
        note(sub)
        output(sub)

    sub = verb(
        "comment", "Comment on text.", 'docx-edit comment F.docx "Effective Date" "Check this" --author Me --in-place'
    )
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
            f"docx-edit {name} F.docx 5 6 --author Me --in-place",
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


def _refusal_output(error: Exception, args: argparse.Namespace) -> None:
    retry = _retry_for(error, args)
    if args.json:
        payload = (
            {"type": "Refusal", "message": str(error), "next": None}
            if isinstance(error, _Refusal)
            else _error_payload(error)
        )
        body = {
            "ok": False,
            "error": payload,
            "retry": _render(retry.argv) if retry else None,
            "retry_argv": retry.argv if retry else None,
            "retry_note": retry.note if retry else None,
        }
        print(json.dumps(body, default=_json_default))
        return
    print(f"error: {error}", file=sys.stderr)
    if retry:
        if retry.note:
            print(f"note: {retry.note}", file=sys.stderr)
        print(f"retry: {_render(retry.argv)}", file=sys.stderr)
    elif (hint := getattr(error, "next", None)) and args.verb != "find":
        print(f"next (Python API): {hint}", file=sys.stderr)


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
    library_warnings = (UnanchoredNoteWarning, UnhandledRevisionWarning)
    cleanup_warnings: list[str] = []
    try:
        # A read-only verb's private workspace; an edit needs none. A failure
        # to remove it is ignored: it never holds anything the caller wants.
        scratch_dir = (
            contextlib.nullcontext("") if mutating else tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        )
        with warnings.catch_warnings(record=True) as caught, scratch_dir as scratch:
            # Only the library's own warnings are the caller's business; a
            # ResourceWarning or DeprecationWarning from a dependency is not.
            for category in library_warnings:
                warnings.simplefilter("always", category)
            doc = _open(args, mutating, scratch)
            out: dict = {}
            try:
                out = _RUNNERS[args.verb](doc, args)
                if mutating:
                    out["saved"] = str(doc.save(args.file if args.in_place else args.output))
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
        try:
            _refusal_output(error, args)
            sys.stdout.flush()
        except BrokenPipeError:
            os.dup2(os.open(os.devnull, os.O_WRONLY), sys.stdout.fileno())
        return 1

    found = (str(w.message) for w in caught if issubclass(w.category, library_warnings))
    out["warnings"] = list(dict.fromkeys([*found, *cleanup_warnings]))
    try:
        if args.json:
            print(json.dumps({"ok": True, "verb": args.verb, **out}, default=_json_default))
        else:
            _print_text(args.verb, out)
        sys.stdout.flush()
    except BrokenPipeError:
        # The reader went away; the command itself succeeded (and saved).
        # Point stdout at devnull so the interpreter's exit flush cannot fail.
        os.dup2(os.open(os.devnull, os.O_WRONLY), sys.stdout.fileno())
    return 0


if __name__ == "__main__":
    sys.exit(main())
