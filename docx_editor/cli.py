"""docx-edit: one tracked-change edit per command (ROADMAP.md #104).

Each verb is one :class:`~docx_editor.Document` method, named in kebab-case,
with its keywords as flags: ``docx-edit replace F OLD NEW --paragraph P`` is
``doc.replace(OLD, NEW, paragraph=P)``. A command opens the file, makes the
one call, saves (mutating verbs only, to ``-o OUT`` or ``--in-place``) and
closes. Nothing is saved when the call refuses.

Output is a receipt the caller can stop on: what changed, the paragraph ref,
the paragraph as it now reads (tracked-change markup), and where it was saved.
A refusal prints the library's message (with its pick-list) and a ``retry:``
line that is the command to run next. ``--json`` prints the same as one JSON
object on stdout. Exit status: 0 success, 1 refused, 2 usage error.
"""

import argparse
import json
import shlex
import sys
import warnings
from collections.abc import Sequence

from . import Document, EditResult
from .exceptions import (
    AmbiguousTextError,
    DocxEditError,
    HashMismatchError,
    MatchCountError,
    TextNotFoundError,
    UnanchoredNoteWarning,
    UnhandledRevisionWarning,
    _error_payload,
    _json_default,
    _match_context,
)

PROG = "docx-edit"

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
  occur once. If it occurs more often the command refuses, lists every match
  as "occurrence=N  P3#a7b2 ...", and prints a retry: line; rerun it with
  --occurrence N. --paragraph P3#a7b2 (a ref from list/find) pins the edit to
  that paragraph as you saw it; a stale ref refuses and prints the fresh one.
- --expect N on replace replaces every match, and refuses unless there are
  exactly N. On find it checks the count without editing.
- --note "why" attaches your rationale as a comment on the change.
- The receipt shows the paragraph after the edit: [del#5:Me]old[/del]
  [ins#6:Me]new[/ins]. It is the check; no follow-up call is needed.
- Exit 0 = done, 1 = refused (nothing saved), 2 = bad command line.
  --json prints one JSON object instead of text.
"""


class _Refusal(Exception):
    """A CLI-level refusal that is not a library error (e.g. an unknown revision id)."""

    def __init__(self, message: str, retry: str | None = None):
        super().__init__(message)
        self.retry = retry


def _retry_for(error: Exception, argv: Sequence[str], args: argparse.Namespace) -> str | None:
    """The command to run next after ``error``, or None when no single command fixes it."""
    cmd = [PROG, *argv]
    if isinstance(error, _Refusal):
        return error.retry
    if isinstance(error, AmbiguousTextError):
        return f"{shlex.join(cmd)} --occurrence N   (N from the list above)"
    if isinstance(error, MatchCountError):
        return _with_flag(cmd, "--expect", str(error.total_occurrences)) + "   (if every match listed should count)"
    if isinstance(error, HashMismatchError) and getattr(args, "paragraph", None):
        return _with_flag(cmd, "--paragraph", f"P{error.paragraph_index}#{error.actual_hash}")
    if isinstance(error, TextNotFoundError):
        if error.total_occurrences:
            return f"{shlex.join(cmd)}   with --occurrence 0..{error.total_occurrences - 1}"
        return f"{PROG} list {shlex.quote(args.file)}"
    return None


def _with_flag(cmd: list[str], flag: str, value: str) -> str:
    """``cmd`` with ``flag``'s value replaced (or the flag appended)."""
    out = list(cmd)
    if flag in out and out.index(flag) + 1 < len(out):
        out[out.index(flag) + 1] = value
    else:
        out += [flag, value]
    return shlex.join(out)


def _paragraph_lines(doc: Document, refs: Sequence[str]) -> list[str]:
    """``ref| markup`` for each distinct ref: the paragraph as it now reads."""
    markup = doc.get_markup_text().splitlines()
    lines = []
    for ref in dict.fromkeys(refs):
        index = int(ref[1 : ref.index("#")])
        lines.append(f"{ref}| {markup[index - 1] if index <= len(markup) else ''}")
    return lines


def _edit_receipt(doc: Document, verb: str, results: Sequence[EditResult]) -> dict:
    refs = [ref for r in results for ref in (r.refs or (str(r),))]
    return {
        "verb": verb,
        "results": [
            {
                "ref": str(r),
                "refs": list(r.refs),
                "revision_ids": list(r.revision_ids),
                "comment_id": r.comment_id,
            }
            for r in results
        ],
        "paragraphs": _paragraph_lines(doc, refs),
    }


def _run_edit(doc: Document, args: argparse.Namespace) -> dict:
    scope = {"paragraph": args.paragraph, "occurrence": args.occurrence}
    if args.verb == "replace":
        if args.expect is not None:
            results = doc.replace_all(args.old, args.new, expected_count=args.expect, note=args.note)
        else:
            results = [doc.replace(args.old, args.new, note=args.note, **scope)]
    elif args.verb == "delete":
        results = [doc.delete(args.text, note=args.note, **scope)]
    elif args.verb == "insert-after":
        results = [doc.insert_after(args.anchor, args.text, note=args.note, **scope)]
    else:
        results = [doc.insert_before(args.anchor, args.text, note=args.note, **scope)]
    return _edit_receipt(doc, args.verb, results)


def _run_comment(doc: Document, args: argparse.Namespace) -> dict:
    comment_id = doc.add_comment(args.anchor, args.text, paragraph=args.paragraph, occurrence=args.occurrence)
    ref = next((c.paragraph_ref for c in doc.list_comments() if c.id == comment_id), None)
    return {
        "verb": "comment",
        "comment_id": comment_id,
        "ref": ref,
        "paragraphs": _paragraph_lines(doc, [ref] if ref else []),
    }


def _run_resolve(doc: Document, args: argparse.Namespace) -> dict:
    if args.all:
        resolved = (doc.accept_all if args.verb == "accept" else doc.reject_all)(author=args.by)
        return {
            "verb": args.verb,
            "resolved": int(resolved),
            "unhandled": resolved.unhandled,
            "unhandled_types": dict(resolved.unhandled_types),
        }
    one = doc.accept_revision if args.verb == "accept" else doc.reject_revision
    for rev_id in args.ids:
        if not one(rev_id, author=args.by):
            by = f" by {args.by}" if args.by else ""
            raise _Refusal(
                f"No pending revision {rev_id}{by} (it may have been resolved together with an earlier id). "
                f"Nothing was saved.",
                retry=f"{PROG} revisions {shlex.quote(args.file)}",
            )
    return {"verb": args.verb, "resolved": len(args.ids), "ids": list(args.ids)}


def _run_find(doc: Document, args: argparse.Namespace) -> dict:
    matches = doc.find_all(args.text, args.paragraph, expected_count=args.expect)
    texts = {p.index: p.text for p in doc.list_paragraphs_structured(limit=None)} if matches else {}
    return {
        "verb": "find",
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


def _run_read(doc: Document, args: argparse.Namespace) -> dict:
    if args.verb == "find":
        return _run_find(doc, args)
    if args.verb == "list":
        kwargs = {"start": args.start, "max_chars": args.max_chars}
        if args.limit is not None:
            kwargs["limit"] = args.limit or None
        return {"verb": "list", "lines": doc.list_paragraphs(**kwargs)}
    return {"verb": "revisions", "revisions": doc.list_revisions(author=args.by, paragraph=args.paragraph)}


def _print_text(out: dict) -> None:
    verb = out["verb"]
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
        line = f"{verb}ed {out['resolved']} revision(s)"
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
            details = []
            if r["revision_ids"]:
                details.append("revisions " + ", ".join(map(str, r["revision_ids"])))
            else:
                details.append("no new revision (amended your own insertion, or nothing changed)")
            if r["comment_id"] is not None:
                details.append(f"comment {r['comment_id']}")
            print(f"{verb} {' '.join(r['refs'] or [r['ref']])}  ({'; '.join(details)})")
    for line in out.get("paragraphs", []):
        print(f"  {line}")
    for w in out.get("warnings", []):
        print(f"warning: {w}")
    if out.get("saved"):
        print(f"saved {out['saved']}")


MUTATING = {"replace", "delete", "insert-after", "insert-before", "comment", "accept", "reject"}


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=PROG,
        description="One tracked-change edit to a .docx per command. `docx-edit skill` prints the agent guide.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    verbs = parser.add_subparsers(dest="verb", required=True, metavar="VERB")

    def verb(name: str, summary: str, example: str) -> argparse.ArgumentParser:
        sub = verbs.add_parser(
            name,
            help=summary,
            description=summary,
            epilog=f"example:\n  {example}",
            formatter_class=argparse.RawDescriptionHelpFormatter,
        )
        if name != "skill":
            sub.add_argument("file", help="the .docx file")
        return sub

    def scope(sub: argparse.ArgumentParser) -> None:
        sub.add_argument("--paragraph", metavar="REF", help="pin to this paragraph ref (e.g. P3#a7b2)")
        sub.add_argument("--occurrence", type=int, metavar="N", help="pick match N (0-based) when text repeats")

    def output(sub: argparse.ArgumentParser) -> None:
        sub.add_argument("--author", required=True, help="name recorded on the tracked change")
        where = sub.add_mutually_exclusive_group(required=True)
        where.add_argument("-o", "--output", metavar="OUT", help="save to OUT")
        where.add_argument("--in-place", action="store_true", help="overwrite FILE")

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
    output(sub)

    sub = verb("find", "List every match of text, with refs.", 'docx-edit find F.docx "30 days"')
    sub.add_argument("text")
    sub.add_argument("--paragraph", metavar="REF", help="search only this paragraph")
    sub.add_argument("--expect", type=int, metavar="N", help="refuse unless there are exactly N matches")

    sub = verb("list", "List paragraphs with their refs.", "docx-edit list F.docx --start 40 --limit 20")
    sub.add_argument("--start", type=int, default=1, help="first paragraph (1-based)")
    sub.add_argument("--limit", type=int, help="most paragraphs to list (default 200, 0 = all)")
    sub.add_argument("--max-chars", type=int, default=80, help="cut each paragraph at this length")

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
        output(sub)

    verb("skill", "Print the agent guide for this tool.", "docx-edit skill")

    for sub in verbs.choices.values():
        if sub.prog != f"{PROG} skill":
            sub.add_argument("--json", action="store_true", help="print one JSON object instead of text")
    return parser


def _check_usage(parser: argparse.ArgumentParser, args: argparse.Namespace) -> None:
    if args.verb == "replace" and args.expect is not None and (args.paragraph or args.occurrence is not None):
        parser.error("replace: --expect replaces every match in the document; drop --paragraph/--occurrence")
    if args.verb in ("accept", "reject") and bool(args.ids) == bool(args.all):
        parser.error(f"{args.verb}: give revision ids or --all, not both or neither")


def main(argv: Sequence[str] | None = None) -> int:
    """Run one docx-edit command; returns 0 (done) or 1 (refused, nothing saved).

    A bad command line exits with status 2 through argparse (SystemExit).
    """
    argv = list(sys.argv[1:] if argv is None else argv)
    parser = _parser()
    args = parser.parse_args(argv)
    if args.verb == "skill":
        print(AGENT_GUIDE, end="")
        return 0
    _check_usage(parser, args)

    mutating = args.verb in MUTATING
    try:
        with warnings.catch_warnings(record=True) as caught:
            # Only the library's own warnings are the caller's business; a
            # ResourceWarning or DeprecationWarning from a dependency is not.
            warnings.simplefilter("always", UnanchoredNoteWarning)
            warnings.simplefilter("always", UnhandledRevisionWarning)
            doc = Document.open(args.file, author=args.author if mutating else None)
            try:
                if args.verb in ("replace", "delete", "insert-after", "insert-before"):
                    out = _run_edit(doc, args)
                elif args.verb == "comment":
                    out = _run_comment(doc, args)
                elif args.verb in ("accept", "reject"):
                    out = _run_resolve(doc, args)
                else:
                    out = _run_read(doc, args)
                if mutating:
                    out["saved"] = str(doc.save(args.file if args.in_place else args.output))
            finally:
                doc.close(cleanup=True)
    except (DocxEditError, ValueError, OSError, _Refusal) as error:
        retry = _retry_for(error, argv, args)
        if args.json:
            payload = (
                {"type": "Refusal", "message": str(error)} if isinstance(error, _Refusal) else _error_payload(error)
            )
            print(json.dumps({"ok": False, "error": payload, "retry": retry}, default=_json_default))
        else:
            print(f"error: {error}", file=sys.stderr)
            if retry:
                print(f"retry: {retry}", file=sys.stderr)
            elif hint := getattr(error, "next", None):
                print(f"next (Python API): {hint}", file=sys.stderr)
        return 1

    library = (UnanchoredNoteWarning, UnhandledRevisionWarning)
    out["warnings"] = list(dict.fromkeys(str(w.message) for w in caught if issubclass(w.category, library)))
    if args.json:
        print(json.dumps({"ok": True, **out}, default=_json_default))
    else:
        _print_text(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
