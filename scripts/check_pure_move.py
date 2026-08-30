"""Verify that a ``docx_editor.track_changes`` refactor is a pure move.

Compares the module/package as it was at the merge base with the working tree
(ROADMAP.md #73: every decomposition step is a literal cut/paste). Usage::

    uv run python scripts/check_pure_move.py [--base main]

Three checks, all of which must pass:

1. **Line multiset.** Every non-plumbing line of the old files appears in the
   new files exactly as often, and vice versa. Plumbing is a deliberately
   narrow allowlist: blank lines, module docstrings, column-0 ``import``/
   ``from`` lines and the ``    Name,`` entries of their parenthesised blocks,
   ``__all__`` blocks, bare ``(``/``)``/``]`` lines, column-0 ``class``
   headers (a mixin changes a base list) and column-0 ``#`` comments.
   Indented comments and every function/method line must match, count for
   count. ``base.py`` is exempt here and covered by check 3.
2. **AST identity.** Every method of ``RevisionManager`` (old) maps to the same
   byte-identical source, decorators included, on ``RevisionManager`` or a
   ``*Mixin`` class (new); no method may appear twice. Module-level
   functions, classes and assignments are compared the same way. This catches
   a reordering inside a method that the multiset cannot see.
3. **``base.py`` is copies only.** Each non-blank line of ``base.py`` must be a
   verbatim line found elsewhere in the package (a copied decorator, ``def``
   or parameter line), ``raise NotImplementedError``, or an attribute
   annotation.

Exit status is 1 on any failure; every offending line is printed with its
file and side.
"""

from __future__ import annotations

import argparse
import ast
import re
import subprocess
import sys
from collections import Counter
from pathlib import Path

OLD_MODULE = "docx_editor/track_changes.py"
PACKAGE = "docx_editor/track_changes"
BASE_FILE = f"{PACKAGE}/base.py"

_IMPORT_LINE = re.compile(r"^(import |from \S+ import )")
_IMPORT_OPEN = re.compile(r"^from \S+ import \($")
_IMPORT_ENTRY = re.compile(r"^    [A-Za-z_][A-Za-z0-9_]*,$")
_ALL_OPEN = re.compile(r"^__all__ = \[$")
_ALL_ENTRY = re.compile(r'^    "[A-Za-z_][A-Za-z0-9_]*",$')
_CLASS_HEADER = re.compile(r"^class [A-Za-z_][A-Za-z0-9_]*(\(.*\))?:$")
_ATTR_ANNOTATION = re.compile(r"^    [A-Za-z_][A-Za-z0-9_]*: .+$")
_NOT_IMPLEMENTED = re.compile(r"^\s+raise NotImplementedError$")


def _git(*args: str) -> str:
    return subprocess.run(["git", *args], capture_output=True, text=True, check=True).stdout


def old_files(ref: str) -> dict[str, str]:
    names = _git("ls-tree", "-r", "--name-only", ref, "--", OLD_MODULE, PACKAGE).split()
    return {name: _git("show", f"{ref}:{name}") for name in sorted(names) if name.endswith(".py")}


def new_files() -> dict[str, str]:
    paths = [Path(OLD_MODULE)] if Path(OLD_MODULE).exists() else []
    paths += sorted(Path(PACKAGE).glob("*.py")) if Path(PACKAGE).is_dir() else []
    return {str(p): p.read_text() for p in paths}


def _docstring_lines(src: str) -> set[int]:
    """1-based line numbers spanned by the module docstring, if any."""
    tree = ast.parse(src)
    if tree.body and isinstance(tree.body[0], ast.Expr) and isinstance(tree.body[0].value, ast.Constant):
        node = tree.body[0]
        return set(range(node.lineno, (node.end_lineno or node.lineno) + 1))
    return set()


def content_lines(src: str) -> list[str]:
    """The lines of ``src`` that a pure move must preserve verbatim."""
    skip = _docstring_lines(src)
    kept: list[str] = []
    block: re.Pattern[str] | None = None  # entry pattern while inside an import/__all__ block
    for lineno, line in enumerate(src.splitlines(), start=1):
        if lineno in skip or not line.strip():
            continue
        if block is not None:
            if line in (")", "]"):
                block = None
                continue
            if block.match(line):
                continue
            block = None  # anything else ends the block and is judged on its own
        if _IMPORT_OPEN.match(line):
            block = _IMPORT_ENTRY
            continue
        if _ALL_OPEN.match(line):
            block = _ALL_ENTRY
            continue
        if _IMPORT_LINE.match(line) or line in ("(", ")", "]"):
            continue
        if _CLASS_HEADER.match(line) or line.startswith("#"):
            continue
        kept.append(line)
    return kept


def check_line_multiset(old: dict[str, str], new: dict[str, str]) -> list[str]:
    def tally(files: dict[str, str]) -> tuple[Counter[str], dict[str, set[str]]]:
        counts: Counter[str] = Counter()
        where: dict[str, set[str]] = {}
        for name, src in files.items():
            if name == BASE_FILE:
                continue
            for line in content_lines(src):
                counts[line] += 1
                where.setdefault(line, set()).add(name)
        return counts, where

    old_counts, old_where = tally(old)
    new_counts, new_where = tally(new)
    problems: list[str] = []
    for line, n in sorted((old_counts - new_counts).items()):
        problems.append(f"lost ×{n} from {', '.join(sorted(old_where[line]))}: {line!r}")
    for line, n in sorted((new_counts - old_counts).items()):
        problems.append(f"added ×{n} in {', '.join(sorted(new_where[line]))}: {line!r}")
    return problems


def _segment(lines: list[str], node: ast.AST) -> str:
    start = getattr(node, "lineno", 0)
    for deco in getattr(node, "decorator_list", []):
        start = min(start, deco.lineno)
    return "\n".join(lines[start - 1 : getattr(node, "end_lineno", start)])


def _names_of(node: ast.stmt) -> str | None:
    if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
        return node.name
    if isinstance(node, ast.Assign):
        return ",".join(t.id for t in node.targets if isinstance(t, ast.Name)) or None
    if isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
        return node.target.id
    return None


def _is_manager_part(cls: ast.ClassDef) -> bool:
    return cls.name == "RevisionManager" or cls.name.endswith("Mixin") or cls.name == "_RevisionManagerBase"


def collect(files: dict[str, str]) -> tuple[dict[str, tuple[str, str]], dict[str, tuple[str, str]], list[str]]:
    """Return (manager members, module-level definitions, problems).

    Both dicts map a name to ``(file, source)``. Manager members are the
    statements in the body of ``RevisionManager`` and every ``*Mixin`` class;
    ``_RevisionManagerBase`` is excluded (its stubs are checked by check 3).
    """
    members: dict[str, tuple[str, str]] = {}
    toplevel: dict[str, tuple[str, str]] = {}
    problems: list[str] = []
    for name, src in files.items():
        lines = src.splitlines()
        for node in ast.parse(src).body:
            key = _names_of(node)
            if key is None or key == "__all__":
                continue
            if isinstance(node, ast.ClassDef) and _is_manager_part(node):
                if node.name == "_RevisionManagerBase":
                    continue
                for i, stmt in enumerate(node.body):
                    mkey = _names_of(stmt) or f"<{node.name} body statement {i}>"
                    if mkey in members:
                        problems.append(f"duplicate manager member {mkey!r} in {members[mkey][0]} and {name}")
                    members[mkey] = (name, _segment(lines, stmt))
                continue
            if key in toplevel:
                problems.append(f"duplicate top-level definition {key!r} in {toplevel[key][0]} and {name}")
            toplevel[key] = (name, _segment(lines, node))
    return members, toplevel, problems


def _diff_dicts(label: str, old: dict[str, tuple[str, str]], new: dict[str, tuple[str, str]]) -> list[str]:
    problems: list[str] = []
    for key in sorted(old.keys() - new.keys()):
        problems.append(f"{label} {key!r} lost from {old[key][0]}")
    for key in sorted(new.keys() - old.keys()):
        problems.append(f"{label} {key!r} added in {new[key][0]}")
    for key in sorted(old.keys() & new.keys()):
        if old[key][1] != new[key][1]:
            problems.append(f"{label} {key!r} differs: {old[key][0]} -> {new[key][0]}")
    return problems


def check_ast_identity(old: dict[str, str], new: dict[str, str]) -> list[str]:
    old_members, old_top, problems = collect(old)
    new_members, new_top, new_problems = collect(new)
    problems += new_problems
    problems += _diff_dicts("manager member", old_members, new_members)
    problems += _diff_dicts("top-level", old_top, new_top)
    return problems


def check_base_is_copies(new: dict[str, str]) -> list[str]:
    if BASE_FILE not in new:
        return []
    pool: set[str] = set()
    for name, src in new.items():
        if name != BASE_FILE:
            pool.update(src.splitlines())
    problems: list[str] = []
    skip = _docstring_lines(new[BASE_FILE])
    for lineno, line in enumerate(new[BASE_FILE].splitlines(), start=1):
        if lineno in skip or not line.strip() or line in pool:
            continue
        if _NOT_IMPLEMENTED.match(line) or _ATTR_ANNOTATION.match(line):
            continue
        if _IMPORT_LINE.match(line) or _IMPORT_ENTRY.match(line) or line in ("(", ")"):
            continue
        if _CLASS_HEADER.match(line) or line.startswith("#"):
            continue
        problems.append(f"{BASE_FILE}:{lineno} is not a verbatim copy: {line!r}")
    return problems


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Verify that a docx_editor.track_changes refactor is a pure move.")
    parser.add_argument("--base", default="main", help="ref whose merge base with HEAD is the 'before' state")
    args = parser.parse_args(argv)
    ref = _git("merge-base", args.base, "HEAD").strip()
    old, new = old_files(ref), new_files()
    print(f"before ({ref[:12]}): {', '.join(old)}")
    print(f"after  (working tree): {', '.join(new)}")

    failed = False
    for title, problems in (
        ("check 1: line multiset", check_line_multiset(old, new)),
        ("check 2: AST identity", check_ast_identity(old, new)),
        ("check 3: base.py is copies only", check_base_is_copies(new)),
    ):
        status = "FAIL" if problems else "ok"
        print(f"{title}: {status}")
        for problem in problems:
            print(f"  {problem}")
        failed = failed or bool(problems)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
