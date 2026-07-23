#!/usr/bin/env python3
"""Benchmark: accept-path scaling with revision count (ISSUES.md #57).

Builds a large multi-run document (via ``batch_edit_scaling.build_large_doc``),
lays down an N-operation ``batch_edit`` redline (one changeset, N groups, 2N
revisions), then times three ways to resolve it on identical revisions:

    * ``accept_all()``        — the multi-pass baseline (never regressed)
    * ``accept_changeset()``  — resolves the whole batch's changeset
    * ``accept_group()`` ×N   — group-by-group resolution loop

Before #57, ``accept_revision``/``reject_revision`` each did a full-document
``getElementsByTagName`` walk per member, repeated per pass, so the changeset
and group-loop paths were O(members x doc): on a 3000-paragraph doc a 240-rev
``accept_changeset`` measured 5.35 s (2.26x ``accept_all``) and the
``accept_group`` ×120 loop 6.13 s (2.59x) — see scratchpad
``dogfood4-perfscale/perf-report.md`` §4. After #57 the group/changeset path
builds one w:id->element index per call and threads it through every member, so
its cost tracks ``accept_all`` (ratio ~1x).

Usage:
    uv run python benchmarks/accept_scaling.py [--ops 500] [--paras 3000]

Timing output is informational. The script exits non-zero only on a
correctness assertion failure (every revision must be resolved, leaving
``list_revisions() == []``).
"""

import argparse
import shutil
import tempfile
import time
from pathlib import Path

from batch_edit_scaling import N_PARAGRAPHS, build_large_doc

from docx_editor import Document, EditOperation


def _build_redline(doc: Document, n_ops: int) -> int:
    """Apply an n_ops ``batch_edit`` redline; return the revision count (2*n_ops)."""
    refs = doc.list_paragraphs(limit=None)
    n_paras = len(refs)
    assert n_ops <= 2 * n_paras, f"at most {2 * n_paras} ops supported"
    # Coprime stride spreads targets; each paragraph holds three 'committee'
    # occurrences, so revisiting a paragraph past n_paras ops stays valid.
    ops = [
        EditOperation.replace(
            "committee",
            f"BOARD_{i}",
            paragraph=refs[(i * 397) % n_paras].split("|")[0],
            occurrence=0,
        )
        for i in range(n_ops)
    ]
    doc.batch_edit(ops)
    return len(doc.list_revisions())


def _unique_group_ids(doc: Document) -> list[int]:
    """Group ids of the current revisions, in first-seen (document) order."""
    seen: list[int] = []
    for rev in doc.list_revisions():
        if rev.group_id is not None and rev.group_id not in seen:
            seen.append(rev.group_id)
    return seen


def run_resolve(persist_path: Path, n_ops: int, mode: str) -> tuple[float, int]:
    """Time one resolution of an n_ops redline on a fresh copy.

    Returns (elapsed seconds, revisions resolved). The redline build is not
    timed — only the resolve call(s).
    """
    tmp = tempfile.mkdtemp(prefix="bench_accept_")
    try:
        dest = Path(tmp) / "a.docx"
        shutil.copy(persist_path, dest)
        doc = Document.open(dest, force_recreate=True)
        try:
            n_revs = _build_redline(doc, n_ops)

            if mode == "accept_all":
                start = time.perf_counter()
                resolved = doc.accept_all()
                elapsed = time.perf_counter() - start
            elif mode == "accept_changeset":
                changeset_id = doc.list_revisions()[0].changeset_id
                assert changeset_id is not None
                start = time.perf_counter()
                resolved = doc.accept_changeset(changeset_id)
                elapsed = time.perf_counter() - start
            elif mode == "accept_group_loop":
                group_ids = _unique_group_ids(doc)
                start = time.perf_counter()
                resolved = sum(doc.accept_group(gid) for gid in group_ids)
                elapsed = time.perf_counter() - start
            else:
                raise ValueError(f"unknown mode: {mode}")

            assert resolved == n_revs, f"{mode}: resolved {resolved} of {n_revs} revisions"
            assert doc.list_revisions() == [], f"{mode}: {len(doc.list_revisions())} revisions left unresolved"
            return elapsed, n_revs
        finally:
            doc.close()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="accept-path scaling benchmark (ISSUES.md #57)")
    parser.add_argument("--ops", type=int, default=None, help="extra op count to measure (e.g. 500)")
    parser.add_argument("--paras", type=int, default=N_PARAGRAPHS, help=f"document size (default {N_PARAGRAPHS})")
    args = parser.parse_args()
    if args.ops is not None and args.ops < 1:
        parser.error("--ops must be >= 1")
    if args.paras < 1:
        parser.error("--paras must be >= 1")

    op_counts = [10, 40, 120]
    if args.ops is not None and args.ops not in op_counts:
        op_counts.append(args.ops)

    print(f"Building {args.paras}-paragraph document...")
    persist_path, tmp = build_large_doc(args.paras)
    try:
        print()
        header = (
            f"{'ops':>5} {'revs':>5} | {'accept_all':>10} | "
            f"{'changeset':>10} {'ratio':>6} | {'group×N':>10} {'ratio':>6}"
        )
        print(header)
        print("-" * len(header))
        for n in op_counts:
            all_s, revs = run_resolve(persist_path, n, "accept_all")
            cs_s, _ = run_resolve(persist_path, n, "accept_changeset")
            grp_s, _ = run_resolve(persist_path, n, "accept_group_loop")
            print(
                f"{n:>5} {revs:>5} | {all_s:>9.3f}s | {cs_s:>9.3f}s {cs_s / all_s:>5.2f}x "
                f"| {grp_s:>9.3f}s {grp_s / all_s:>5.2f}x"
            )
        print()
        print("Ratios near 1.00x mean the group/changeset path no longer pays O(members x doc)")
        print("(pre-#57 baseline on doc3000/240-rev: changeset 2.26x, group×N 2.59x).")
        print("All correctness assertions passed (every revision resolved).")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
