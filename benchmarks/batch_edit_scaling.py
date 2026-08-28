#!/usr/bin/env python3
"""Benchmark: batch_edit / batch_rewrite scaling (ISSUES.md #51, #56).

Builds a ~600-paragraph document (multi-run paragraphs, ~400 chars each) and
measures batch_edit at increasing operation counts, apply and dry_run. Before
the per-batch <w:p> snapshot + seeded change-id fixes, per-op cost was
O(document size) and flat in the op count (~26 ms/op apply, ~3.5 ms/op
dry_run at this size); after, ~1.3 ms/op apply and ~0.3 ms/op dry_run,
falling further as the per-batch fixed cost amortizes over more ops. Since
#56 moved change-id seeding into ``DocxXMLEditor.__init__`` (parse time), the
first measured op no longer absorbs a one-time seed-scan on top of its own
cost — every op, including the first, times the same steady-state work.

Also measures ``batch_rewrite`` the same way (#56 threaded the same shared
<w:p> snapshot into its validation and apply loops, closing the O(items x
doc) gap batch_edit had before #51).

Usage:
    uv run python benchmarks/batch_edit_scaling.py [--ops 1000]

Timing output is informational. The script exits non-zero only on a
correctness assertion failure (every returned ref must be valid against a
fresh list_paragraphs() after the batch).
"""

import argparse
import shutil
import tempfile
import time
from pathlib import Path

from docx_editor import Document, EditOperation

TEST_DATA = Path(__file__).parent.parent / "tests" / "test_data" / "simple.docx"

N_PARAGRAPHS = 600

FILLER = (
    "The committee shall review the quarterly compliance report and "
    "record all findings in the shared register for later audit. "
)


def build_large_doc(n_paragraphs: int = N_PARAGRAPHS) -> tuple[Path, Path]:
    """Build a document with many multi-run paragraphs via direct XML injection.

    Each paragraph has a unique marker run plus three filler runs (~400 chars
    total, 'committee' appears three times). Returns (docx_path, tmp_dir);
    the caller removes tmp_dir.
    """
    tmp = tempfile.mkdtemp(prefix="bench_scale_")
    dest = Path(tmp) / "bench.docx"
    shutil.copy(TEST_DATA, dest)

    doc = Document.open(dest)
    doc.accept_all()
    doc.save()
    doc.close()

    doc = Document.open(dest, force_recreate=True)
    editor = doc._document_editor
    body = editor.dom.getElementsByTagName("w:body")[0]

    for p in list(editor.dom.getElementsByTagName("w:p")):
        if p.parentNode == body:
            body.removeChild(p)

    sect_pr = editor.dom.getElementsByTagName("w:sectPr")
    insert_before = sect_pr[0] if sect_pr else None

    filler_runs = f'<w:r><w:t xml:space="preserve">{FILLER}</w:t></w:r>' * 3
    for i in range(1, n_paragraphs + 1):
        p_xml = f'<w:p><w:r><w:t xml:space="preserve">[P{i:04d}] </w:t></w:r>{filler_runs}</w:p>'
        for node in editor._parse_fragment(p_xml):
            if insert_before:
                body.insertBefore(node, insert_before)
            else:
                body.appendChild(node)

    editor.save()
    save_path = doc.save()
    doc.close()
    return save_path, Path(tmp)


def run_batch(persist_path: Path, n_ops: int, *, dry_run: bool) -> float:
    """Time one batch_edit of n_ops on a fresh copy; return elapsed seconds."""
    tmp = tempfile.mkdtemp(prefix="bench_scale_run_")
    try:
        dest = Path(tmp) / "b.docx"
        shutil.copy(persist_path, dest)
        doc = Document.open(dest, force_recreate=True)
        try:
            return _timed_batch(doc, n_ops, dry_run=dry_run)
        finally:
            doc.close()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _timed_batch(doc: Document, n_ops: int, *, dry_run: bool) -> float:
    """Build and time the batch; assert the results are correct."""
    refs = doc.list_paragraphs(limit=None)
    n_paras = len(refs)
    # Coprime stride spreads targets across the document; beyond n_paras
    # ops it revisits paragraphs (valid: hashes are validated upfront
    # against the pre-batch state, and each paragraph holds three
    # 'committee' occurrences).
    assert n_ops <= 2 * n_paras, f"at most {2 * n_paras} ops supported"
    ops = [
        EditOperation.replace(
            "committee",
            f"BOARD_{i}",
            paragraph=refs[(i * 397) % n_paras].split("|")[0],
            occurrence=0,
        )
        for i in range(n_ops)
    ]

    start = time.perf_counter()
    if dry_run:
        results = doc.batch_edit(ops, dry_run=True)
        elapsed = time.perf_counter() - start
        invalid = [r for r in results if not r.valid]
        assert not invalid, f"dry_run flagged {len(invalid)} valid op(s) as invalid: {invalid[0].error}"
    else:
        results = doc.batch_edit(ops)
        elapsed = time.perf_counter() - start
        fresh = {entry.split("|")[0] for entry in doc.list_paragraphs(limit=None)}
        stale = [r for r in results if str(r) not in fresh]
        assert not stale, f"{len(stale)} returned ref(s) invalid after batch, e.g. {stale[0]}"
    return elapsed


def run_rewrite_batch(persist_path: Path, n_ops: int) -> float:
    """Time one batch_rewrite of n_ops on a fresh copy; return elapsed seconds."""
    tmp = tempfile.mkdtemp(prefix="bench_scale_rw_")
    try:
        dest = Path(tmp) / "b.docx"
        shutil.copy(persist_path, dest)
        doc = Document.open(dest, force_recreate=True)
        try:
            return _timed_rewrite_batch(doc, n_ops)
        finally:
            doc.close()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _timed_rewrite_batch(doc: Document, n_ops: int) -> float:
    """Build and time the rewrite batch; assert the results are correct.

    Unlike batch_edit, batch_rewrite rejects duplicate paragraph targets
    within one call, so (unlike ``_timed_batch``) n_ops is capped at the
    paragraph count rather than allowed to revisit paragraphs.
    """
    refs = doc.list_paragraphs(limit=None)
    n_paras = len(refs)
    assert n_ops <= n_paras, f"at most {n_paras} ops supported (one rewrite per paragraph)"
    # Distinct targets by construction. batch_edit's coprime-stride trick is
    # not reusable here: batch_rewrite rejects duplicate paragraphs, so a
    # stride sharing a factor with n_paras would abort the benchmark.
    rewrites = [(refs[i].split("|")[0], f"[REWRITTEN {i:04d}] {FILLER}") for i in range(n_ops)]

    start = time.perf_counter()
    group_ids = doc.batch_rewrite(rewrites)
    elapsed = time.perf_counter() - start
    assert len(group_ids) == n_ops, f"expected {n_ops} group ids, got {len(group_ids)}"
    return elapsed


def main() -> None:
    parser = argparse.ArgumentParser(description="batch_edit / batch_rewrite scaling benchmark")
    parser.add_argument("--ops", type=int, default=None, help="extra op count to measure (e.g. 1000)")
    args = parser.parse_args()
    if args.ops is not None and args.ops < 1:
        parser.error("--ops must be >= 1")

    op_counts = [10, 40, 120]
    if args.ops is not None and args.ops not in op_counts:
        op_counts.append(args.ops)

    print(f"Building {N_PARAGRAPHS}-paragraph document...")
    persist_path, tmp = build_large_doc()
    try:
        print()
        print(f"{'ops':>5} | {'apply s':>8} {'ms/op':>7} | {'dry_run s':>9} {'ms/op':>7}")
        print("-" * 46)
        for n in op_counts:
            apply_s = run_batch(persist_path, n, dry_run=False)
            dry_s = run_batch(persist_path, n, dry_run=True)
            print(f"{n:>5} | {apply_s:>8.3f} {apply_s / n * 1000:>7.2f} | {dry_s:>9.3f} {dry_s / n * 1000:>7.2f}")
        print()
        print("All correctness assertions passed (refs valid after each batch).")

        # batch_rewrite: same shared-<w:p>-snapshot fix (#56), but each rewrite
        # must target a distinct paragraph, so op counts are capped at N_PARAGRAPHS.
        with Document.open(persist_path, force_recreate=True) as probe:
            n_paras = len(probe.list_paragraphs(limit=None))
        rewrite_op_counts = [n for n in op_counts if n <= n_paras]
        skipped = [n for n in op_counts if n > n_paras]
        print()
        if skipped:
            print(f"Skipping batch_rewrite for op counts > {n_paras} (one rewrite per paragraph): {skipped}")
        print(f"{'ops':>5} | {'rewrite s':>9} {'ms/op':>7}")
        print("-" * 26)
        for n in rewrite_op_counts:
            rewrite_s = run_rewrite_batch(persist_path, n)
            print(f"{n:>5} | {rewrite_s:>9.3f} {rewrite_s / n * 1000:>7.2f}")
        print()
        print("All correctness assertions passed (group id returned for each rewrite).")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
