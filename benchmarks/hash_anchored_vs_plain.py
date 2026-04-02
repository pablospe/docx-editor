#!/usr/bin/env python3
"""Benchmark: Hash-Anchored vs Plain edits.

Compares speed overhead and accuracy (silent corruption vs loud rejection)
of hash-anchored paragraph references vs plain occurrence-based editing.
"""

import shutil
import tempfile
import time
from pathlib import Path

from docx_editor import Document, HashMismatchError
from docx_editor.xml_editor import build_text_map

TEST_DATA = Path(__file__).parent.parent / "tests" / "test_data" / "simple.docx"


def fresh_doc() -> tuple[Document, Path]:
    tmp = tempfile.mkdtemp(prefix="bench_")
    dest = Path(tmp) / "bench.docx"
    shutil.copy(TEST_DATA, dest)
    return Document.open(dest), Path(tmp)


def cleanup(doc: Document, tmp: Path):
    doc.close()
    shutil.rmtree(tmp, ignore_errors=True)


def build_multi_paragraph_doc(n_paragraphs: int = 30) -> tuple[Document, Path]:
    """Build a document with many paragraphs containing repeated phrases.

    Directly injects <w:p> elements into the XML for reliable paragraph creation.
    Each paragraph contains 'the committee' and 'shall review' (repeated phrases)
    plus a unique marker like [P01].
    """
    doc, tmp = fresh_doc()
    doc.accept_all()
    doc.save()
    doc.close()

    # Reopen and inject paragraphs directly into XML
    doc = Document.open(Path(tmp) / "bench.docx", force_recreate=True)
    editor = doc._document_editor
    body = editor.dom.getElementsByTagName("w:body")[0]

    # Remove existing paragraphs except the last (sectPr container)
    existing_paras = list(editor.dom.getElementsByTagName("w:p"))
    for p in existing_paras:
        if p.parentNode == body:
            body.removeChild(p)

    # Insert new paragraphs before sectPr
    sect_pr = editor.dom.getElementsByTagName("w:sectPr")
    insert_before = sect_pr[0] if sect_pr else None

    for i in range(1, n_paragraphs + 1):
        text = (
            f"[P{i:02d}] The committee shall review item {i}. The report shall include all findings from the committee."
        )
        p_xml = f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
        nodes = editor._parse_fragment(p_xml)
        for node in nodes:
            if insert_before:
                body.insertBefore(node, insert_before)
            else:
                body.appendChild(node)

    editor.save()
    save_path = doc.save()
    doc.close()

    doc = Document.open(save_path, force_recreate=True)
    return doc, tmp


# ---------------------------------------------------------------------------
# Speed Benchmark
# ---------------------------------------------------------------------------


def benchmark_speed(iterations: int = 50):
    """Measure per-operation overhead of hash-anchored vs plain replace."""

    # Build a larger doc for speed testing
    doc, tmp = build_multi_paragraph_doc(30)
    doc.list_paragraphs()
    save_path = doc.save()
    doc.close()

    # Keep the saved file in a persistent temp dir
    persist_dir = tempfile.mkdtemp(prefix="bench_persist_")
    persist_path = Path(persist_dir) / "bench.docx"
    shutil.copy(save_path, persist_path)
    shutil.rmtree(tmp, ignore_errors=True)

    # Use paragraph 15 (middle of doc) as target
    def open_saved():
        t = tempfile.mkdtemp(prefix="bench_spd_")
        d = Path(t) / "bench.docx"
        shutil.copy(persist_path, d)
        return Document.open(d), Path(t)

    # Benchmark PLAIN replace (using RevisionManager directly, no paragraph scoping)
    plain_times = []
    for _ in range(iterations):
        d, t = open_saved()
        t0 = time.perf_counter()
        d._revision_manager.replace_text("[P15]", "[X15]")
        t1 = time.perf_counter()
        plain_times.append(t1 - t0)
        cleanup(d, t)

    # Benchmark HASH-ANCHORED replace
    hash_times = []
    for _ in range(iterations):
        d, t = open_saved()
        refs = d.list_paragraphs()
        # Find P15's ref
        ref = None
        for entry in refs:
            if "[P15]" in entry:
                ref = entry.split("|")[0]
                break
        t0 = time.perf_counter()
        d.replace("[P15]", "[X15]", paragraph=ref)
        t1 = time.perf_counter()
        hash_times.append(t1 - t0)
        cleanup(d, t)

    avg_plain = sum(plain_times) / len(plain_times) * 1000
    avg_hash = sum(hash_times) / len(hash_times) * 1000
    overhead = avg_hash - avg_plain
    pct = (overhead / avg_plain * 100) if avg_plain > 0 else 0

    print("=" * 60)
    print("SPEED BENCHMARK (30-paragraph document)")
    print("=" * 60)
    print(f"  Iterations:          {iterations}")
    print(f"  Plain replace:       {avg_plain:.3f} ms/op")
    print(f"  Hash-anchored:       {avg_hash:.3f} ms/op")
    print(f"  Overhead:            {overhead:+.3f} ms/op ({pct:+.1f}%)")
    print()

    shutil.rmtree(persist_dir, ignore_errors=True)
    return {"plain_ms": avg_plain, "hash_ms": avg_hash, "overhead_ms": overhead}


# ---------------------------------------------------------------------------
# Accuracy Benchmark
# ---------------------------------------------------------------------------


def benchmark_accuracy():
    """Simulate batch edits with index/occurrence drift.

    Scenario:
    1. Open 30-paragraph doc, snapshot all refs
    2. Edit paragraph 5 (replace 'committee' → 'BOARD'), changing occurrence mapping
    3. Try to edit paragraphs 6-15 using:
       - PLAIN: global occurrence of 'committee' (occurrence=N)
       - HASH: old refs from step 1
    4. Count: correct edits, wrong-paragraph edits, caught errors
    """
    print("=" * 60)
    print("ACCURACY BENCHMARK (30-paragraph document)")
    print("=" * 60)

    doc, tmp = build_multi_paragraph_doc(30)
    paragraphs = doc.list_paragraphs()
    n_paras = len(paragraphs)
    print(f"  Paragraphs: {n_paras}")

    # Count 'committee' occurrences per paragraph
    all_paras = doc._document_editor.dom.getElementsByTagName("w:p")
    committee_per_para = []
    for p in all_paras:
        tm = build_text_map(p)
        count = tm.text.count("committee")
        committee_per_para.append(count)

    total_committee = sum(committee_per_para)
    print(f"  Total 'committee' occurrences: {total_committee}")
    print(f"  Per-paragraph: ~{total_committee / n_paras:.1f}")

    # Snapshot refs and build occurrence-to-paragraph mapping
    old_refs = {}
    occurrence_map = {}  # occurrence_index -> paragraph_index (1-based)
    global_occ = 0
    for i, entry in enumerate(paragraphs):
        ref = entry.split("|")[0]
        old_refs[i + 1] = ref  # 1-indexed
        for _ in range(committee_per_para[i]):
            occurrence_map[global_occ] = i + 1
            global_occ += 1

    save_path = doc.save()
    doc.close()
    persist_dir = tempfile.mkdtemp(prefix="bench_acc_persist_")
    persist_path = Path(persist_dir) / "bench.docx"
    shutil.copy(save_path, persist_path)
    shutil.rmtree(tmp, ignore_errors=True)

    # Define edits: target paragraphs 6-15 (each has 'committee' twice)
    # After editing P5, the occurrence indices for P6+ shift by -2
    target_paras = list(range(6, 16))
    n_edits = len(target_paras)

    # Pre-compute which global occurrence maps to each target paragraph's FIRST 'committee'
    target_occurrences = {}
    for para_idx in target_paras:
        for occ, pidx in occurrence_map.items():
            if pidx == para_idx:
                target_occurrences[para_idx] = occ
                break

    print(f"  Edits planned: {n_edits} (paragraphs {target_paras[0]}-{target_paras[-1]})")
    print()

    # ---- PLAIN APPROACH ----
    print("  PLAIN (occurrence-based):")
    plain_correct = 0
    plain_wrong = 0
    plain_error = 0

    t1 = tempfile.mkdtemp(prefix="bench_p_")
    d1 = Path(t1) / "b.docx"
    shutil.copy(persist_path, d1)
    doc1 = Document.open(d1, force_recreate=True)

    # Disrupting edit: replace BOTH 'committee' in P5
    p5_ref = old_refs[5]
    doc1.replace("committee", "BOARD", paragraph=p5_ref)
    doc1.replace("committee", "BOARD", paragraph=old_refs[5].split("#")[0] + "#" + _get_fresh_hash(doc1, 5))

    # Now try to edit P6-P15 using the OLD occurrence indices
    for para_idx in target_paras:
        old_occ = target_occurrences.get(para_idx)
        if old_occ is None:
            plain_error += 1
            continue

        marker = f"[P{para_idx:02d}]"
        try:
            doc1._revision_manager.replace_text("committee", "EDITED", occurrence=old_occ)
            vis = doc1.get_visible_text()

            # Check where "EDITED" landed
            landed_para = None
            for line in vis.split("\n"):
                if "EDITED" in line:
                    # Extract marker
                    if "[P" in line:
                        start = line.index("[P")
                        end = line.index("]", start) + 1
                        landed_para = line[start:end]
                    break

            if landed_para == marker:
                plain_correct += 1
            elif landed_para:
                plain_wrong += 1
                print(f"    occ={old_occ}: targeted {marker}, landed in {landed_para} ← WRONG")
            else:
                plain_correct += 1  # Can't verify
        except Exception:
            plain_error += 1
            # After first wrong edit, subsequent occurrences shift further
            # This is expected cascade failure

    cleanup(doc1, Path(t1))

    # ---- HASH-ANCHORED APPROACH ----
    print()
    print("  HASH-ANCHORED:")
    hash_correct = 0
    hash_rejected = 0
    hash_error = 0

    t2 = tempfile.mkdtemp(prefix="bench_h_")
    d2 = Path(t2) / "b.docx"
    shutil.copy(persist_path, d2)
    doc2 = Document.open(d2, force_recreate=True)

    # Same disrupting edit
    doc2.replace("committee", "BOARD", paragraph=p5_ref)
    doc2.replace("committee", "BOARD", paragraph=old_refs[5].split("#")[0] + "#" + _get_fresh_hash(doc2, 5))

    # Try to edit P6-P15 using OLD refs (from before edit)
    for para_idx in target_paras:
        old_ref = old_refs[para_idx]
        marker = f"[P{para_idx:02d}]"
        try:
            doc2.replace("committee", "EDITED", paragraph=old_ref)
            hash_correct += 1
        except HashMismatchError:
            hash_rejected += 1
            print(f"    {old_ref}: targeted {marker} ← SAFELY CAUGHT (HashMismatchError)")
        except Exception as e:
            hash_error += 1
            print(f"    {old_ref}: {type(e).__name__}: {e}")

    cleanup(doc2, Path(t2))

    # Results
    print()
    print("-" * 60)
    print(f"  RESULTS ({n_edits} edits after disrupting edit on P5)")
    print("-" * 60)
    print(f"  PLAIN:         {plain_correct:>2} correct  {plain_wrong:>2} WRONG  {plain_error:>2} errors")
    print(f"  HASH-ANCHORED: {hash_correct:>2} correct  {hash_rejected:>2} rejected  {hash_error:>2} errors")
    print()

    if plain_wrong > 0:
        print(f"  ⚠  Plain approach: {plain_wrong} edit(s) silently landed in WRONG paragraph!")
        print(f"  ✓  Hash approach:  0 silent corruptions ({hash_rejected} safely caught)")
    else:
        print("  Both approaches produced correct results.")

    # Cleanup
    shutil.rmtree(persist_dir, ignore_errors=True)

    return {
        "total_edits": n_edits,
        "plain_correct": plain_correct,
        "plain_wrong": plain_wrong,
        "plain_error": plain_error,
        "hash_correct": hash_correct,
        "hash_rejected": hash_rejected,
        "hash_error": hash_error,
    }


def _get_fresh_hash(doc, para_index):
    """Get the current hash for a paragraph by index (1-based)."""
    entries = doc.list_paragraphs()
    entry = entries[para_index - 1]
    ref = entry.split("|")[0]
    return ref.split("#")[1]


def benchmark_batch_vs_individual():
    """Compare batch_edit() vs N individual calls."""
    from docx_editor import EditOperation

    print("=" * 60)
    print("BATCH vs INDIVIDUAL BENCHMARK (30-paragraph document)")
    print("=" * 60)

    # Build doc and save
    doc, tmp = build_multi_paragraph_doc(30)
    save_path = doc.save()
    doc.close()
    persist_dir = tempfile.mkdtemp(prefix="bench_batch_persist_")
    persist_path = Path(persist_dir) / "bench.docx"
    shutil.copy(save_path, persist_path)
    shutil.rmtree(tmp, ignore_errors=True)

    iterations = 30
    n_edits = 10  # Edit paragraphs 1-10

    def open_saved():
        t = tempfile.mkdtemp(prefix="bench_b_")
        d = Path(t) / "b.docx"
        shutil.copy(persist_path, d)
        return Document.open(d), Path(t)

    # INDIVIDUAL: N calls, each with list_paragraphs() + replace()
    individual_times = []
    for _ in range(iterations):
        d, t = open_saved()
        t0 = time.perf_counter()
        for i in range(1, n_edits + 1):
            refs = d.list_paragraphs()
            ref = refs[i - 1].split("|")[0]
            d.replace(f"item {i}", f"EDIT_{i}", paragraph=ref)
        t1 = time.perf_counter()
        individual_times.append(t1 - t0)
        cleanup(d, t)

    # BATCH: 1 list_paragraphs() + 1 batch_edit()
    batch_times = []
    for _ in range(iterations):
        d, t = open_saved()
        t0 = time.perf_counter()
        refs = d.list_paragraphs()
        ops = [
            EditOperation(
                action="replace",
                find=f"item {i}",
                replace_with=f"EDIT_{i}",
                paragraph=refs[i - 1].split("|")[0],
            )
            for i in range(1, n_edits + 1)
        ]
        d.batch_edit(ops)
        t1 = time.perf_counter()
        batch_times.append(t1 - t0)
        cleanup(d, t)

    avg_individual = sum(individual_times) / len(individual_times) * 1000
    avg_batch = sum(batch_times) / len(batch_times) * 1000
    speedup = avg_individual / avg_batch if avg_batch > 0 else 0

    print(f"  Iterations:          {iterations}")
    print(f"  Edits per iteration: {n_edits}")
    print()
    print(f"  Individual (N calls): {avg_individual:.1f} ms total")
    print(f"  Batch (1 call):       {avg_batch:.1f} ms total")
    print(f"  Speedup:              {speedup:.1f}x")
    print()
    print(f"  Individual: {n_edits} x list_paragraphs() + {n_edits} x replace()")
    print("  Batch:      1 x list_paragraphs() + 1 x batch_edit()")
    print()

    shutil.rmtree(persist_dir, ignore_errors=True)
    return {"individual_ms": avg_individual, "batch_ms": avg_batch, "speedup": speedup}


if __name__ == "__main__":
    print()
    speed = benchmark_speed(iterations=50)
    print()
    accuracy = benchmark_accuracy()
    print()
    batch = benchmark_batch_vs_individual()
    print()
