#!/usr/bin/env python3
"""Assemble the docx corpus described by manifest.json into files/.

manifest.json is the single source of truth. Each entry has a "kind":
  local      copied from a repo-relative path (tests/test_data fixtures)
  download   fetched from a commit-pinned raw.githubusercontent.com URL and
             verified against the recorded sha256 (mismatch = hard failure)
  generated  produced locally with LibreOffice (soffice) or pandoc from the
             sources in srcgen/; skipped with a notice if the tool is absent

Exits 1 if any local/download entry is missing at the end; tool-absence
skips of generated entries are not failures.
"""

import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
import urllib.request
from pathlib import Path

CORPUS = Path(__file__).resolve().parent
REPO_ROOT = CORPUS.parents[1]
FILES = CORPUS / "files"
SRCGEN = CORPUS / "srcgen"
LO_PROFILE = CORPUS / "loprofile"
MANIFEST_PATH = CORPUS / "manifest.json"

MAX_SIZE = 2 * 1024 * 1024  # 2MB per corpus file
DOWNLOAD_TIMEOUT = 30  # seconds
CONVERT_TIMEOUT = 120  # seconds, per soffice/pandoc invocation

# Recipes for kind=generated entries: output name -> (tool, input path).
# Inputs under srcgen/ are corpus-relative; tests/... are repo-relative.
RECIPES: dict[str, tuple[str, str]] = {
    "lo_from_txt.docx": ("soffice", "srcgen/plain.txt"),
    "lo_from_html.docx": ("soffice", "srcgen/rich.html"),
    "lo_from_html_unicode.docx": ("soffice", "srcgen/unicode.html"),
    "lo_from_rtf.docx": ("soffice", "srcgen/doc.rtf"),
    "lo_from_odt.docx": ("soffice", "srcgen/plain.odt"),
    "lo_resave_tricky_track_changes.docx": ("soffice", "tests/test_data/tricky-track-changes.docx"),
    "lo_resave_with_tables.docx": ("soffice", "tests/test_data/with_tables.docx"),
    "lo_resave_oxml_trackchanges.docx": ("soffice", "tests/test_data/OXML_TrackChanges_Test.docx"),
    "pandoc_notes.docx": ("pandoc", "srcgen/notes.md"),
    "pandoc_unicode.docx": ("pandoc", "srcgen/unicode.html"),
}


def sha16(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()[:16]


def resolve_input(source: str) -> Path:
    if source.startswith("srcgen/"):
        return CORPUS / source
    return REPO_ROOT / source


def download(url: str, dest: Path, expected_sha: str) -> str | None:
    """Fetch url to dest and verify its truncated sha256. Returns an error string or None."""
    try:
        with urllib.request.urlopen(url, timeout=DOWNLOAD_TIMEOUT) as r:
            data = r.read(MAX_SIZE + 1)
    except Exception as e:
        return f"download failed: {e}"
    if len(data) > MAX_SIZE:
        return "too large (>2MB)"
    actual = hashlib.sha256(data).hexdigest()[:16]
    if actual != expected_sha:
        return f"sha256 mismatch: expected {expected_sha}, got {actual}"
    dest.write_bytes(data)
    return None


def soffice_convert(soffice: str, src: Path, to_format: str, dest: Path) -> str | None:
    """Convert src with LibreOffice into dest. Returns an error string or None.

    soffice names its output by the input stem, so convert into a temp dir
    and move the result (also avoids stem collisions, e.g. plain.txt/plain.odt).
    docx output names its export filter explicitly — HTML inputs open as
    Writer/Web documents, which have no default docx filter.
    """
    convert_to = f"{to_format}:MS Word 2007 XML" if to_format == "docx" else to_format
    with tempfile.TemporaryDirectory(dir=CORPUS) as tmp:
        cmd = [
            soffice,
            "--headless",
            f"-env:UserInstallation=file://{LO_PROFILE}",
            "--convert-to",
            convert_to,
            "--outdir",
            tmp,
            str(src),
        ]
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=CONVERT_TIMEOUT)
        except subprocess.TimeoutExpired:
            return "soffice timeout"
        produced = Path(tmp) / f"{src.stem}.{to_format}"
        if proc.returncode != 0 or not produced.exists():
            return f"soffice failed: {(proc.stderr or proc.stdout or 'no output produced')[:200]}"
        shutil.move(str(produced), dest)
    return None


def pandoc_convert(pandoc: str, src: Path, dest: Path) -> str | None:
    try:
        proc = subprocess.run(
            [pandoc, str(src), "-o", str(dest)],
            capture_output=True,
            text=True,
            timeout=CONVERT_TIMEOUT,
        )
    except subprocess.TimeoutExpired:
        return "pandoc timeout"
    if proc.returncode != 0 or not dest.exists():
        return f"pandoc failed: {(proc.stderr or 'no output produced')[:200]}"
    return None


def generate(name: str, tools: dict[str, str | None]) -> str | None:
    """Generate one kind=generated corpus file. Returns an error string or None."""
    if name not in RECIPES:
        return f"no generation recipe for {name}"
    tool_name, source = RECIPES[name]
    tool = tools[tool_name]
    if tool is None:
        return f"skip: {tool_name} not found"
    src = resolve_input(source)
    if source == "srcgen/plain.odt" and not src.exists():
        # Intermediate: plain.odt is itself generated from plain.txt (not committed).
        err = soffice_convert(tool, SRCGEN / "plain.txt", "odt", src)
        if err:
            return f"intermediate plain.odt: {err}"
    if not src.exists():
        return f"input missing: {source}"
    if tool_name == "soffice":
        return soffice_convert(tool, src, "docx", FILES / name)
    return pandoc_convert(tool, src, FILES / name)


def main() -> int:
    manifest: dict[str, dict] = json.loads(MANIFEST_PATH.read_text())
    FILES.mkdir(exist_ok=True)
    tools = {"soffice": shutil.which("soffice"), "pandoc": shutil.which("pandoc")}

    failed: list[tuple[str, str]] = []  # local/download problems -> exit 1
    skipped: list[tuple[str, str]] = []  # generated entries skipped (tool absent)

    for name, entry in manifest.items():
        kind = entry["kind"]
        dest = FILES / name
        if kind == "local":
            src = REPO_ROOT / entry["source"]
            if not src.exists():
                failed.append((name, f"local source missing: {entry['source']}"))
                continue
            shutil.copy(src, dest)
        elif kind == "download":
            if dest.exists() and sha16(dest) == entry["sha256"]:
                continue
            err = download(entry["source"], dest, entry["sha256"])
            if err:
                failed.append((name, err))
        elif kind == "generated":
            if dest.exists():
                continue
            err = generate(name, tools)
            if err and err.startswith("skip: "):
                skipped.append((name, err[len("skip: ") :]))
            elif err:
                failed.append((name, err))
        else:
            failed.append((name, f"unknown kind: {kind}"))

    assembled = sum(1 for name in manifest if (FILES / name).exists())
    print(f"assembled {assembled}/{len(manifest)}", end="")
    if skipped:
        reasons = ", ".join(sorted({r for _, r in skipped}))
        print(f" (skipped {len(skipped)}: {reasons})", end="")
    print()
    for name, err in failed:
        print(f"FAILED: {name}: {err}", file=sys.stderr)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
