"""Tests for atomic save.

``pack_document`` must write the archive to a temp file in the destination's own
directory and promote it with a single ``os.replace``. This guards two properties:

  * A failure mid-pack, or a failed ``validate=True``, never touches the
    destination. This fixes the historic data-loss bug where a failed validation
    ran ``output_file.unlink()`` on the *original* during a save-in-place.
  * A successful pack is a single rename, so the destination is never observed
    half-written — critical inside cloud-synced folders (OneDrive/Dropbox/...).
"""

import shutil
import zipfile
from pathlib import Path

import pytest

from docx_editor.ooxml import pack, unpack_document
from docx_editor.ooxml.pack import pack_document


@pytest.fixture
def unpacked_dir(simple_docx, temp_dir) -> Path:
    """A valid unpacked workspace directory to pack from."""
    d = temp_dir / "unpacked"
    unpack_document(simple_docx, d)
    return d


def _temp_litter(dest: Path) -> list[Path]:
    """Leftover atomic-save temp files next to the destination, if any."""
    return list(dest.parent.glob(f".{dest.name}.*"))


def test_failure_mid_pack_preserves_original(unpacked_dir, simple_docx, temp_dir, monkeypatch):
    """A write error mid-pack leaves the existing destination byte-for-byte intact."""
    dest = temp_dir / "original.docx"
    shutil.copy(simple_docx, dest)
    original_bytes = dest.read_bytes()

    def boom(*args, **kwargs):
        raise OSError("simulated write failure")

    monkeypatch.setattr(pack.shutil, "copyfileobj", boom)

    with pytest.raises(OSError):
        pack_document(unpacked_dir, dest)

    assert dest.read_bytes() == original_bytes
    assert _temp_litter(dest) == []


def test_validation_failure_preserves_original(unpacked_dir, simple_docx, temp_dir, monkeypatch):
    """A failed validation returns False and leaves the original untouched (data-loss regression)."""
    dest = temp_dir / "original.docx"
    shutil.copy(simple_docx, dest)
    original_bytes = dest.read_bytes()

    monkeypatch.setattr(pack, "validate_document", lambda p: False)

    result = pack_document(unpacked_dir, dest, validate=True)

    assert result is False
    assert dest.read_bytes() == original_bytes
    assert _temp_litter(dest) == []


def test_success_is_single_replace_promotion(unpacked_dir, temp_dir, monkeypatch):
    """A successful pack promotes one temp file — in the destination's own dir — via os.replace."""
    dest = temp_dir / "out.docx"

    real_replace = pack.os.replace
    calls: list[tuple[Path, Path]] = []

    def spy_replace(src, dst):
        calls.append((Path(src), Path(dst)))
        return real_replace(src, dst)

    monkeypatch.setattr(pack.os, "replace", spy_replace)

    result = pack_document(unpacked_dir, dest)

    assert result is True
    with zipfile.ZipFile(dest) as zf:
        assert zf.testzip() is None
    assert len(calls) == 1
    src, dst = calls[0]
    assert src.parent == dest.parent  # same volume ⇒ atomic replace
    assert dst == dest
    assert _temp_litter(dest) == []
