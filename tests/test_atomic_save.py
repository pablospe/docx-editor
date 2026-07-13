"""Tests for atomic save.

``pack_document`` must write the archive to a temp file in the destination's own
directory and promote it with a single ``os.replace``. This guards two properties:

  * A failure mid-pack, or a failed ``validate=True``, never touches the
    destination. This fixes the historic data-loss bug where a failed validation
    ran ``output_file.unlink()`` on the *original* during a save-in-place.
  * A successful pack is a single rename, so the destination is never observed
    half-written — critical inside cloud-synced folders (OneDrive/Dropbox/...).

Because ``os.replace`` swaps in the temp file's *inode*, anything riding on the
destination's inode must be carried over: these tests also pin the destination's
permissions and the handling of a symlinked destination.
"""

import os
import shutil
import stat
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
    """A write error mid-archive leaves the existing destination byte-for-byte intact."""
    dest = temp_dir / "original.docx"
    shutil.copy(simple_docx, dest)
    original_bytes = dest.read_bytes()

    # Fail on the *second* member so the archive is genuinely half-written, and
    # record that a temp file existed at that moment — proving the failure happened
    # after promotion started, not before the temp file was ever created (which
    # would make this test pass vacuously).
    state = {"calls": 0, "temp_existed_mid_write": False}
    real_copyfileobj = shutil.copyfileobj

    def boom(src, dst, *args, **kwargs):
        state["calls"] += 1
        if state["calls"] < 2:
            return real_copyfileobj(src, dst, *args, **kwargs)
        state["temp_existed_mid_write"] = bool(_temp_litter(dest))
        raise OSError("simulated write failure")

    monkeypatch.setattr(pack.shutil, "copyfileobj", boom)

    with pytest.raises(OSError):
        pack_document(unpacked_dir, dest)

    assert state["temp_existed_mid_write"], "failure did not occur mid-archive"
    assert dest.read_bytes() == original_bytes
    assert _temp_litter(dest) == []


def test_validation_failure_preserves_original(unpacked_dir, simple_docx, temp_dir, monkeypatch):
    """A failed validation returns False and leaves the original untouched (data-loss regression)."""
    dest = temp_dir / "original.docx"
    shutil.copy(simple_docx, dest)
    original_bytes = dest.read_bytes()

    monkeypatch.setattr(pack, "validate_document", lambda p, **kw: False)

    result = pack_document(unpacked_dir, dest, validate=True)

    assert result is False
    assert dest.read_bytes() == original_bytes
    assert _temp_litter(dest) == []


def test_validation_receives_real_extension(unpacked_dir, temp_dir, monkeypatch):
    """The temp file is named .tmp, so the real suffix must be passed to soffice explicitly.

    Otherwise validate_document() picks its export filter from ".tmp" and falls to
    the default (Writer) filter — wrong for .pptx/.xlsx.
    """
    dest = temp_dir / "out.docx"
    seen = {}

    def spy_validate(path, suffix=None):
        seen["suffix"] = suffix
        return True

    monkeypatch.setattr(pack, "validate_document", spy_validate)

    assert pack_document(unpacked_dir, dest, validate=True) is True
    assert seen["suffix"] == ".docx"


def test_success_is_single_replace_promotion(unpacked_dir, temp_dir, monkeypatch):
    """A successful pack promotes one temp file — in the destination's own dir — via os.replace."""
    dest = temp_dir / "out.docx"
    resolved = Path(os.path.realpath(dest))

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
    assert src.parent == resolved.parent  # same volume ⇒ atomic replace
    assert dst == resolved
    assert _temp_litter(dest) == []


def test_existing_destination_keeps_its_permissions(unpacked_dir, simple_docx, temp_dir):
    """os.replace swaps the inode — the destination's mode must survive the save.

    Without this, mkstemp's 0600 would silently strip group/world access from every
    saved document.
    """
    dest = temp_dir / "shared.docx"
    shutil.copy(simple_docx, dest)
    dest.chmod(0o664)

    pack_document(unpacked_dir, dest)

    assert stat.S_IMODE(dest.stat().st_mode) == 0o664


def test_new_destination_respects_umask(unpacked_dir, temp_dir):
    """A brand-new file gets the mode a plain open() would have produced, not 0600."""
    dest = temp_dir / "fresh.docx"
    old_umask = os.umask(0o022)
    try:
        pack_document(unpacked_dir, dest)
    finally:
        os.umask(old_umask)

    assert stat.S_IMODE(dest.stat().st_mode) == 0o644


def test_symlinked_destination_is_followed(unpacked_dir, simple_docx, temp_dir):
    """Saving to a symlink must update the file it points at, not replace the link."""
    real = temp_dir / "real.docx"
    shutil.copy(simple_docx, real)
    original_bytes = real.read_bytes()
    link = temp_dir / "link.docx"
    link.symlink_to(real)

    pack_document(unpacked_dir, link)

    assert link.is_symlink(), "the symlink itself was replaced"
    assert real.read_bytes() != original_bytes, "the real document was never updated"
    with zipfile.ZipFile(real) as zf:
        assert zf.testzip() is None
    assert _temp_litter(real) == []


def test_long_destination_filename(unpacked_dir, temp_dir):
    """A near-max-length filename must still save (the temp name is derived from it)."""
    dest = temp_dir / ("a" * 245 + ".docx")  # 250 chars, legal on ext4/APFS
    pack_document(unpacked_dir, dest)

    with zipfile.ZipFile(dest) as zf:
        assert zf.testzip() is None
