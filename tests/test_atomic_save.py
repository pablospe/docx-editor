"""Tests for atomic save.

``pack_document`` must write the archive to a temp file in the destination's own
directory and promote it with a single ``os.replace``. This guards two properties:

  * A failure mid-pack, or a failed ``validate=True``, never touches the
    destination. This fixes the historic data-loss bug where a failed validation
    ran ``output_file.unlink()`` on the *original* during a save-in-place.
  * A successful pack is a single rename, so the destination is never observed
    half-written — critical inside cloud-synced folders (OneDrive/Dropbox/...).

Because ``os.replace`` swaps in the temp file's *inode*, anything riding on the
destination's inode must be carried over deliberately: these tests also pin the
destination's permissions and the handling of a symlinked destination.
"""

import os
import shutil
import stat
import subprocess
import zipfile
from pathlib import Path

import pytest

from docx_editor import Document
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


class TestAtomicFailurePaths:
    """A failed pack must leave the existing destination exactly as it was."""

    def test_failure_mid_pack_preserves_original(self, unpacked_dir, simple_docx, temp_dir, monkeypatch):
        """A write error mid-archive leaves the existing destination byte-for-byte intact."""
        dest = temp_dir / "original.docx"
        shutil.copy(simple_docx, dest)
        original_bytes = dest.read_bytes()

        # Fail on the *second* member so the archive is genuinely half-written, and
        # record that a temp file existed at that moment — proving the failure struck
        # after the temp was created, not before (which would pass vacuously).
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

    def test_validation_failure_preserves_original(self, unpacked_dir, simple_docx, temp_dir, monkeypatch):
        """A failed validation returns False and leaves the original untouched.

        Regression for the data-loss bug: validation used to run against the
        destination and unlink() it on failure.
        """
        dest = temp_dir / "original.docx"
        shutil.copy(simple_docx, dest)
        original_bytes = dest.read_bytes()

        monkeypatch.setattr(pack, "validate_document", lambda p, **kw: False)

        result = pack_document(unpacked_dir, dest, validate=True)

        assert result is False
        assert dest.read_bytes() == original_bytes
        assert _temp_litter(dest) == []

    def test_validation_receives_real_extension(self, unpacked_dir, temp_dir, monkeypatch):
        """The temp file is named .tmp, so the real suffix must reach soffice explicitly.

        Otherwise validate_document() picks its export filter from ".tmp", falling to
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


class TestAtomicPromotion:
    """A successful pack is a single rename from a temp file in the destination's dir."""

    def test_success_is_single_replace_promotion(self, unpacked_dir, temp_dir, monkeypatch):
        """Exactly one os.replace, sourced from the destination's own directory."""
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

    def test_long_destination_filename(self, unpacked_dir, temp_dir):
        """A near-max-length filename must still save (the temp name is derived from it)."""
        dest = temp_dir / ("a" * 245 + ".docx")  # 250 bytes, legal on ext4/APFS
        pack_document(unpacked_dir, dest)

        with zipfile.ZipFile(dest) as zf:
            assert zf.testzip() is None

    def test_long_non_ascii_destination_filename(self, unpacked_dir, temp_dir):
        """NAME_MAX is a byte budget, not a character count.

        A legal non-ASCII name near the limit must not overflow the temp name — 80
        Chinese characters is only 80 chars but 240 bytes.
        """
        dest = temp_dir / ("报" * 80 + ".docx")  # 245 bytes, 85 characters
        pack_document(unpacked_dir, dest)

        with zipfile.ZipFile(dest) as zf:
            assert zf.testzip() is None

    def test_undecodable_destination_filename(self, unpacked_dir, temp_dir):
        """Linux filenames are bytes; an undecodable one arrives as surrogates.

        Deriving the temp name with str.encode() (strict UTF-8) would reject those and
        make the document unsaveable.
        """
        dest = temp_dir / os.fsdecode(b"caf\xe9.docx")
        pack_document(unpacked_dir, dest)

        with zipfile.ZipFile(dest) as zf:
            assert zf.testzip() is None


class TestAtomicInodeState:
    """os.replace swaps the inode — what rides on it must be carried over."""

    def test_existing_destination_keeps_its_permissions(self, unpacked_dir, simple_docx, temp_dir):
        """Without this, mkstemp's 0600 silently strips group/world access on every save."""
        dest = temp_dir / "shared.docx"
        shutil.copy(simple_docx, dest)
        dest.chmod(0o664)

        pack_document(unpacked_dir, dest)

        assert stat.S_IMODE(dest.stat().st_mode) == 0o664

    def test_read_only_destination_is_refused(self, unpacked_dir, simple_docx, temp_dir):
        """A write-protected document must not be silently overwritten.

        rename(2) only needs write permission on the *directory*, so an atomic promotion
        would happily replace a 0444 file — something the pre-atomic in-place write
        could never do, and that Windows still refuses.
        """
        dest = temp_dir / "readonly.docx"
        shutil.copy(simple_docx, dest)
        original_bytes = dest.read_bytes()
        dest.chmod(0o444)

        with pytest.raises(PermissionError):
            pack_document(unpacked_dir, dest)

        assert dest.read_bytes() == original_bytes
        assert _temp_litter(dest) == []

    def test_new_destination_respects_umask(self, unpacked_dir, temp_dir):
        """A brand-new file gets the mode a plain open() would have produced, not 0600."""
        dest = temp_dir / "fresh.docx"
        old_umask = os.umask(0o022)
        try:
            pack_document(unpacked_dir, dest)
        finally:
            os.umask(old_umask)

        assert stat.S_IMODE(dest.stat().st_mode) == 0o644

    @pytest.mark.skipif(shutil.which("setfacl") is None, reason="setfacl not available")
    def test_saves_into_a_directory_with_a_restrictive_default_acl(self, unpacked_dir, simple_docx, temp_dir):
        """A default ACL can leave a newly created file not writable by its owner.

        Reopening the temp file by name to write or fsync it would then fail, breaking
        a save that worked before this feature existed. Writing through the fd it was
        created with (opened O_RDWR before the mode applied) is what keeps this working.
        """
        acl_dir = temp_dir / "acl"
        acl_dir.mkdir()
        subprocess.run(["setfacl", "-d", "-m", "u::r--", str(acl_dir)], check=True)
        dest = acl_dir / "doc.docx"
        shutil.copy(simple_docx, dest)
        original_bytes = dest.read_bytes()

        pack_document(unpacked_dir, dest)

        assert dest.read_bytes() != original_bytes
        with zipfile.ZipFile(dest) as zf:
            assert zf.testzip() is None

    def test_symlinked_destination_is_followed(self, unpacked_dir, simple_docx, temp_dir):
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

    def test_saving_through_a_symlink_keeps_the_workspace_in_sync(self, clean_workspace, temp_dir):
        """Saving via another name for the source still writes the source.

        So the workspace's recorded mtime has to be refreshed, or the next open() sees a
        workspace that looks stale and refuses.
        """
        link = temp_dir / "alias.docx"
        link.symlink_to(clean_workspace)

        doc = Document.open(clean_workspace)
        doc.save(link)  # a different spelling of the same file
        doc.close()

        reopened = Document.open(clean_workspace)  # must not raise WorkspaceSyncError
        reopened.close()
