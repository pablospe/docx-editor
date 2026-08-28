"""Tests for the two settings.xml halves: the track-changes switch and protection.

``word/settings.xml`` carries a document's application-level state. Two elements
there matter to this library:

* ``<w:trackChanges/>`` — Word's *Track Changes* switch. Our revisions are in
  document.xml and show up without it, but a recipient who keeps typing in Word
  produces *untracked* edits unless the switch is on. ``save()`` turns it on
  when the document carries a revision we authored.
* ``<w:documentProtection>`` — Word's *Restrict Editing*. When enforced with a
  body-locking mode the document is refused at open, so a protected document is
  never silently edited.

Assertions read the saved .docx with zipfile + minidom rather than the workspace,
so they pin what the reviewer's Word actually opens.
"""

import warnings
import zipfile
from pathlib import Path

import defusedxml.minidom
import pytest
from conftest import find_ref, replace_docx_parts

from docx_editor import Document, DocumentProtectedError

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
)

# A Word-shaped settings.xml: the sequence around trackChanges' slot, plus a
# trailing extension element from a namespace the schema list does not know.
WORD_SHAPED_BODY = (
    '<w:zoom w:val="bestFit"/>'
    '<w:proofState w:spelling="clean" w:grammar="clean"/>'
    '<w:defaultTabStop w:val="720"/>'
    '<w:characterSpacingControl w:val="doNotCompress"/>'
    "<w:compat/>"
    '<w15:docId w15:val="{7A0B1C2D-0000-0000-0000-000000000000}"/>'
)

AUTHOR_A = "Reviewer A"


def settings_xml(body: str) -> str:
    """A settings.xml part with ``body`` as the children of ``w:settings``."""
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}>{body}</w:settings>'


def with_settings(src: Path, dest: Path, body: str) -> Path:
    """Copy ``src`` to ``dest`` with a settings.xml built from ``body``."""
    replace_docx_parts(src, dest, {"word/settings.xml": settings_xml(body)})
    return dest


def saved_settings(path: Path):
    """Parse ``word/settings.xml`` out of a saved .docx."""
    with zipfile.ZipFile(path) as archive:
        return defusedxml.minidom.parseString(archive.read("word/settings.xml"))


def track_changes_elements(path: Path) -> list:
    """Every ``w:trackChanges`` element in a saved document's settings.xml."""
    return saved_settings(path).getElementsByTagName("w:trackChanges")


def edit_and_save(source: Path, dest: Path, old: str = "quick", new: str = "slow", **save_kwargs) -> Path:
    """Open ``source``, make one tracked replacement, save to ``dest``."""
    doc = Document.open(source)
    try:
        doc.replace(old, new, paragraph=find_ref(doc, "brown fox"))
        return doc.save(dest, **save_kwargs)
    finally:
        doc.close()


class TestTrackChangesFlag:
    """save() writes Word's track-changes switch when we leave a redline."""

    def test_redline_turns_the_flag_on(self, temp_docx, temp_dir):
        """A tracked edit puts exactly one bare w:trackChanges in the output."""
        out = edit_and_save(temp_docx, temp_dir / "out.docx")

        elements = track_changes_elements(out)
        assert len(elements) == 1
        assert not elements[0].getAttribute("w:val")

    def test_open_and_save_without_edits_leaves_it_off(self, temp_docx, temp_dir):
        """We do not flip a switch on a document we did not redline."""
        doc = Document.open(temp_docx)
        out = doc.save(temp_dir / "out.docx")
        doc.close()

        assert track_changes_elements(out) == []

    def test_track_changes_false_skips_the_write(self, temp_docx, temp_dir):
        """save(track_changes=False) leaves settings.xml alone despite the redline."""
        out = edit_and_save(temp_docx, temp_dir / "out.docx", track_changes=False)

        assert track_changes_elements(out) == []

    def test_track_changes_true_writes_without_revisions(self, temp_docx, temp_dir):
        """save(track_changes=True) writes the flag even with nothing to redline."""
        doc = Document.open(temp_docx)
        out = doc.save(temp_dir / "out.docx", track_changes=True)
        doc.close()

        assert len(track_changes_elements(out)) == 1

    def test_existing_flag_is_not_duplicated(self, temp_docx, temp_dir):
        """A document that already has the flag keeps exactly one."""
        source = with_settings(temp_docx, temp_dir / "already_on.docx", "<w:zoom/><w:trackChanges/>")
        out = edit_and_save(source, temp_dir / "out.docx")

        assert len(track_changes_elements(out)) == 1

    @pytest.mark.parametrize("off", ["false", "0", "off"])
    def test_explicit_off_is_respected_and_warned_about(self, temp_docx, temp_dir, off):
        """An explicit off stays as it is; the save says what that costs.

        ST_OnOff has three falsy spellings and producers do not all pick the
        same one, so all three have to read as "the author turned this off".
        """
        source = with_settings(temp_docx, temp_dir / "off.docx", f'<w:zoom/><w:trackChanges w:val="{off}"/>')

        doc = Document.open(source)
        try:
            doc.replace("quick", "slow", paragraph=find_ref(doc, "brown fox"))
            with pytest.warns(UserWarning, match="will not be tracked"):
                out = doc.save(temp_dir / "out.docx")
        finally:
            doc.close()

        elements = track_changes_elements(out)
        assert len(elements) == 1
        assert elements[0].getAttribute("w:val") == off

    def test_explicit_true_overrides_the_documents_explicit_off(self, temp_docx, temp_dir):
        """track_changes=True outranks w:val="false" — the caller said it in as many words."""
        source = with_settings(temp_docx, temp_dir / "off.docx", '<w:zoom/><w:trackChanges w:val="false"/>')

        doc = Document.open(source)
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("error", UserWarning)
                out = doc.save(temp_dir / "out.docx", track_changes=True)
        finally:
            doc.close()

        elements = track_changes_elements(out)
        assert len(elements) == 1
        assert not elements[0].getAttribute("w:val")

    def test_flag_lands_in_its_schema_slot(self, temp_docx, temp_dir):
        """CT_Settings is a sequence: after w:proofState, before w:defaultTabStop.

        The unknown trailing w15:docId must not become the anchor — anchoring on
        it would put the flag at the end of the sequence, out of its slot.
        """
        source = with_settings(temp_docx, temp_dir / "word_shaped.docx", WORD_SHAPED_BODY)
        out = edit_and_save(source, temp_dir / "out.docx")

        root = saved_settings(out).documentElement
        names = [c.tagName for c in root.childNodes if c.nodeType == c.ELEMENT_NODE]
        assert names.index("w:proofState") + 1 == names.index("w:trackChanges")
        assert names.index("w:trackChanges") + 1 == names.index("w:defaultTabStop")

    def test_missing_settings_part_still_saves(self, temp_docx, temp_dir):
        """A document with no settings.xml saves cleanly, just without a flag."""
        source = temp_dir / "no_settings.docx"
        replace_docx_parts(temp_docx, source, {"word/settings.xml": None})

        out = edit_and_save(source, temp_dir / "out.docx")

        with zipfile.ZipFile(out) as archive:
            assert "word/settings.xml" not in archive.namelist()

    def test_accepted_edits_leave_the_flag_off(self, temp_docx, temp_dir):
        """accept_all() removes the redline, so the predicate answers honestly."""
        doc = Document.open(temp_docx)
        try:
            doc.replace("quick", "slow", paragraph=find_ref(doc, "brown fox"))
            doc.accept_all()
            with warnings.catch_warnings():
                warnings.simplefilter("error", UserWarning)
                out = doc.save(temp_dir / "out.docx")
        finally:
            doc.close()

        assert track_changes_elements(out) == []

    def test_foreign_revisions_alone_leave_the_flag_off(self, temp_docx, temp_dir):
        """Passing a foreign redline through untouched is not our redline."""
        source = temp_dir / "foreign.docx"
        with zipfile.ZipFile(temp_docx) as archive:
            document_xml = archive.read("word/document.xml").decode("utf-8")
        foreign = (
            f'<w:p><w:ins w:id="900" w:author="{AUTHOR_A}" w:date="2024-01-01T00:00:00Z">'
            "<w:r><w:t>Their clause.</w:t></w:r></w:ins></w:p>"
        )
        replace_docx_parts(
            temp_docx,
            source,
            {"word/document.xml": document_xml.replace("</w:body>", f"{foreign}</w:body>")},
        )

        doc = Document.open(source)
        out = doc.save(temp_dir / "out.docx")
        doc.close()

        assert track_changes_elements(out) == []

    def test_reopening_a_flagged_document_keeps_one_flag(self, temp_docx, temp_dir):
        """Round two of the same redline does not stack a second element."""
        first = edit_and_save(temp_docx, temp_dir / "first.docx")
        second = edit_and_save(first, temp_dir / "second.docx", old="lazy", new="sleepy")

        assert len(track_changes_elements(second)) == 1


class TestDocumentProtection:
    """Enforced editing protection is refused at open, with one bypass."""

    def protected(self, temp_docx, temp_dir, mode: str, enforcement: str | None = "1") -> Path:
        """A copy of the document protected with ``mode``."""
        attrs = f'w:edit="{mode}"'
        if enforcement is not None:
            attrs += f' w:enforcement="{enforcement}"'
        name = f"protected_{mode}_{enforcement}.docx"
        return with_settings(temp_docx, temp_dir / name, f"<w:zoom/><w:documentProtection {attrs}/>")

    @pytest.mark.parametrize("mode", ["readOnly", "forms", "comments"])
    def test_body_locking_modes_raise(self, temp_docx, temp_dir, mode):
        """readOnly, forms and comments all lock the body text, so open refuses."""
        source = self.protected(temp_docx, temp_dir, mode)

        with pytest.raises(DocumentProtectedError) as exc_info:
            Document.open(source)

        assert exc_info.value.mode == mode
        assert exc_info.value.path == source.resolve()

    @pytest.mark.parametrize("mode", ["readOnly", "forms", "comments"])
    def test_message_names_the_mode_and_the_way_out(self, temp_docx, temp_dir, mode):
        """The message has to carry both halves: what is wrong, and what to do."""
        source = self.protected(temp_docx, temp_dir, mode)

        with pytest.raises(DocumentProtectedError) as exc_info:
            Document.open(source)

        message = str(exc_info.value)
        assert "Restrict Editing" in message
        assert "allow_protected=True" in message
        if mode == "comments":
            assert "add_comment()" in message

    def test_refused_open_leaves_no_workspace_or_lock(self, temp_docx, temp_dir, isolated_workspace_base):
        """A raise inside __init__ must not lock the document against its own retry."""
        source = self.protected(temp_docx, temp_dir, "readOnly")

        with pytest.raises(DocumentProtectedError):
            Document.open(source)

        # The workspace dir and its .lock sidecar are siblings under the base,
        # so an empty base is both halves of "nothing leaked".
        leftovers = sorted(entry.name for entry in isolated_workspace_base.iterdir())
        assert leftovers == [], f"refused open left {leftovers} behind"

        # And the retry that the bypass makes possible actually works.
        doc = Document.open(source, allow_protected=True)
        doc.close()

    def test_tracked_changes_mode_never_raises(self, temp_docx, temp_dir):
        """Enforced trackedChanges asks for exactly what this library does."""
        source = self.protected(temp_docx, temp_dir, "trackedChanges")

        with warnings.catch_warnings():
            warnings.simplefilter("error", UserWarning)
            doc = Document.open(source)
        try:
            doc.replace("quick", "slow", paragraph=find_ref(doc, "brown fox"))
            out = doc.save(temp_dir / "out.docx")
        finally:
            doc.close()

        assert len(track_changes_elements(out)) == 1

    @pytest.mark.parametrize("enforcement", ["1", "true", "on"])
    def test_every_enforcement_spelling_counts(self, temp_docx, temp_dir, enforcement):
        """ST_OnOff is truthy in three spellings; a producer may write any of them."""
        source = self.protected(temp_docx, temp_dir, "readOnly", enforcement=enforcement)

        with pytest.raises(DocumentProtectedError):
            Document.open(source)

    @pytest.mark.parametrize("enforcement", ["0", "false", "off", None])
    def test_unenforced_protection_opens_silently(self, temp_docx, temp_dir, enforcement):
        """A configured but switched-off protection is editable in Word, so here too."""
        source = self.protected(temp_docx, temp_dir, "readOnly", enforcement=enforcement)

        doc = Document.open(source)
        try:
            assert doc.paragraph_count() > 0
        finally:
            doc.close()

    def test_enforced_protection_without_a_mode_opens(self, temp_docx, temp_dir):
        """No w:edit is no mode to police — the guard is about locked content."""
        source = with_settings(
            temp_docx, temp_dir / "no_mode.docx", '<w:zoom/><w:documentProtection w:enforcement="1"/>'
        )

        doc = Document.open(source)
        try:
            assert doc.paragraph_count() > 0
        finally:
            doc.close()

    def test_bypass_opens_and_saves(self, temp_docx, temp_dir):
        """allow_protected=True is a real open: edit and save work through it."""
        source = self.protected(temp_docx, temp_dir, "readOnly")

        doc = Document.open(source, allow_protected=True)
        try:
            doc.replace("quick", "slow", paragraph=find_ref(doc, "brown fox"))
            out = doc.save(temp_dir / "out.docx")
        finally:
            doc.close()

        assert len(saved_settings(out).getElementsByTagName("w:documentProtection")) == 1

    def test_bypass_allows_commenting_a_comments_protected_document(self, temp_docx, temp_dir):
        """The mode that permits comments still permits them once bypassed."""
        source = self.protected(temp_docx, temp_dir, "comments")

        doc = Document.open(source, allow_protected=True)
        try:
            comment_id = doc.add_comment("quick brown fox", "Please review")
            out = doc.save(temp_dir / "out.docx")
        finally:
            doc.close()

        assert comment_id is not None
        with zipfile.ZipFile(out) as archive:
            assert "word/comments.xml" in archive.namelist()
