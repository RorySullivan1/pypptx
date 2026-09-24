"""Unit-test suite for `pptx.oxml.slide` module."""

from __future__ import annotations

from pptx.oxml.slide import CT_NotesMaster, CT_NotesSlide

from ..unitutil.file import snippet_text


class DescribeCT_NotesMaster:
    """Unit-test suite for `pptx.oxml.slide.CT_NotesMaster` objects."""

    def it_can_create_a_default_notesMaster_element(self):
        notesMaster = CT_NotesMaster.new_default()
        assert notesMaster.xml == snippet_text("default-notesMaster")


class DescribeCT_NotesSlide:
    """Unit-test suite for `pptx.oxml.slide.CT_NotesSlide` objects."""

    def it_can_create_a_new_notes_element(self):
        notes = CT_NotesSlide.new()
        assert notes.xml == snippet_text("default-notes")


class DescribeCT_Slide:
    """Unit-test suite for the modern-comment extension of `pptx.oxml.slide.CT_Slide`."""

    def _sld(self, *ext_uris: str):
        """`p:sld` whose `p:extLst` has one `p:ext` per uri; "C" is the comment-rel extension."""
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls
        from pptx.oxml.slide import COMMENT_REL_EXT_URI

        exts = "".join(
            '<p:ext uri="%s"><p188:commentRel r:id="rId9"/></p:ext>' % COMMENT_REL_EXT_URI
            if uri == "C"
            else '<p:ext uri="%s"/>' % uri
            for uri in ext_uris
        )
        extLst = "<p:extLst>%s</p:extLst>" % exts if ext_uris else ""
        return parse_xml(
            "<p:sld %s><p:cSld><p:spTree/></p:cSld>%s</p:sld>" % (nsdecls("p", "r", "p188"), extLst)
        )

    def it_knows_the_rId_of_its_comment_relationship(self):
        assert self._sld("{X}", "C").comment_rel_rId == "rId9"

    def but_it_has_none_without_the_comment_extension(self):
        assert self._sld().comment_rel_rId is None
        assert self._sld("{X}").comment_rel_rId is None

    def it_can_remove_its_comment_relationship_extension(self):
        sld = self._sld("{X}", "C")

        sld.remove_comment_rel()

        assert sld.comment_rel_rId is None
        assert [ext.get("uri") for ext in sld.xpath("./p:extLst/p:ext")] == ["{X}"]

    def and_it_removes_the_extLst_when_that_empties_it(self):
        sld = self._sld("C")

        sld.remove_comment_rel()

        assert sld.xpath("./p:extLst") == []
