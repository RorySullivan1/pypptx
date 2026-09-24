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

    def it_can_point_its_comment_relationship_at_an_rId(self):
        sld = self._sld("{X}", "C")

        sld.set_comment_rel("rId4")

        assert sld.comment_rel_rId == "rId4"
        assert len(sld.xpath("./p:extLst/p:ext")) == 2

    def and_it_adds_the_extLst_after_the_other_children_when_there_is_none(self):
        sld = self._sld()

        sld.set_comment_rel("rId4")

        assert sld.comment_rel_rId == "rId4"
        assert sld[-1].tag.endswith("}extLst")


class DescribeCT_CommonSlideData:
    """Unit-test suite for the `p14:creationId` extension of `CT_CommonSlideData`."""

    def it_knows_its_creation_id(self):
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls
        from pptx.oxml.slide import CREATION_ID_EXT_URI

        cSld = parse_xml(
            '<p:cSld %s><p:spTree/><p:extLst><p:ext uri="%s"><p14:creationId val="42"/></p:ext>'
            "</p:extLst></p:cSld>" % (nsdecls("p", "p14"), CREATION_ID_EXT_URI)
        )

        assert cSld.creation_id == 42
        assert cSld.get_or_add_creation_id() == 42

    def it_adds_a_creation_id_when_there_is_none(self):
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls

        cSld = parse_xml("<p:cSld %s><p:spTree/></p:cSld>" % nsdecls("p"))

        creation_id = cSld.get_or_add_creation_id()

        assert 1 <= creation_id <= 0xFFFFFFFF
        assert cSld.creation_id == creation_id
        assert cSld[-1].tag.endswith("}extLst")
