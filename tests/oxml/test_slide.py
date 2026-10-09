"""Unit-test suite for `pptx.oxml.slide` module."""

from __future__ import annotations

import pytest

from pptx.exc import InvalidValueError
from pptx.oxml.slide import (
    CT_HandoutMaster,
    CT_NotesMaster,
    CT_NotesSlide,
    CT_SlideLayout,
    CT_SlideLayoutIdList,
)

from ..unitutil.cxml import element, xml
from ..unitutil.file import snippet_text


class DescribeCT_HandoutMaster:
    """Unit-test suite for `pptx.oxml.slide.CT_HandoutMaster` objects."""

    def it_is_the_element_class_for_a_handout_master(self):
        handoutMaster = element("p:handoutMaster/(p:cSld/p:spTree,p:clrMap,p:hf)")

        assert isinstance(handoutMaster, CT_HandoutMaster)
        assert handoutMaster.cSld.spTree is not None
        assert handoutMaster.hf is not None


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

    def it_can_renew_its_creation_id(self):
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls
        from pptx.oxml.slide import CREATION_ID_EXT_URI

        cSld = parse_xml(
            '<p:cSld %s><p:spTree/><p:extLst><p:ext uri="%s"><p14:creationId val="42"/></p:ext>'
            "</p:extLst></p:cSld>" % (nsdecls("p", "p14"), CREATION_ID_EXT_URI)
        )

        cSld.renew_creation_id()

        assert cSld.creation_id not in (None, 42)

    def but_it_does_not_add_one_when_renewing(self):
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls

        cSld = parse_xml("<p:cSld %s><p:spTree/></p:cSld>" % nsdecls("p"))

        cSld.renew_creation_id()

        assert cSld.creation_id is None


class DescribeCT_SlideLayout:
    """Unit-test suite for `pptx.oxml.slide.CT_SlideLayout` objects."""

    def it_can_create_a_new_layout_as_PowerPoint_writes_a_user_made_one(self):
        sldLayout = CT_SlideLayout.new("Quote")

        assert isinstance(sldLayout, CT_SlideLayout)
        assert sldLayout.cSld.name == "Quote"
        assert sldLayout.preserve is True
        assert sldLayout.userDrawn is True
        assert sldLayout.get("type") is None
        assert list(sldLayout.cSld.spTree.iter_shape_elms()) == []
        assert sldLayout.xpath("./p:clrMapOvr/a:masterClrMapping")

    def and_it_has_its_flags_off_by_default(self):
        sldLayout = element("p:sldLayout/p:cSld/p:spTree")

        assert sldLayout.preserve is False
        assert sldLayout.userDrawn is False


class DescribeCT_SlideLayoutIdList:
    """Unit-test suite for `pptx.oxml.slide.CT_SlideLayoutIdList` objects."""

    def it_can_append_an_entry_with_an_id(self):
        sldLayoutIdLst = element("p:sldLayoutIdLst/p:sldLayoutId{r:id=rId1,id=2147483649}")
        assert isinstance(sldLayoutIdLst, CT_SlideLayoutIdList)

        sldLayoutIdLst.add_sldLayoutId("rId2", id=2147483650)

        assert sldLayoutIdLst.xml == xml(
            "p:sldLayoutIdLst/(p:sldLayoutId{r:id=rId1,id=2147483649}"
            ",p:sldLayoutId{r:id=rId2,id=2147483650})"
        )

    def and_it_can_insert_one_after_a_given_entry(self):
        sldLayoutIdLst = element(
            "p:sldLayoutIdLst/(p:sldLayoutId{r:id=rId1},p:sldLayoutId{r:id=rId2})"
        )
        first = sldLayoutIdLst.sldLayoutId_lst[0]

        sldLayoutIdLst.add_sldLayoutId("rId3", id=2147483651, after=first)

        assert [e.rId for e in sldLayoutIdLst.sldLayoutId_lst] == ["rId1", "rId3", "rId2"]

    def but_it_rejects_an_id_outside_the_shared_range(self):
        sldLayoutIdLst = element("p:sldLayoutIdLst")

        with pytest.raises(InvalidValueError):
            sldLayoutIdLst.add_sldLayoutId("rId1", id=256)
