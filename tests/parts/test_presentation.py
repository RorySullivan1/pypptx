"""Unit-test suite for `pptx.parts.presentation` module."""

from __future__ import annotations

import pytest

from pptx.exc import SlideError
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
from pptx.package import Package
from pptx.parts.coreprops import CorePropertiesPart
from pptx.parts.presentation import PresentationPart
from pptx.parts.slide import NotesMasterPart, SlideMasterPart, SlidePart
from pptx.presentation import Presentation
from pptx.slide import NotesMaster, Slide, SlideLayout, SlideMaster

from ..unitutil.cxml import element
from ..unitutil.mock import call, class_mock, instance_mock, method_mock, property_mock


class DescribePresentationPart:
    """Unit-test suite for `pptx.parts.presentation.PresentationPart` objects."""

    def it_provides_access_to_its_presentation(self, request):
        prs_ = instance_mock(request, Presentation)
        Presentation_ = class_mock(
            request, "pptx.parts.presentation.Presentation", return_value=prs_
        )
        prs_elm = element("p:presentation")
        prs_part = PresentationPart(None, None, None, prs_elm)

        prs = prs_part.presentation

        Presentation_.assert_called_once_with(prs_elm, prs_part)
        assert prs is prs_

    def it_provides_access_to_its_core_properties(self, request, package_):
        core_properties_ = instance_mock(request, CorePropertiesPart)
        package_.core_properties = core_properties_
        prs_part = PresentationPart(None, None, package_, None)

        assert prs_part.core_properties is core_properties_

    def it_provides_access_to_an_existing_notes_master_part(
        self, notes_master_part_, part_related_by_
    ):
        """This is the first of a two-part test to cover the existing notes master case.

        The notes master not-present case follows.
        """
        prs_part = PresentationPart(None, None, None, None)
        part_related_by_.return_value = notes_master_part_

        notes_master_part = prs_part.notes_master_part

        prs_part.part_related_by.assert_called_once_with(prs_part, RT.NOTES_MASTER)
        assert notes_master_part is notes_master_part_

    def but_it_adds_a_notes_master_part_when_needed(
        self, request, package_, notes_master_part_, part_related_by_, relate_to_
    ):
        """This is the second of a two-part test to cover notes-master-not-present case.

        The notes master present case is just above.
        """
        NotesMasterPart_ = class_mock(request, "pptx.parts.presentation.NotesMasterPart")
        NotesMasterPart_.create_default.return_value = notes_master_part_
        part_related_by_.side_effect = KeyError
        prs_part = PresentationPart(None, None, package_, None)

        notes_master_part = prs_part.notes_master_part

        NotesMasterPart_.create_default.assert_called_once_with(package_)
        relate_to_.assert_called_once_with(prs_part, notes_master_part_, RT.NOTES_MASTER)
        assert notes_master_part is notes_master_part_

    def it_provides_access_to_its_notes_master(self, request, notes_master_part_):
        notes_master_ = instance_mock(request, NotesMaster)
        property_mock(
            request,
            PresentationPart,
            "notes_master_part",
            return_value=notes_master_part_,
        )
        notes_master_part_.notes_master = notes_master_
        prs_part = PresentationPart(None, None, None, None)

        assert prs_part.notes_master is notes_master_

    def it_provides_access_to_a_related_slide(self, request, slide_, related_part_):
        slide_part_ = instance_mock(request, SlidePart, slide=slide_)
        related_part_.return_value = slide_part_
        prs_part = PresentationPart(None, None, None, None)

        slide = prs_part.related_slide("rId42")

        related_part_.assert_called_once_with(prs_part, "rId42")
        assert slide is slide_

    def it_provides_access_to_a_related_master(self, request, slide_master_, related_part_):
        slide_master_part_ = instance_mock(request, SlideMasterPart, slide_master=slide_master_)
        related_part_.return_value = slide_master_part_
        prs_part = PresentationPart(None, None, None, None)

        slide_master = prs_part.related_slide_master("rId42")

        related_part_.assert_called_once_with(prs_part, "rId42")
        assert slide_master is slide_master_

    def it_can_rename_related_slide_parts(self, request, related_part_):
        rIds = tuple("rId%d" % n for n in range(5, 0, -1))
        slide_parts = tuple(instance_mock(request, SlidePart) for _ in range(5))
        related_part_.side_effect = iter(slide_parts)
        prs_part = PresentationPart(None, None, None, None)

        prs_part.rename_slide_parts(rIds)

        assert related_part_.call_args_list == [call(prs_part, rId) for rId in rIds]
        assert [s.partname for s in slide_parts] == [
            PackURI("/ppt/slides/slide%d.xml" % (i + 1)) for i in range(len(rIds))
        ]

    def it_can_save_the_package_to_a_file(self, package_):
        PresentationPart(None, None, package_, None).save("prs.pptx")
        package_.save.assert_called_once_with("prs.pptx")

    def it_can_add_a_new_slide(self, request, package_, slide_part_, slide_, relate_to_):
        slide_layout_ = instance_mock(request, SlideLayout)
        partname = PackURI("/ppt/slides/slide9.xml")
        property_mock(request, PresentationPart, "_next_slide_partname", return_value=partname)
        SlidePart_ = class_mock(request, "pptx.parts.presentation.SlidePart")
        SlidePart_.new.return_value = slide_part_
        relate_to_.return_value = "rId42"
        slide_layout_part_ = slide_layout_.part
        slide_part_.slide = slide_
        prs_part = PresentationPart(None, None, package_, None)

        rId, slide = prs_part.add_slide(slide_layout_)

        SlidePart_.new.assert_called_once_with(partname, package_, slide_layout_part_)
        prs_part.relate_to.assert_called_once_with(prs_part, slide_part_, RT.SLIDE)
        assert rId == "rId42"
        assert slide is slide_

    def it_finds_the_slide_id_of_a_slide_part(self, slide_part_, related_part_):
        prs_elm = element(
            "p:presentation/p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id="
            "b,id=257},p:sldId{r:id=c,id=258})"
        )
        related_part_.side_effect = iter((None, slide_part_, None))
        prs_part = PresentationPart(None, None, None, prs_elm)

        _slide_id = prs_part.slide_id(slide_part_)

        assert related_part_.call_args_list == [
            call(prs_part, "a"),
            call(prs_part, "b"),
        ]
        assert _slide_id == 257

    def it_raises_on_slide_id_not_found(self, slide_part_, related_part_):
        prs_elm = element(
            "p:presentation/p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id="
            "b,id=257},p:sldId{r:id=c,id=258})"
        )
        related_part_.return_value = "not the slide you're looking for"
        prs_part = PresentationPart(None, None, None, prs_elm)

        with pytest.raises(SlideError):
            prs_part.slide_id(slide_part_)

    @pytest.mark.parametrize("is_present", (True, False))
    def it_finds_a_slide_by_slide_id(self, is_present, slide_, slide_part_, related_part_):
        prs_elm = element(
            "p:presentation/p:sldIdLst/(p:sldId{r:id=a,id=256},p:sldId{r:id="
            "b,id=257},p:sldId{r:id=c,id=258})"
        )
        slide_id = 257 if is_present else 666
        expected_value = slide_ if is_present else None
        related_part_.return_value = slide_part_
        slide_part_.slide = slide_
        prs_part = PresentationPart(None, None, None, prs_elm)

        slide = prs_part.get_slide(slide_id)

        assert slide == expected_value

    def it_knows_the_next_slide_partname_to_help(self):
        prs_elm = element("p:presentation/p:sldIdLst/(p:sldId,p:sldId)")
        prs_part = PresentationPart(None, None, None, prs_elm)

        assert prs_part._next_slide_partname == PackURI("/ppt/slides/slide3.xml")

    def it_can_duplicate_a_slide(self, request, package_):
        from pptx.opc.constants import CONTENT_TYPE as CT
        from pptx.oxml.slide import CT_Slide

        # Create a source slide part with a relationship to a layout
        partname = PackURI("/ppt/slides/slide1.xml")
        slide_elm = CT_Slide.new()
        src_slide_part = SlidePart(partname, CT.PML_SLIDE, package_, slide_elm)

        layout_part = instance_mock(request, SlidePart)
        layout_part.partname = PackURI("/ppt/slideLayouts/slideLayout1.xml")
        src_slide_part.relate_to(layout_part, RT.SLIDE_LAYOUT)

        # Package.next_partname provides the new slide partname
        package_.next_partname.return_value = PackURI("/ppt/slides/slide2.xml")

        # Set up the PresentationPart with a valid partname
        prs_elm = element("p:presentation/p:sldIdLst/p:sldId{r:id=rId1,id=256}")
        prs_part = PresentationPart(
            PackURI("/ppt/presentation.xml"), CT.PML_PRESENTATION, package_, prs_elm
        )

        rId, new_slide = prs_part.duplicate_slide(src_slide_part)

        assert rId is not None
        assert new_slide is not None
        package_.next_partname.assert_called_once_with("/ppt/slides/slide%d.xml")
        # Verify the new slide part has a relationship to the same layout
        new_slide_part = new_slide.part
        layout_rel = new_slide_part.part_related_by(RT.SLIDE_LAYOUT)
        assert layout_rel is layout_part

    def it_remaps_rIds_across_all_relationship_attribute_types(self):
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls, qn
        from pptx.parts.presentation import _remap_rIds

        # XML with r:id (hyperlink), r:embed (image), and r:link (linked resource)
        xml_str = (
            '<p:sld %s><p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr/>'
            '<p:sp><p:nvSpPr><p:cNvPr id="2" name="sp1">'
            '<a:hlinkClick r:id="rId1"/>'
            '</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>'
            '<p:pic><p:nvPicPr><p:cNvPr id="3" name="pic1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
            '<p:blipFill><a:blip r:embed="rId2" r:link="rId3"/></p:blipFill>'
            '<p:spPr/></p:pic>'
            '</p:spTree></p:cSld></p:sld>' % nsdecls("p", "a", "r")
        )
        sld = parse_xml(xml_str)

        rId_map = {"rId1": "rId10", "rId2": "rId20", "rId3": "rId30"}
        _remap_rIds(sld, rId_map)

        # Verify r:id was remapped (hyperlink)
        hlinkClick = sld.xpath("//a:hlinkClick")[0]
        assert hlinkClick.get(qn("r:id")) == "rId10"

        # Verify r:embed was remapped (image)
        blip = sld.xpath("//a:blip")[0]
        assert blip.get(qn("r:embed")) == "rId20"

        # Verify r:link was remapped (linked resource)
        assert blip.get(qn("r:link")) == "rId30"

    # fixture components ---------------------------------------------

    @pytest.fixture
    def notes_master_part_(self, request):
        return instance_mock(request, NotesMasterPart)

    @pytest.fixture
    def package_(self, request):
        return instance_mock(request, Package)

    @pytest.fixture
    def part_related_by_(self, request):
        return method_mock(request, PresentationPart, "part_related_by")

    @pytest.fixture
    def relate_to_(self, request):
        return method_mock(request, PresentationPart, "relate_to")

    @pytest.fixture
    def related_part_(self, request):
        return method_mock(request, PresentationPart, "related_part")

    @pytest.fixture
    def slide_(self, request):
        return instance_mock(request, Slide)

    @pytest.fixture
    def slide_master_(self, request):
        return instance_mock(request, SlideMaster)

    @pytest.fixture
    def slide_part_(self, request):
        return instance_mock(request, SlidePart)
