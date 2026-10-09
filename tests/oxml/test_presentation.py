# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.oxml.presentation` module."""

from __future__ import annotations

from typing import cast

import pytest

from pptx.oxml.presentation import CT_Presentation, CT_SlideIdList, CT_SlideMasterIdList

from ..unitutil.cxml import element, xml


class DescribeCT_Presentation:
    """Unit-test suite for `pptx.oxml.presentation.CT_Presentation` objects."""

    def it_provides_access_to_its_embeddedFontLst(self):
        presentation = cast(
            CT_Presentation,
            element("p:presentation/p:embeddedFontLst/p:embeddedFont/p:font{typeface=Calibri}"),
        )
        assert presentation.embeddedFontLst is not None
        assert presentation.embeddedFontLst.embeddedFont_lst[0].font.typeface == "Calibri"

    def it_returns_None_for_embeddedFontLst_when_not_present(self):
        presentation = cast(CT_Presentation, element("p:presentation"))
        assert presentation.embeddedFontLst is None

    def it_adds_embeddedFontLst_between_notesSz_and_kinsoku(self):
        presentation = cast(CT_Presentation, element("p:presentation/(p:notesSz,p:kinsoku)"))

        presentation.get_or_add_embeddedFontLst()

        assert [child.tag.split("}")[1] for child in presentation] == [
            "notesSz",
            "embeddedFontLst",
            "kinsoku",
        ]

    @pytest.mark.parametrize(
        ("prs_cxml", "expected_value"),
        [
            ("p:presentation", None),
            ("p:presentation{embedTrueTypeFonts=1}", True),
            ("p:presentation{embedTrueTypeFonts=0}", False),
        ],
    )
    def it_knows_whether_it_embeds_truetype_fonts(self, prs_cxml: str, expected_value):
        presentation = cast(CT_Presentation, element(prs_cxml))
        assert presentation.embedTrueTypeFonts is expected_value

    def it_can_change_whether_it_embeds_truetype_fonts(self):
        presentation = cast(CT_Presentation, element("p:presentation"))
        presentation.embedTrueTypeFonts = False
        assert presentation.xml == xml("p:presentation{embedTrueTypeFonts=0}")


class DescribeCT_SlideIdList:
    """Unit-test suite for `pptx.oxml.presentation.CT_SlideIdLst` objects."""

    def it_can_add_a_sldId_element_as_a_child(self):
        sldIdLst = cast(CT_SlideIdList, element("p:sldIdLst/p:sldId{r:id=rId4,id=256}"))

        sldIdLst.add_sldId("rId1")

        assert sldIdLst.xml == xml(
            "p:sldIdLst/(p:sldId{r:id=rId4,id=256},p:sldId{r:id=rId1,id=257})"
        )

    @pytest.mark.parametrize(
        ("sldIdLst_cxml", "expected_value"),
        [
            ("p:sldIdLst", 256),
            ("p:sldIdLst/p:sldId{id=42}", 256),
            ("p:sldIdLst/p:sldId{id=256}", 257),
            ("p:sldIdLst/(p:sldId{id=256},p:sldId{id=712})", 713),
            ("p:sldIdLst/(p:sldId{id=280},p:sldId{id=257})", 281),
        ],
    )
    def it_knows_the_next_available_slide_id(self, sldIdLst_cxml: str, expected_value: int):
        sldIdLst = cast(CT_SlideIdList, element(sldIdLst_cxml))
        assert sldIdLst._next_id == expected_value

    @pytest.mark.parametrize(
        ("sldIdLst_cxml", "expected_value"),
        [
            ("p:sldIdLst/p:sldId{id=2147483646}", 2147483647),
            ("p:sldIdLst/p:sldId{id=2147483647}", 256),
            # -- 2147483648 is not a valid id but shouldn't stop us from finding a one that is --
            ("p:sldIdLst/p:sldId{id=2147483648}", 256),
            ("p:sldIdLst/(p:sldId{id=256},p:sldId{id=2147483647})", 257),
            ("p:sldIdLst/(p:sldId{id=256},p:sldId{id=2147483647},p:sldId{id=257})", 258),
            # -- 245 is also not a valid id but that shouldn't change the result either --
            ("p:sldIdLst/(p:sldId{id=245},p:sldId{id=2147483647},p:sldId{id=256})", 257),
        ],
    )
    def and_it_chooses_a_valid_slide_id_when_max_slide_id_is_used_for_a_slide(
        self, sldIdLst_cxml: str, expected_value: int
    ):
        sldIdLst = cast(CT_SlideIdList, element(sldIdLst_cxml))

        slide_id = sldIdLst._next_id

        assert 256 <= slide_id <= 2147483647
        assert slide_id == expected_value


class DescribeCT_SlideMasterIdList:
    """Unit-test suite for `pptx.oxml.presentation.CT_SlideMasterIdList` objects."""

    def it_can_add_an_entry_with_an_id(self):
        sldMasterIdLst = cast(CT_SlideMasterIdList, element("p:sldMasterIdLst"))

        sldMasterIdLst.add_sldMasterId("rId9", id=2147483660)

        (sldMasterId,) = sldMasterIdLst.sldMasterId_lst
        assert (sldMasterId.rId, sldMasterId.id) == ("rId9", 2147483660)

    def and_it_can_add_an_entry_without_one(self):
        sldMasterIdLst = cast(CT_SlideMasterIdList, element("p:sldMasterIdLst"))

        sldMasterIdLst.add_sldMasterId("rId9")

        (sldMasterId,) = sldMasterIdLst.sldMasterId_lst
        assert (sldMasterId.rId, sldMasterId.id) == ("rId9", None)
