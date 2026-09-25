# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.oxml.presprops` module."""

from __future__ import annotations

from typing import cast

import pytest

from pptx.oxml.presprops import CT_PresentationProperties, CT_ShowProperties

from ..unitutil.cxml import element, xml


class DescribeCT_PresentationProperties:
    """Unit-test suite for `pptx.oxml.presprops.CT_PresentationProperties`."""

    def it_has_no_showPr_by_default(self):
        presentationPr = cast(CT_PresentationProperties, element("p:presentationPr"))
        assert presentationPr.showPr is None

    def it_can_get_or_add_a_showPr_element(self):
        presentationPr = cast(CT_PresentationProperties, element("p:presentationPr"))
        showPr = presentationPr.get_or_add_showPr()
        assert presentationPr.xml == xml("p:presentationPr/p:showPr")
        assert presentationPr.get_or_add_showPr() is showPr

    def it_inserts_showPr_before_an_existing_extLst(self):
        presentationPr = cast(CT_PresentationProperties, element("p:presentationPr/p:extLst"))
        presentationPr.get_or_add_showPr()
        assert presentationPr.xml == xml("p:presentationPr/(p:showPr,p:extLst)")


class DescribeCT_ShowProperties:
    """Unit-test suite for `pptx.oxml.presprops.CT_ShowProperties`."""

    @pytest.mark.parametrize(
        ("showPr_cxml", "expected_value"),
        [
            ("p:showPr", False),
            ("p:showPr{loop=1}", True),
            ("p:showPr{loop=0}", False),
        ],
    )
    def it_knows_the_loop_setting(self, showPr_cxml: str, expected_value: bool):
        showPr = cast(CT_ShowProperties, element(showPr_cxml))
        assert showPr.loop is expected_value

    def it_has_showAnimation_and_useTimings_true_by_default(self):
        showPr = cast(CT_ShowProperties, element("p:showPr"))
        assert showPr.showAnimation is True
        assert showPr.useTimings is True

    def it_can_change_to_each_show_type(self):
        showPr = cast(CT_ShowProperties, element("p:showPr"))

        showPr.get_or_change_to_browse()
        assert showPr.showTypeChoice.tag.endswith("}browse")

        showPr.get_or_change_to_kiosk()
        assert showPr.showTypeChoice.tag.endswith("}kiosk")
        # -- changing choice removes the prior one --
        assert showPr.xml.count("browse") == 0

        showPr.get_or_change_to_present()
        assert showPr.showTypeChoice.tag.endswith("}present")

    def it_can_change_to_each_slide_range_choice(self):
        showPr = cast(CT_ShowProperties, element("p:showPr"))

        sldRg = showPr.get_or_change_to_sldRg()
        sldRg.st, sldRg.end = 2, 5
        assert showPr.slideRangeChoice.tag.endswith("}sldRg")

        custShow = showPr.get_or_change_to_custShow()
        custShow.id = 3
        assert showPr.slideRangeChoice.tag.endswith("}custShow")
        assert showPr.xml.count("sldRg") == 0

        showPr.get_or_change_to_sldAll()
        assert showPr.slideRangeChoice.tag.endswith("}sldAll")

    def it_can_get_or_add_a_penClr_element(self):
        showPr = cast(CT_ShowProperties, element("p:showPr"))
        penClr = showPr.get_or_add_penClr()
        srgbClr = penClr.get_or_change_to_srgbClr()
        srgbClr.val = "FF0000"

        assert showPr.penClr is penClr
        assert penClr.eg_colorChoice.val == "FF0000"


class DescribeShowPrCustomShow:
    """`p:showPr/p:custShow` shares its tag with the `p:custShowLst/p:custShow` definition."""

    def it_reads_the_custom_show_id_through_the_shared_element_class(self):
        from pptx.oxml.presentation import CT_CustomShow

        showPr = element("p:showPr/p:custShow{id=3}")
        definition = element("p:custShowLst/p:custShow{name=A,id=3}")[0]

        assert isinstance(showPr.custShow, CT_CustomShow)
        assert showPr.custShow.id == 3
        assert definition.name == "A"
