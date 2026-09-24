"""Unit-test suite for `pptx.text.styles` module."""

from __future__ import annotations

import pytest

from pptx.enum.text import PP_ALIGN
from pptx.text.styles import MasterTextStyles, TextLevelStyle, TextListStyle
from pptx.text.text import Font
from pptx.util import Emu, Pt

from ..unitutil.cxml import element, xml


class DescribeCT_TextListStyle:
    """Unit-test suite for `pptx.oxml.text.CT_TextListStyle` objects."""

    def it_finds_the_paragraph_properties_for_a_level(self):
        lstStyle = element("a:lstStyle/(a:defPPr,a:lvl1pPr{marL=1},a:lvl3pPr{marL=3})")

        assert lstStyle.lvl_pPr(0).marL == 1
        assert lstStyle.lvl_pPr(1) is None
        assert lstStyle.lvl_pPr(2).marL == 3

    def it_adds_level_properties_in_schema_order(self):
        lstStyle = element("a:lstStyle/(a:defPPr,a:lvl1pPr,a:lvl9pPr)")

        lstStyle.get_or_add_lvl_pPr(4)

        assert lstStyle.xml == xml("a:lstStyle/(a:defPPr,a:lvl1pPr,a:lvl5pPr,a:lvl9pPr)")

    @pytest.mark.parametrize("level", [-1, 9, "1", 1.0, True])
    def it_rejects_a_level_outside_0_to_8(self, level):
        lstStyle = element("a:lstStyle")
        with pytest.raises(IndexError):
            lstStyle.lvl_pPr(level)


class DescribeMasterTextStyles:
    """Unit-test suite for `pptx.text.styles.MasterTextStyles` objects."""

    @pytest.mark.parametrize(
        ("prop_name", "tagname"),
        [("title", "p:titleStyle"), ("body", "p:bodyStyle"), ("other", "p:otherStyle")],
    )
    def it_provides_access_to_each_text_style(self, prop_name, tagname):
        sldMaster = element(
            "p:sldMaster/(p:cSld,p:txStyles/(p:titleStyle/a:lvl1pPr{marL=1},"
            "p:bodyStyle/a:lvl1pPr{marL=2},p:otherStyle/a:lvl1pPr{marL=3}))"
        )
        expected = {"title": 1, "body": 2, "other": 3}[prop_name]

        style = getattr(MasterTextStyles(sldMaster), prop_name)

        assert isinstance(style, TextListStyle)
        assert style[0].margin_left == expected

    def it_reads_None_without_adding_txStyles(self):
        sldMaster = element("p:sldMaster/p:cSld")

        assert MasterTextStyles(sldMaster).body[0].margin_left is None
        assert sldMaster.xml == xml("p:sldMaster/p:cSld")

    def it_adds_txStyles_in_schema_order_when_written(self):
        sldMaster = element("p:sldMaster/(p:cSld,p:clrMap,p:extLst)")

        MasterTextStyles(sldMaster).body[0].font.size = Pt(20)

        assert [child.tag.split("}")[1] for child in sldMaster] == [
            "cSld",
            "clrMap",
            "txStyles",
            "extLst",
        ]
        assert sldMaster.xpath("p:txStyles/*/*/*/@sz") == ["2000"]
        assert sldMaster.xpath("p:txStyles/p:bodyStyle/a:lvl1pPr/a:defRPr/@sz") == ["2000"]


class DescribeTextListStyle:
    """Unit-test suite for `pptx.text.styles.TextListStyle` objects."""

    def it_has_nine_levels(self):
        lstStyle = element("a:lstStyle")
        style = TextListStyle(lambda: lstStyle, lambda: lstStyle)

        assert len(style) == 9
        assert all(isinstance(level, TextLevelStyle) for level in style)

    def it_maps_level_indexes_to_lvlNpPr_elements(self):
        lstStyle = element("a:lstStyle")
        style = TextListStyle(lambda: lstStyle, lambda: lstStyle)

        style[0].margin_left = Emu(10)
        style[8].margin_left = Emu(90)

        assert lstStyle.xml == xml("a:lstStyle/(a:lvl1pPr{marL=10},a:lvl9pPr{marL=90})")

    def it_provides_access_to_the_default_paragraph_properties(self):
        lstStyle = element("a:lstStyle/(a:defPPr{marL=5},a:lvl1pPr)")
        style = TextListStyle(lambda: lstStyle, lambda: lstStyle)

        assert style.default.margin_left == 5

    @pytest.mark.parametrize("level", [-1, 9, "0"])
    def it_raises_on_a_level_out_of_range(self, level):
        lstStyle = element("a:lstStyle")
        with pytest.raises(IndexError):
            TextListStyle(lambda: lstStyle, lambda: lstStyle)[level]


class DescribeTextLevelStyle:
    """Unit-test suite for `pptx.text.styles.TextLevelStyle` objects."""

    @pytest.mark.parametrize(
        ("prop_name", "pPr_cxml", "expected"),
        [
            ("alignment", "a:lvl1pPr{algn=ctr}", PP_ALIGN.CENTER),
            ("indent", "a:lvl1pPr{indent=-342900}", -342900),
            ("margin_left", "a:lvl1pPr{marL=342900}", 342900),
            ("line_spacing", "a:lvl1pPr/a:lnSpc/a:spcPct{val=90000}", 0.9),
            ("space_after", "a:lvl1pPr/a:spcAft/a:spcPts{val=600}", Pt(6)),
            ("space_before", "a:lvl1pPr/a:spcBef/a:spcPts{val=1200}", Pt(12)),
            ("alignment", "a:lvl1pPr", None),
            ("space_before", "a:lvl1pPr/a:spcBef/a:spcPct{val=20000}", None),
        ],
    )
    def it_reads_its_paragraph_properties(self, prop_name, pPr_cxml, expected):
        pPr = element(pPr_cxml)
        assert getattr(TextLevelStyle(lambda: pPr, lambda: pPr), prop_name) == expected

    @pytest.mark.parametrize(
        "prop_name",
        ["alignment", "indent", "margin_left", "line_spacing", "space_after", "space_before"],
    )
    def it_reads_None_when_the_level_is_not_defined(self, prop_name):
        def fail():
            raise AssertionError("read must not add XML")

        assert getattr(TextLevelStyle(lambda: None, fail), prop_name) is None

    @pytest.mark.parametrize(
        ("prop_name", "value", "expected_cxml"),
        [
            ("alignment", PP_ALIGN.RIGHT, "a:lvl1pPr{algn=r}"),
            ("indent", Emu(-100), "a:lvl1pPr{indent=-100}"),
            ("margin_left", Emu(200), "a:lvl1pPr{marL=200}"),
            ("line_spacing", Pt(20), "a:lvl1pPr/a:lnSpc/a:spcPts{val=2000}"),
            ("space_after", Pt(3), "a:lvl1pPr/a:spcAft/a:spcPts{val=300}"),
            ("space_before", Pt(4), "a:lvl1pPr/a:spcBef/a:spcPts{val=400}"),
        ],
    )
    def it_writes_its_paragraph_properties(self, prop_name, value, expected_cxml):
        pPr = element("a:lvl1pPr")

        setattr(TextLevelStyle(lambda: None, lambda: pPr), prop_name, value)

        assert pPr.xml == xml(expected_cxml)

    def it_provides_the_default_run_font(self):
        pPr = element("a:lvl1pPr/a:defRPr{sz=3200}")

        font = TextLevelStyle(lambda: pPr, lambda: pPr).font

        assert isinstance(font, Font)
        assert font.size == Pt(32)
        font.bold = True
        assert pPr.xml == xml("a:lvl1pPr/a:defRPr{sz=3200,b=1}")
