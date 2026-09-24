"""Unit-test suite for `pptx.oxml.theme` module."""

from __future__ import annotations

import pytest

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import PartFactory, XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.theme import (
    CT_BaseStylesOverride,
    CT_ColorScheme,
    CT_FillStyleList,
    CT_LineStyleList,
    CT_OfficeStyleSheet,
)

from ..unitutil.file import snippet_text


class DescribeCT_OfficeStyleSheet:
    def it_can_create_a_default_theme_element(self, new_fixture):
        expected_xml = new_fixture
        theme = CT_OfficeStyleSheet.new_default()
        assert theme.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def new_fixture(self):
        expected_xml = snippet_text("default-theme")
        return expected_xml


class DescribeCT_BaseStylesOverride:
    """Unit-test suite for `a:themeOverride`, the root of a theme-override part."""

    def it_provides_access_to_its_schemes(self):
        themeOverride = parse_xml(
            '<a:themeOverride %s><a:clrScheme name="x"/></a:themeOverride>' % nsdecls("a")
        )

        assert isinstance(themeOverride, CT_BaseStylesOverride)
        assert isinstance(themeOverride.clrScheme, CT_ColorScheme)
        assert themeOverride.fontScheme is None
        assert themeOverride.fmtScheme is None

    def it_loads_theme_override_parts_as_xml_parts(self):
        blob = ('<a:themeOverride %s><a:clrScheme name="x"/></a:themeOverride>' % nsdecls("a"))

        part = PartFactory(
            PackURI("/ppt/theme/themeOverride1.xml"),
            CT.OFC_THEME_OVERRIDE,
            None,  # pyright: ignore[reportArgumentType]
            blob.encode("utf-8"),
        )

        assert type(part) is XmlPart
        assert isinstance(part._element, CT_BaseStylesOverride)


class DescribeCT_StyleMatrix:
    """Unit-test suite for the fill and line style lists of `a:fmtScheme`."""

    def it_provides_access_to_its_style_lists(self):
        fmtScheme = parse_xml(
            "<a:fmtScheme %s><a:fillStyleLst><a:solidFill/><a:noFill/></a:fillStyleLst>"
            "<a:lnStyleLst><a:ln/><a:ln/></a:lnStyleLst><a:effectStyleLst/>"
            "<a:bgFillStyleLst><a:gradFill/></a:bgFillStyleLst></a:fmtScheme>" % nsdecls("a")
        )

        assert isinstance(fmtScheme.fillStyleLst, CT_FillStyleList)
        assert [e.tag.split("}")[1] for e in fmtScheme.fillStyleLst.fill_elms] == [
            "solidFill",
            "noFill",
        ]
        assert isinstance(fmtScheme.lnStyleLst, CT_LineStyleList)
        assert len(fmtScheme.lnStyleLst.ln_lst) == 2
        assert isinstance(fmtScheme.bgFillStyleLst, CT_FillStyleList)
        assert len(fmtScheme.bgFillStyleLst.fill_elms) == 1
