"""Unit-test suite for `pptx.theme` module."""

from __future__ import annotations

import os

import pytest

import pptx
from pptx.dml.fill import FillFormat
from pptx.dml.line import LineFormat
from pptx.enum.dml import MSO_FILL, MSO_THEME_COLOR
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.package import Package
from pptx.theme import EffectScheme, EffectStyle, Theme
from pptx.util import Pt

from .unitutil.cxml import element

_DEFAULT_TEMPLATE = os.path.join(os.path.dirname(pptx.__file__), "templates", "default.pptx")


class DescribeTheme:
    """Unit-test suite for `pptx.theme.Theme` objects."""

    def it_provides_access_to_effect_scheme(self):
        theme_xml = (
            '<a:theme %s><a:themeElements>'
            '<a:clrScheme name="x"><a:dk1><a:srgbClr val="000000"/></a:dk1>'
            '<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
            '<a:dk2><a:srgbClr val="000000"/></a:dk2>'
            '<a:lt2><a:srgbClr val="FFFFFF"/></a:lt2>'
            '<a:accent1><a:srgbClr val="000000"/></a:accent1>'
            '<a:accent2><a:srgbClr val="000000"/></a:accent2>'
            '<a:accent3><a:srgbClr val="000000"/></a:accent3>'
            '<a:accent4><a:srgbClr val="000000"/></a:accent4>'
            '<a:accent5><a:srgbClr val="000000"/></a:accent5>'
            '<a:accent6><a:srgbClr val="000000"/></a:accent6>'
            '<a:hlink><a:srgbClr val="000000"/></a:hlink>'
            '<a:folHlink><a:srgbClr val="000000"/></a:folHlink>'
            '</a:clrScheme>'
            '<a:fontScheme name="x"><a:majorFont><a:latin typeface="Arial"/>'
            '<a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
            '<a:minorFont><a:latin typeface="Arial"/>'
            '<a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
            '<a:fmtScheme name="Office">'
            '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
            '<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/>'
            '</a:solidFill></a:ln><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/>'
            '</a:solidFill></a:ln><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/>'
            '</a:solidFill></a:ln></a:lnStyleLst>'
            '<a:effectStyleLst>'
            '<a:effectStyle><a:effectLst/></a:effectStyle>'
            '<a:effectStyle><a:effectLst/></a:effectStyle>'
            '<a:effectStyle><a:effectLst/><a:scene3d/><a:sp3d/></a:effectStyle>'
            '</a:effectStyleLst>'
            '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>'
            '</a:fmtScheme>'
            '</a:themeElements></a:theme>' % nsdecls("a")
        )
        theme_elm = parse_xml(theme_xml)
        theme = Theme(theme_elm)

        effect_scheme = theme.effect_scheme
        assert effect_scheme is not None
        assert isinstance(effect_scheme, EffectScheme)
        assert effect_scheme.name == "Office"

    def it_returns_None_when_no_format_scheme(self):
        theme_xml = (
            '<a:theme %s><a:themeElements>'
            '<a:clrScheme name="x"><a:dk1><a:srgbClr val="000000"/></a:dk1>'
            '<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
            '<a:dk2><a:srgbClr val="000000"/></a:dk2>'
            '<a:lt2><a:srgbClr val="FFFFFF"/></a:lt2>'
            '<a:accent1><a:srgbClr val="000000"/></a:accent1>'
            '<a:accent2><a:srgbClr val="000000"/></a:accent2>'
            '<a:accent3><a:srgbClr val="000000"/></a:accent3>'
            '<a:accent4><a:srgbClr val="000000"/></a:accent4>'
            '<a:accent5><a:srgbClr val="000000"/></a:accent5>'
            '<a:accent6><a:srgbClr val="000000"/></a:accent6>'
            '<a:hlink><a:srgbClr val="000000"/></a:hlink>'
            '<a:folHlink><a:srgbClr val="000000"/></a:folHlink>'
            '</a:clrScheme>'
            '<a:fontScheme name="x"><a:majorFont><a:latin typeface="Arial"/>'
            '<a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
            '<a:minorFont><a:latin typeface="Arial"/>'
            '<a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
            '</a:themeElements></a:theme>' % nsdecls("a")
        )
        theme_elm = parse_xml(theme_xml)
        theme = Theme(theme_elm)
        assert theme.effect_scheme is None


class DescribeEffectScheme:
    """Unit-test suite for `pptx.theme.EffectScheme` objects."""

    @pytest.fixture
    def effect_scheme(self):
        xml = (
            '<a:fmtScheme %s name="TestScheme">'
            '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
            '<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/>'
            '</a:solidFill></a:ln><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/>'
            '</a:solidFill></a:ln><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/>'
            '</a:solidFill></a:ln></a:lnStyleLst>'
            '<a:effectStyleLst>'
            '<a:effectStyle><a:effectLst/></a:effectStyle>'
            '<a:effectStyle><a:effectLst/></a:effectStyle>'
            '<a:effectStyle><a:effectLst/><a:scene3d/><a:sp3d/></a:effectStyle>'
            '</a:effectStyleLst>'
            '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>'
            '</a:fmtScheme>' % nsdecls("a")
        )
        fmtScheme = parse_xml(xml)
        return EffectScheme(fmtScheme.effectStyleLst, fmtScheme.name)

    def it_has_the_format_scheme_name(self, effect_scheme):
        assert effect_scheme.name == "TestScheme"

    def it_has_three_styles(self, effect_scheme):
        assert len(effect_scheme) == 3

    def it_supports_iteration(self, effect_scheme):
        styles = list(effect_scheme)
        assert len(styles) == 3
        assert all(isinstance(s, EffectStyle) for s in styles)
        assert [s.name for s in styles] == ["subtle", "moderate", "intense"]

    def it_supports_indexed_access(self, effect_scheme):
        subtle = effect_scheme[0]
        assert subtle.name == "subtle"
        intense = effect_scheme[2]
        assert intense.name == "intense"

    def it_provides_named_accessors(self, effect_scheme):
        assert effect_scheme.subtle.name == "subtle"
        assert effect_scheme.moderate.name == "moderate"
        assert effect_scheme.intense.name == "intense"


class DescribeEffectStyle:
    """Unit-test suite for `pptx.theme.EffectStyle` objects."""

    @pytest.fixture
    def effect_scheme(self):
        xml = (
            '<a:fmtScheme %s name="Office">'
            '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
            '<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/>'
            '</a:solidFill></a:ln><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/>'
            '</a:solidFill></a:ln><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/>'
            '</a:solidFill></a:ln></a:lnStyleLst>'
            '<a:effectStyleLst>'
            '<a:effectStyle><a:effectLst/></a:effectStyle>'
            '<a:effectStyle><a:effectLst/></a:effectStyle>'
            '<a:effectStyle><a:effectLst/><a:scene3d/><a:sp3d/></a:effectStyle>'
            '</a:effectStyleLst>'
            '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>'
            '</a:fmtScheme>' % nsdecls("a")
        )
        fmtScheme = parse_xml(xml)
        return EffectScheme(fmtScheme.effectStyleLst, fmtScheme.name)

    def it_knows_if_it_has_an_effect_list(self, effect_scheme):
        assert effect_scheme.subtle.has_effect_list is True

    def it_knows_if_it_has_3d_scene(self, effect_scheme):
        assert effect_scheme.subtle.has_3d_scene is False
        assert effect_scheme.intense.has_3d_scene is True

    def it_knows_if_it_has_3d_shape(self, effect_scheme):
        assert effect_scheme.subtle.has_3d_shape is False
        assert effect_scheme.intense.has_3d_shape is True


_FMT_SCHEME_THEME_XML = (
    "<a:theme %s><a:themeElements>"
    '<a:fmtScheme name="Office">'
    '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"/></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs>'
    '</a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>'
    '<a:pattFill prst="pct5"/></a:fillStyleLst>'
    '<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '</a:ln><a:ln w="25400"><a:noFill/></a:ln><a:ln w="38100"/></a:lnStyleLst>'
    "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>"
    '<a:bgFillStyleLst><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>'
    "<a:noFill/></a:bgFillStyleLst>"
    "</a:fmtScheme>"
    "</a:themeElements></a:theme>" % nsdecls("a")
)


class DescribeTheme_FormatSchemeStyles:
    """Unit-test suite for the fill- and line-style lists of `pptx.theme.Theme`."""

    def it_provides_the_fill_styles_in_order(self):
        theme = Theme(parse_xml(_FMT_SCHEME_THEME_XML))

        fill_styles = theme.fill_styles

        assert all(isinstance(fill, FillFormat) for fill in fill_styles)
        assert [fill.type for fill in fill_styles] == [
            MSO_FILL.SOLID,
            MSO_FILL.GRADIENT,
            MSO_FILL.PATTERNED,
        ]
        assert fill_styles[1].gradient_angle == 90.0

    def it_reports_the_placeholder_color_used_by_a_fill_style(self):
        theme = Theme(parse_xml(_FMT_SCHEME_THEME_XML))
        assert theme.fill_styles[0].fore_color.theme_color == MSO_THEME_COLOR.PLACEHOLDER

    def it_provides_the_line_styles_in_order(self):
        theme = Theme(parse_xml(_FMT_SCHEME_THEME_XML))

        line_styles = theme.line_styles

        assert all(isinstance(line, LineFormat) for line in line_styles)
        assert [line.width for line in line_styles] == [Pt(0.75), Pt(2), Pt(3)]
        assert [line.fill.type for line in line_styles] == [
            MSO_FILL.SOLID,
            MSO_FILL.BACKGROUND,
            None,
        ]

    def it_provides_the_background_fill_styles_in_order(self):
        theme = Theme(parse_xml(_FMT_SCHEME_THEME_XML))

        bg_styles = theme.background_fill_styles

        assert [fill.type for fill in bg_styles] == [MSO_FILL.SOLID, MSO_FILL.BACKGROUND]
        assert str(bg_styles[0].fore_color.rgb) == "FF0000"

    @pytest.mark.parametrize(
        "theme_cxml",
        [
            "a:theme",
            "a:theme/a:themeElements",
            "a:theme/a:themeElements/a:fmtScheme",
        ],
    )
    def it_returns_empty_lists_when_the_theme_has_none(self, theme_cxml):
        theme = Theme(element(theme_cxml))

        assert theme.fill_styles == ()
        assert theme.line_styles == ()
        assert theme.background_fill_styles == ()

    def it_returns_the_three_fill_styles_of_the_default_theme(self):
        theme_part = Package.open(_DEFAULT_TEMPLATE).presentation_part.presentation.slide_master.part.theme_part
        theme = Theme(theme_part._element)

        assert len(theme.fill_styles) == 3
        assert len(theme.line_styles) == 3
        assert len(theme.background_fill_styles) == 3
        assert [fill.type for fill in theme.fill_styles] == [
            MSO_FILL.SOLID,
            MSO_FILL.GRADIENT,
            MSO_FILL.GRADIENT,
        ]
