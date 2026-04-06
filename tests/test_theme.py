"""Unit-test suite for `pptx.theme` module."""

from __future__ import annotations

import pytest

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.theme import EffectScheme, EffectStyle, Theme


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
