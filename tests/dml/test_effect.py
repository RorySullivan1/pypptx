"""Unit-test suite for `pptx.dml.effect` module."""

from __future__ import annotations

import pytest

from pptx.dml.effect import ShadowFormat
from pptx.enum.dml import MSO_RECT_ALIGNMENT
from pptx.util import Emu, Pt

from ..unitutil.cxml import element, xml


class DescribeShadowFormat:
    @pytest.mark.parametrize(
        ("spPr_cxml", "expected_value"),
        [
            ("p:spPr", None),
            ("p:spPr/a:effectLst", None),
            ("p:spPr/a:effectLst/a:outerShdw{blurRad=50800}", Emu(50800)),
        ],
    )
    def it_knows_its_blur_radius(self, spPr_cxml: str, expected_value):
        shadow = ShadowFormat(element(spPr_cxml))
        assert shadow.blur_radius == expected_value

    def it_can_change_its_blur_radius(self):
        spPr = element("p:spPr{a:b=c}")
        shadow = ShadowFormat(spPr)
        shadow.blur_radius = 50800
        assert spPr.xml == xml("p:spPr{a:b=c}/a:effectLst/a:outerShdw{blurRad=50800}")

    @pytest.mark.parametrize(
        ("spPr_cxml", "expected_value"),
        [
            ("p:spPr", None),
            ("p:spPr/a:effectLst/a:outerShdw{dist=38100}", Emu(38100)),
        ],
    )
    def it_knows_its_distance(self, spPr_cxml: str, expected_value):
        shadow = ShadowFormat(element(spPr_cxml))
        assert shadow.distance == expected_value

    def it_can_change_its_distance(self):
        spPr = element("p:spPr{a:b=c}")
        shadow = ShadowFormat(spPr)
        shadow.distance = 38100
        assert spPr.xml == xml("p:spPr{a:b=c}/a:effectLst/a:outerShdw{dist=38100}")

    @pytest.mark.parametrize(
        ("spPr_cxml", "expected_value"),
        [
            ("p:spPr", None),
            ("p:spPr/a:effectLst/a:outerShdw{dir=2700000}", 45.0),
        ],
    )
    def it_knows_its_direction(self, spPr_cxml: str, expected_value):
        shadow = ShadowFormat(element(spPr_cxml))
        assert shadow.direction == expected_value

    def it_can_change_its_direction(self):
        spPr = element("p:spPr{a:b=c}")
        shadow = ShadowFormat(spPr)
        shadow.direction = 45.0
        assert spPr.xml == xml("p:spPr{a:b=c}/a:effectLst/a:outerShdw{dir=2700000}")

    @pytest.mark.parametrize(
        ("spPr_cxml", "expected_value"),
        [
            ("p:spPr", None),
            ("p:spPr/a:effectLst/a:outerShdw{algn=bl}", MSO_RECT_ALIGNMENT.BOTTOM_LEFT),
        ],
    )
    def it_knows_its_alignment(self, spPr_cxml: str, expected_value):
        shadow = ShadowFormat(element(spPr_cxml))
        assert shadow.alignment == expected_value

    @pytest.mark.parametrize(
        ("spPr_cxml", "expected_value"),
        [
            ("p:spPr", None),
            ("p:spPr/a:effectLst/a:outerShdw", "outer"),
            ("p:spPr/a:effectLst/a:innerShdw", "inner"),
        ],
    )
    def it_knows_its_shadow_type(self, spPr_cxml: str, expected_value: str | None):
        shadow = ShadowFormat(element(spPr_cxml))
        assert shadow.shadow_type == expected_value

    def it_knows_whether_it_inherits(self, inherit_get_fixture):
        shadow, expected_value = inherit_get_fixture
        inherit = shadow.inherit
        assert inherit is expected_value

    def it_can_change_whether_it_inherits(self, inherit_set_fixture):
        shadow, value, expected_xml = inherit_set_fixture
        shadow.inherit = value
        assert shadow._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("p:spPr", True),
            ("p:spPr/a:effectLst", False),
            ("p:grpSpPr", True),
            ("p:grpSpPr/a:effectLst", False),
        ]
    )
    def inherit_get_fixture(self, request):
        cxml, expected_value = request.param
        shadow = ShadowFormat(element(cxml))
        return shadow, expected_value

    @pytest.fixture(
        params=[
            ("p:spPr{a:b=c}", False, "p:spPr{a:b=c}/a:effectLst"),
            ("p:grpSpPr{a:b=c}", False, "p:grpSpPr{a:b=c}/a:effectLst"),
            ("p:spPr{a:b=c}/a:effectLst", True, "p:spPr{a:b=c}"),
            ("p:grpSpPr{a:b=c}/a:effectLst", True, "p:grpSpPr{a:b=c}"),
            ("p:spPr", True, "p:spPr"),
            ("p:grpSpPr", True, "p:grpSpPr"),
            ("p:spPr/a:effectLst", False, "p:spPr/a:effectLst"),
            ("p:grpSpPr/a:effectLst", False, "p:grpSpPr/a:effectLst"),
        ]
    )
    def inherit_set_fixture(self, request):
        cxml, value, expected_cxml = request.param
        shadow = ShadowFormat(element(cxml))
        expected_value = xml(expected_cxml)
        return shadow, value, expected_value
