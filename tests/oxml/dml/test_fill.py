"""Unit-test suite for `pptx.oxml.dml.fill` module."""

from __future__ import annotations

import pytest

from pptx.enum.dml import MSO_RECT_ALIGNMENT
from pptx.oxml.dml.fill import (
    CT_BlipFillProperties,
    CT_GradientFillProperties,
    CT_StretchInfoProperties,
    CT_TileInfoProperties,
)
from pptx.util import Emu

from ...unitutil.cxml import element, xml


class Describe_CT_BlipFillProperties:
    """Unit-test suite for `pptx.oxml.dml.fill.CT_BlipFillProperties`."""

    def it_provides_access_to_its_tile_and_stretch_choice_group(self):
        blipFill = element("a:blipFill")
        assert isinstance(blipFill, CT_BlipFillProperties)
        assert blipFill.tile is None
        assert blipFill.stretch is None

        stretch = blipFill.get_or_change_to_stretch()
        assert isinstance(stretch, CT_StretchInfoProperties)
        assert blipFill.stretch is stretch
        assert blipFill.tile is None

    def it_replaces_stretch_with_tile_and_vice_versa(self):
        blipFill = element("a:blipFill/a:stretch/a:fillRect")

        tile = blipFill.get_or_change_to_tile()

        assert isinstance(tile, CT_TileInfoProperties)
        assert blipFill.tile is tile
        assert blipFill.stretch is None
        assert blipFill.xml == xml("a:blipFill/a:tile")


class Describe_CT_StretchInfoProperties:
    """Unit-test suite for `pptx.oxml.dml.fill.CT_StretchInfoProperties`."""

    def it_can_get_or_add_a_fillRect_child(self):
        stretch = element("a:stretch")
        fillRect = stretch.get_or_add_fillRect()
        assert stretch.xml == xml("a:stretch/a:fillRect")
        assert stretch.fillRect is fillRect


class Describe_CT_TileInfoProperties:
    """Unit-test suite for `pptx.oxml.dml.fill.CT_TileInfoProperties`."""

    def it_provides_defaults_for_its_attributes(self):
        tile = element("a:tile")
        assert tile.tx == Emu(0)
        assert tile.ty == Emu(0)
        assert tile.sx == 1.0
        assert tile.sy == 1.0
        assert tile.flip == "none"
        assert tile.algn == MSO_RECT_ALIGNMENT.TOP_LEFT

    def it_can_change_its_attributes(self):
        tile = element("a:tile")

        tile.tx = Emu(1000)
        tile.ty = Emu(2000)
        tile.sx = 0.5
        tile.sy = 0.75
        tile.flip = "xy"
        tile.algn = MSO_RECT_ALIGNMENT.CENTER

        assert tile.xml == xml(
            'a:tile{tx=1000,ty=2000,sx=50000,sy=75000,flip=xy,algn=ctr}'
        )


class Describe_CT_GradientFillProperties:
    """Unit-test suite for `pptx.oxml.dml.fill.CT_GradientFillProperties`."""

    @pytest.mark.parametrize(
        ("gradFill_cxml", "expected_value"),
        [
            ("a:gradFill", None),
            ("a:gradFill/a:lin", None),
            ("a:gradFill/a:path{path=circle}", "circle"),
            ("a:gradFill/a:path{path=rect}", "rect"),
            ("a:gradFill/a:path{path=shape}", "shape"),
        ],
    )
    def it_knows_the_path_val(self, gradFill_cxml, expected_value):
        gradFill = element(gradFill_cxml)
        assert isinstance(gradFill, CT_GradientFillProperties)
        assert gradFill.path_val == expected_value

    def it_can_change_the_path_val(self):
        gradFill = element("a:gradFill")
        gradFill.path_val = "circle"
        assert gradFill.xml == xml('a:gradFill/a:path{path=circle}')

    def it_provides_access_to_fillToRect(self):
        gradFill = element("a:gradFill")
        assert gradFill.fillToRect is None

        fillToRect = gradFill.get_or_add_fillToRect()
        fillToRect.l, fillToRect.t, fillToRect.r, fillToRect.b = 0.1, 0.2, 0.3, 0.4

        assert gradFill.xml == xml(
            'a:gradFill/a:path/a:fillToRect{l=10000,t=20000,r=30000,b=40000}'
        )
        assert gradFill.fillToRect is fillToRect

    def it_can_remove_fillToRect(self):
        gradFill = element("a:gradFill/a:path{path=circle}/a:fillToRect{l=10000}")
        gradFill._remove_fillToRect()
        assert gradFill.xml == xml("a:gradFill/a:path{path=circle}")
        # -- no-op when a:path is absent --
        gradFill2 = element("a:gradFill")
        gradFill2._remove_fillToRect()
        assert gradFill2.xml == xml("a:gradFill")
