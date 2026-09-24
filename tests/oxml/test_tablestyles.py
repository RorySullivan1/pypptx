# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.oxml.tablestyles` module."""

from __future__ import annotations

from typing import cast

from pptx.oxml.tablestyles import CT_TableStyleList

from ..unitutil.cxml import element


class DescribeCT_TableStyleList:
    """Unit-test suite for `pptx.oxml.tablestyles.CT_TableStyleList`."""

    def it_provides_access_to_its_def_attribute(self):
        tblStyleLst = cast(
            CT_TableStyleList, element("a:tblStyleLst{def=GUID-0000-DEFAULT}")
        )
        assert tblStyleLst.def_ == "GUID-0000-DEFAULT"

    def it_defaults_def_to_None(self):
        tblStyleLst = cast(CT_TableStyleList, element("a:tblStyleLst"))
        assert tblStyleLst.def_ is None

    def it_provides_access_to_its_tblStyle_children(self):
        tblStyleLst = cast(
            CT_TableStyleList,
            element(
                "a:tblStyleLst/("
                "a:tblStyle{styleId=GUID-1,styleName=Light},"
                "a:tblStyle{styleId=GUID-2,styleName=Dark}"
                ")"
            ),
        )
        styles = tblStyleLst.tblStyle_lst
        assert [(s.styleId, s.styleName) for s in styles] == [
            ("GUID-1", "Light"),
            ("GUID-2", "Dark"),
        ]
