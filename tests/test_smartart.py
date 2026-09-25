# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.smartart` module."""

from __future__ import annotations

import re

import pytest

from pptx.parts.diagram import DiagramDataPart
from pptx.parts.slide import SlidePart
from pptx.smartart import SmartArt, SmartArtNode

from .unitutil.cxml import element, xml
from .unitutil.mock import instance_mock

DATA_MODEL_EXT_URI = "http://schemas.microsoft.com/office/drawing/2008/diagram"

# -- doc "0" -> ("1" -> ("11", "12"), "2") --
TREE_CXML = (
    "dgm:dataModel/("
    "dgm:ptLst/("
    "dgm:pt{modelId=0,type=doc},"
    'dgm:pt{modelId=12}/dgm:t/(a:bodyPr,a:p/a:r/a:t"Budget"),'
    'dgm:pt{modelId=1}/dgm:t/(a:bodyPr,a:p/a:r/a:t"Plan"),'
    'dgm:pt{modelId=2}/dgm:t/(a:bodyPr,a:p/a:r/a:t"Build"),'
    'dgm:pt{modelId=11}/dgm:t/(a:bodyPr,a:p/a:r/a:t"Scope")),'
    "dgm:cxnLst/("
    "dgm:cxn{modelId=c4,srcId=1,destId=12,srcOrd=1},"
    "dgm:cxn{modelId=c2,srcId=0,destId=2,srcOrd=1},"
    "dgm:cxn{modelId=c3,srcId=1,destId=11,srcOrd=0},"
    "dgm:cxn{modelId=c1,srcId=0,destId=1,srcOrd=0}),"
    "dgm:extLst/a:ext{uri=%s}/dsp:dataModelExt{relId=rId6})" % DATA_MODEL_EXT_URI
)


def _without_nsdecls(xml_str: str) -> str:
    return re.sub(r' xmlns:\w+="[^"]*"', "", xml_str)


class DescribeSmartArt:
    """Unit-test suite for `pptx.smartart.SmartArt` objects."""

    def it_provides_access_to_its_top_level_nodes(self, smartart: SmartArt):
        nodes = smartart.nodes

        assert [(n.model_id, n.text, n.level) for n in nodes] == [
            ("1", "Plan", 0),
            ("2", "Build", 0),
        ]

    def it_has_no_nodes_when_there_is_no_doc_point(self, request):
        data_part_ = instance_mock(request, DiagramDataPart)
        data_part_.data_model = element("dgm:dataModel/dgm:ptLst/dgm:pt{modelId=1}")

        assert SmartArt(data_part_, None).nodes == ()

    def it_walks_all_its_nodes_depth_first(self, smartart: SmartArt):
        assert [(n.text, n.level) for n in smartart.iter_nodes()] == [
            ("Plan", 0),
            ("Scope", 1),
            ("Budget", 1),
            ("Build", 0),
        ]

    def it_drops_the_cached_drawing_when_invalidated(
        self, smartart: SmartArt, slide_part_: SlidePart
    ):
        smartart._invalidate_drawing()

        slide_part_.drop_rel.assert_called_once_with("rId6")
        assert smartart._data_model.drawing_rId is None

    def but_it_does_nothing_when_there_is_no_cached_drawing(self, request, slide_part_):
        data_part_ = instance_mock(request, DiagramDataPart)
        data_part_.data_model = element("dgm:dataModel/dgm:ptLst")

        SmartArt(data_part_, slide_part_)._invalidate_drawing()

        slide_part_.drop_rel.assert_not_called()

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def slide_part_(self, request):
        return instance_mock(request, SlidePart)

    @pytest.fixture
    def smartart(self, request, slide_part_) -> SmartArt:
        data_part_ = instance_mock(request, DiagramDataPart)
        data_part_.data_model = element(TREE_CXML)
        return SmartArt(data_part_, slide_part_)


class DescribeSmartArtNode:
    """Unit-test suite for `pptx.smartart.SmartArtNode` objects."""

    def it_provides_access_to_its_children(self, smartart: SmartArt):
        plan, build = smartart.nodes

        assert [(n.text, n.level) for n in plan.children] == [("Scope", 1), ("Budget", 1)]
        assert build.children == ()

    @pytest.mark.parametrize(
        ("pt_cxml", "expected_value"),
        [
            ("dgm:pt{modelId=1}", ""),
            ('dgm:pt{modelId=1}/dgm:t/(a:bodyPr,a:p/a:r/a:t"foo")', "foo"),
            ('dgm:pt{modelId=1}/dgm:t/(a:bodyPr,a:p/a:r/a:t"foo",a:p/a:r/a:t"bar")', "foo\nbar"),
        ],
    )
    def it_knows_its_text(self, pt_cxml: str, expected_value: str):
        assert SmartArtNode(element(pt_cxml), None, 0).text == expected_value

    @pytest.mark.parametrize(
        ("pt_cxml", "text", "expected_cxml"),
        [
            (
                "dgm:pt{modelId=1}/(dgm:prSet{phldr=1},dgm:spPr)",
                "foo",
                'dgm:pt{modelId=1}/(dgm:prSet,dgm:spPr,dgm:t/(a:bodyPr,a:lstStyle,a:p/a:r/a:t"foo"))',
            ),
            (
                'dgm:pt{modelId=1}/dgm:t/(a:bodyPr,a:p/a:r/a:t"foo",a:p/a:r/a:t"bar")',
                "baz\nqux",
                'dgm:pt{modelId=1}/dgm:t/(a:bodyPr,a:p/a:r/a:t"baz",a:p/a:r/a:t"qux")',
            ),
            (
                "dgm:pt{modelId=1}/dgm:prSet{phldr=1}",
                "",
                "dgm:pt{modelId=1}/(dgm:prSet{phldr=1},dgm:t/(a:bodyPr,a:lstStyle,a:p))",
            ),
        ],
    )
    def it_can_change_its_text(self, request, pt_cxml: str, text: str, expected_cxml: str):
        smartart_ = instance_mock(request, SmartArt)
        pt = element(pt_cxml)
        node = SmartArtNode(pt, smartart_, 0)

        node.text = text

        assert _without_nsdecls(pt.xml) == _without_nsdecls(xml(expected_cxml))
        smartart_._invalidate_drawing.assert_called_once_with()

    def it_is_equal_to_another_node_for_the_same_point(self):
        pt, other_pt = element("dgm:pt{modelId=1}"), element("dgm:pt{modelId=1}")

        assert SmartArtNode(pt, None, 0) == SmartArtNode(pt, None, 0)
        assert SmartArtNode(pt, None, 0) != SmartArtNode(other_pt, None, 0)
        assert SmartArtNode(pt, None, 0) != "foo"
        assert hash(SmartArtNode(pt, None, 0)) == hash(SmartArtNode(pt, None, 0))

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def smartart(self, request) -> SmartArt:
        data_part_ = instance_mock(request, DiagramDataPart)
        data_part_.data_model = element(TREE_CXML)
        return SmartArt(data_part_, instance_mock(request, SlidePart))
