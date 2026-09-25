"""Unit-test suite for `pptx.oxml.diagram` module."""

from __future__ import annotations

import re

import pytest

from pptx.oxml.diagram import CT_Cxn, CT_DataModel, CT_DiagramRelIds, CT_Pt
from pptx.oxml.text import CT_TextBody

from ..unitutil.cxml import element, xml

DATA_MODEL_EXT_URI = "http://schemas.microsoft.com/office/drawing/2008/diagram"


def _without_nsdecls(xml_str: str) -> str:
    """`xml_str` with its namespace declarations removed; cxml declares only prefixes it uses."""
    return re.sub(r' xmlns:\w+="[^"]*"', "", xml_str)


class DescribeCT_DiagramRelIds:
    """Unit-test suite for `pptx.oxml.diagram.CT_DiagramRelIds` objects."""

    def it_knows_the_rIds_of_its_four_diagram_parts(self):
        relIds = element("dgm:relIds{r:dm=rId2,r:lo=rId3,r:qs=rId4,r:cs=rId5}")

        assert isinstance(relIds, CT_DiagramRelIds)
        assert (relIds.dm, relIds.lo, relIds.qs, relIds.cs) == ("rId2", "rId3", "rId4", "rId5")


class DescribeCT_DataModel:
    """Unit-test suite for `pptx.oxml.diagram.CT_DataModel` objects."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("dgm:dataModel/dgm:ptLst/(dgm:pt{modelId=1},dgm:pt{modelId=0,type=doc})", "0"),
            ("dgm:dataModel/dgm:ptLst/dgm:pt{modelId=1}", None),
            ("dgm:dataModel", None),
        ],
    )
    def it_knows_its_doc_point(self, cxml: str, expected_value: str | None):
        dataModel = element(cxml)
        assert isinstance(dataModel, CT_DataModel)

        doc_pt = dataModel.doc_pt

        assert (None if doc_pt is None else doc_pt.modelId) == expected_value

    def it_finds_the_ordered_content_children_of_a_point(self):
        dataModel = element(
            "dgm:dataModel/("
            "dgm:ptLst/("
            "dgm:pt{modelId=0,type=doc},dgm:pt{modelId=3},dgm:pt{modelId=1},"
            "dgm:pt{modelId=2,type=asst},dgm:pt{modelId=9,type=pres},"
            "dgm:pt{modelId=8,type=parTrans}),"
            "dgm:cxnLst/("
            "dgm:cxn{modelId=c3,srcId=0,destId=3,srcOrd=2},"
            "dgm:cxn{modelId=c1,type=parOf,srcId=0,destId=1,srcOrd=0},"
            "dgm:cxn{modelId=c2,srcId=0,destId=2,srcOrd=1},"
            "dgm:cxn{modelId=c9,type=presOf,srcId=0,destId=9,srcOrd=0},"
            "dgm:cxn{modelId=c8,srcId=0,destId=8,srcOrd=3},"
            "dgm:cxn{modelId=cx,srcId=0,destId=missing,srcOrd=4},"
            "dgm:cxn{modelId=c4,srcId=1,destId=3,srcOrd=0}))"
        )

        assert [pt.modelId for pt in dataModel.child_pts("0")] == ["1", "2", "3"]
        assert [pt.modelId for pt in dataModel.child_pts("1")] == ["3"]
        assert dataModel.child_pts("3") == []

    def it_has_no_points_or_connections_when_the_lists_are_absent(self):
        dataModel = element("dgm:dataModel")

        assert list(dataModel.iter_pts()) == []
        assert list(dataModel.iter_cxns()) == []
        assert dataModel.child_pts("0") == []

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            (
                "dgm:dataModel/dgm:extLst/a:ext{uri=%s}/dsp:dataModelExt{relId=rId6}"
                % DATA_MODEL_EXT_URI,
                "rId6",
            ),
            ("dgm:dataModel/dgm:extLst/a:ext{uri=foo}/dsp:dataModelExt{relId=rId6}", None),
            ("dgm:dataModel", None),
        ],
    )
    def it_knows_the_rId_of_its_cached_drawing(self, cxml: str, expected_value: str | None):
        assert element(cxml).drawing_rId == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_cxml"),
        [
            (
                "dgm:dataModel/(dgm:ptLst,dgm:extLst/a:ext{uri=%s}/dsp:dataModelExt{relId=rId6})"
                % DATA_MODEL_EXT_URI,
                "dgm:dataModel/dgm:ptLst",
            ),
            (
                "dgm:dataModel/dgm:extLst/(a:ext{uri=%s}/dsp:dataModelExt{relId=rId6},"
                "a:ext{uri=foo})" % DATA_MODEL_EXT_URI,
                "dgm:dataModel/dgm:extLst/a:ext{uri=foo}",
            ),
            ("dgm:dataModel/dgm:ptLst", "dgm:dataModel/dgm:ptLst"),
        ],
    )
    def it_can_remove_its_reference_to_the_cached_drawing(self, cxml: str, expected_cxml: str):
        dataModel = element(cxml)

        dataModel.remove_drawing_ref()

        assert _without_nsdecls(dataModel.xml) == _without_nsdecls(xml(expected_cxml))


class DescribeCT_Pt:
    """Unit-test suite for `pptx.oxml.diagram.CT_Pt` objects."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("dgm:pt{modelId=1}", "node"),
            ("dgm:pt{modelId=1,type=asst}", "asst"),
            ("dgm:pt{modelId=1,type=doc}", "doc"),
        ],
    )
    def it_knows_its_type(self, cxml: str, expected_value: str):
        pt = element(cxml)
        assert isinstance(pt, CT_Pt)
        assert pt.type == expected_value

    def it_provides_access_to_its_text_body(self):
        pt = element('dgm:pt{modelId=1}/(dgm:prSet,dgm:spPr,dgm:t/(a:bodyPr,a:p/a:r/a:t"foo"))')

        t = pt.t

        assert isinstance(t, CT_TextBody)
        assert [p.text for p in t.p_lst] == ["foo"]

    def it_adds_a_text_body_in_schema_order(self):
        pt = element("dgm:pt{modelId=1}/(dgm:prSet,dgm:spPr,dgm:extLst)")

        t = pt.get_or_add_t()

        assert isinstance(t, CT_TextBody)
        assert _without_nsdecls(pt.xml) == _without_nsdecls(
            xml("dgm:pt{modelId=1}/(dgm:prSet,dgm:spPr,dgm:t/(a:bodyPr,a:lstStyle,a:p),dgm:extLst)")
        )

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("dgm:pt{modelId=1}/dgm:prSet{phldr=1}", True),
            ("dgm:pt{modelId=1}/dgm:prSet", None),
        ],
    )
    def it_knows_whether_its_text_is_placeholder_text(
        self, cxml: str, expected_value: bool | None
    ):
        assert element(cxml).prSet.phldr is expected_value


class DescribeCT_Cxn:
    """Unit-test suite for `pptx.oxml.diagram.CT_Cxn` objects."""

    def it_knows_its_endpoints_and_order(self):
        cxn = element("dgm:cxn{modelId=c1,type=presOf,srcId=1,destId=2,srcOrd=3,destOrd=4}")

        assert isinstance(cxn, CT_Cxn)
        assert (cxn.modelId, cxn.type, cxn.srcId, cxn.destId) == ("c1", "presOf", "1", "2")
        assert (cxn.srcOrd, cxn.destOrd) == (3, 4)

    def it_defaults_to_a_first_parent_of_connection(self):
        cxn = element("dgm:cxn{modelId=c1,srcId=1,destId=2}")

        assert (cxn.type, cxn.srcOrd, cxn.destOrd) == ("parOf", 0, 0)
