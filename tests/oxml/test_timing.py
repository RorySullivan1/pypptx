"""Unit-test suite for the `pptx.oxml.timing` module."""

from __future__ import annotations

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import ST_TLTime
from pptx.oxml.timing import (
    CT_BuildList,
    CT_SlideTiming,
    CT_TimeNodeList,
    CT_TLBehavior,
    CT_TLCommonTimeNodeData,
    CT_TLMediaNodeVideo,
    CT_TLTimeNodeContainer,
)

from ..unitutil.cxml import element

# -- the shape of what PowerPoint writes for one by-paragraph "Fly In" entrance that starts
# -- after the slide begins, then one media node; trimmed from customGeo.pptx slide 18 --
TIMING_XML = (
    f"<p:timing {nsdecls('p')}>"
    "<p:tnLst><p:par>"
    '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
    '<p:seq concurrent="1" nextAc="seek">'
    '<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
    '<p:par><p:cTn id="3" fill="hold" nodeType="clickPar">'
    '<p:stCondLst><p:cond delay="indefinite"/>'
    '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst>'
    "<p:childTnLst>"
    '<p:par><p:cTn id="4" fill="hold" nodeType="withGroup">'
    '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
    '<p:par><p:cTn id="5" presetID="2" presetClass="entr" presetSubtype="4" fill="hold"'
    ' grpId="0" nodeType="afterEffect">'
    '<p:stCondLst><p:cond delay="250"/></p:stCondLst><p:childTnLst>'
    '<p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"/>'
    '<p:tgtEl><p:spTgt spid="3"><p:txEl><p:pRg st="0" end="1"/></p:txEl></p:spTgt></p:tgtEl>'
    "</p:cBhvr></p:set>"
    '<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
    '<p:cTn id="7" dur="500" fill="hold"/>'
    '<p:tgtEl><p:spTgt spid="3"><p:txEl><p:pRg st="0" end="1"/></p:txEl></p:spTgt></p:tgtEl>'
    "</p:cBhvr></p:anim>"
    "</p:childTnLst></p:cTn></p:par>"
    "</p:childTnLst></p:cTn></p:par>"
    "</p:childTnLst></p:cTn></p:par>"
    "</p:childTnLst></p:cTn>"
    '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
    "</p:prevCondLst>"
    '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
    "</p:nextCondLst>"
    "</p:seq>"
    '<p:video><p:cMediaNode vol="80000"><p:cTn id="8" fill="hold" display="0">'
    '<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn>'
    '<p:tgtEl><p:spTgt spid="7"/></p:tgtEl></p:cMediaNode></p:video>'
    "</p:childTnLst></p:cTn>"
    "</p:par></p:tnLst>"
    '<p:bldLst><p:bldP spid="3" grpId="0" build="p"/></p:bldLst>'
    "</p:timing>"
)


def _timing() -> CT_SlideTiming:
    return parse_xml(TIMING_XML)


def _root_cTn(timing: CT_SlideTiming) -> CT_TLCommonTimeNodeData:
    tnLst = timing.tnLst
    assert tnLst is not None
    par = tnLst.time_nodes[0]
    assert isinstance(par, CT_TLTimeNodeContainer)
    return par.cTn


def _effect_cTn(timing: CT_SlideTiming) -> CT_TLCommonTimeNodeData:
    return timing.xpath(".//p:cTn[@presetClass]")[0]


class DescribeCT_SlideTiming:
    """Unit-test suite for `pptx.oxml.timing.CT_SlideTiming` objects."""

    def it_provides_access_to_its_time_node_list_and_build_list(self):
        timing = _timing()

        assert isinstance(timing, CT_SlideTiming)
        assert isinstance(timing.tnLst, CT_TimeNodeList)
        assert isinstance(timing.bldLst, CT_BuildList)

    def and_both_are_optional(self):
        timing = element("p:timing")

        assert timing.tnLst is None
        assert timing.bldLst is None


class DescribeCT_TimeNodeList:
    """Unit-test suite for `pptx.oxml.timing.CT_TimeNodeList` objects."""

    def it_lists_its_time_nodes_in_document_order(self):
        root_cTn = _root_cTn(_timing())

        tags = [node.tag.split("}")[1] for node in root_cTn.child_time_nodes]

        assert tags == ["seq", "video"]

    def but_it_skips_comments(self):
        childTnLst = parse_xml(
            f"<p:childTnLst {nsdecls('p')}><!-- x --><p:par><p:cTn/></p:par></p:childTnLst>"
        )

        assert [n.tag.split("}")[1] for n in childTnLst.time_nodes] == ["par"]


class DescribeCT_TLCommonTimeNodeData:
    """Unit-test suite for `pptx.oxml.timing.CT_TLCommonTimeNodeData` objects."""

    def it_knows_its_timing_and_node_type(self):
        root_cTn = _root_cTn(_timing())

        assert root_cTn.id == 1
        assert root_cTn.dur == ST_TLTime.INDEFINITE
        assert root_cTn.nodeType == "tmRoot"
        assert root_cTn.presetClass is None

    def it_knows_the_preset_of_an_effect(self):
        cTn = _effect_cTn(_timing())

        assert cTn.presetClass == "entr"
        assert cTn.presetID == 2
        assert cTn.presetSubtype == 4
        assert cTn.grpId == 0
        assert cTn.nodeType == "afterEffect"

    def it_provides_access_to_its_start_conditions(self):
        cTn = _effect_cTn(_timing())

        stCondLst = cTn.stCondLst
        assert stCondLst is not None
        assert [cond.delay for cond in stCondLst.cond_lst] == [250]

    def and_its_child_time_nodes(self):
        cTn = _effect_cTn(_timing())

        behaviors = cTn.child_time_nodes

        assert [type(b) for b in behaviors] == [CT_TLBehavior, CT_TLBehavior]
        assert element("p:cTn").child_time_nodes == []


class DescribeCT_TLTimeCondition:
    """Unit-test suite for `pptx.oxml.timing.CT_TLTimeCondition` objects."""

    def it_knows_its_event_and_delay(self):
        clickPar_cTn = _timing().xpath(".//p:cTn[@nodeType='clickPar']")[0]

        conds = clickPar_cTn.stCondLst.cond_lst

        assert [(c.evt, c.delay) for c in conds] == [
            (None, ST_TLTime.INDEFINITE),
            ("onBegin", 0),
        ]

    def it_provides_access_to_its_target(self):
        seq = _timing().xpath(".//p:seq")[0]

        cond = seq.nextCondLst.cond_lst[0]

        assert cond.evt == "onNext"
        assert cond.tgtEl is not None
        assert cond.tgtEl.spTgt is None


class DescribeCT_TLBehavior:
    """Unit-test suite for the behavior elements, `pptx.oxml.timing.CT_TLBehavior`."""

    def it_provides_access_to_its_timing_and_its_target(self):
        anim = _timing().xpath(".//p:anim")[0]

        cBhvr = anim.cBhvr
        assert cBhvr.cTn.dur == 500
        spTgt = cBhvr.tgtEl.spTgt
        assert spTgt is not None
        assert spTgt.spid == 3
        assert spTgt.txEl is not None
        assert spTgt.txEl.charRg is None
        pRg = spTgt.txEl.pRg
        assert pRg is not None
        assert (pRg.st, pRg.end) == (0, 1)


class DescribeCT_TLMediaNode:
    """Unit-test suite for the media nodes, e.g. `pptx.oxml.timing.CT_TLMediaNodeVideo`."""

    def it_provides_access_to_its_timing_and_its_target(self):
        video = _timing().xpath(".//p:video")[0]

        assert isinstance(video, CT_TLMediaNodeVideo)
        assert video.cMediaNode.cTn.id == 8
        spTgt = video.cMediaNode.tgtEl.spTgt
        assert spTgt is not None
        assert spTgt.spid == 7


class DescribeCT_BuildList:
    """Unit-test suite for `pptx.oxml.timing.CT_BuildList` objects."""

    def it_provides_access_to_its_paragraph_builds(self):
        bldLst = _timing().bldLst
        assert bldLst is not None

        (bldP,) = bldLst.bldP_lst

        assert (bldP.spid, bldP.grpId, bldP.build) == ("3", 0, "p")

    def and_a_build_is_whole_by_default(self):
        assert element("p:bldP{spid=3,grpId=0}").build == "whole"
