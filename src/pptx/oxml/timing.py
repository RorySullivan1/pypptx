"""Custom element classes for a slide's timing tree (`p:timing`): animations and media nodes.

The tree is a nest of time nodes. `p:par` (parallel), `p:seq` (sequence) and `p:excl`
(exclusive) containers each hold a `p:cTn` (common time-node data) whose `p:childTnLst` holds
the next level. The leaves are behaviors such as `p:set` or `p:animEffect`, and the media nodes
`p:audio` / `p:video`. pypptx reads this tree; apart from the media nodes that `add_movie()` and
`add_audio()` append, it does not write it.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import (
    ST_DrawingElementId,
    ST_TLTime,
    ST_TLTimeNodePresetClassType,
    ST_TLTimeNodeType,
    XsdInt,
    XsdString,
    XsdUnsignedInt,
)
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from lxml.etree import _Element


class CT_SlideTiming(BaseOxmlElement):
    """`p:timing` element, specifying animations and timed behaviors."""

    _tag_seq = ("p:tnLst", "p:bldLst", "p:extLst")
    tnLst: CT_TimeNodeList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:tnLst", successors=_tag_seq[1:]
    )
    bldLst: CT_BuildList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:bldLst", successors=_tag_seq[2:]
    )
    del _tag_seq


class CT_TimeNodeList(BaseOxmlElement):
    """`p:tnLst`, `p:childTnLst` or `p:subTnLst` element, a list of time nodes."""

    @property
    def time_nodes(self) -> list[_Element]:
        """The time-node child elements, e.g. `p:par`, `p:seq` or `p:set`, in document order."""
        return [child for child in self.iterchildren() if isinstance(child.tag, str)]

    def add_video(self, shape_id):
        """Add a new `p:video` child element for movie having *shape_id*."""
        video_xml = (
            "<p:video %s>\n"
            '  <p:cMediaNode vol="80000">\n'
            '    <p:cTn id="%d" fill="hold" display="0">\n'
            "      <p:stCondLst>\n"
            '        <p:cond delay="indefinite"/>\n'
            "      </p:stCondLst>\n"
            "    </p:cTn>\n"
            "    <p:tgtEl>\n"
            '      <p:spTgt spid="%d"/>\n'
            "    </p:tgtEl>\n"
            "  </p:cMediaNode>\n"
            "</p:video>\n" % (nsdecls("p"), self._next_cTn_id, shape_id)
        )
        video = parse_xml(video_xml)
        self.append(video)

    def add_audio(self, shape_id):
        """Add a new `p:audio` child element for audio clip having *shape_id*."""
        audio_xml = (
            "<p:audio %s>\n"
            '  <p:cMediaNode vol="80000">\n'
            '    <p:cTn id="%d" fill="hold" display="0">\n'
            "      <p:stCondLst>\n"
            '        <p:cond delay="indefinite"/>\n'
            "      </p:stCondLst>\n"
            "    </p:cTn>\n"
            "    <p:tgtEl>\n"
            '      <p:spTgt spid="%d"/>\n'
            "    </p:tgtEl>\n"
            "  </p:cMediaNode>\n"
            "</p:audio>\n" % (nsdecls("p"), self._next_cTn_id, shape_id)
        )
        audio = parse_xml(audio_xml)
        self.append(audio)

    @property
    def _next_cTn_id(self):
        """Return the next available unique ID (int) for p:cTn element."""
        cTn_id_strs = self.xpath("/p:sld/p:timing//p:cTn/@id")
        ids = [int(id_str) for id_str in cTn_id_strs]
        return max(ids) + 1


class CT_TLTimeNodeContainer(BaseOxmlElement):
    """`p:par`, `p:seq` or `p:excl` element, a time node holding other time nodes.

    `p:seq` also has `p:prevCondLst` and `p:nextCondLst` children, the conditions that move
    the sequence back and forward.
    """

    cTn: CT_TLCommonTimeNodeData = OneAndOnlyOne("p:cTn")  # pyright: ignore[reportAssignmentType]
    prevCondLst: CT_TLTimeConditionList | None = ZeroOrOne(  # pyright: ignore
        "p:prevCondLst", successors=("p:nextCondLst",)
    )
    nextCondLst: CT_TLTimeConditionList | None = ZeroOrOne(  # pyright: ignore
        "p:nextCondLst", successors=()
    )


class CT_TLCommonTimeNodeData(BaseOxmlElement):
    """`p:cTn` element, the timing data shared by every kind of time node.

    On an animation effect it also carries the effect's preset: `presetClass` (entrance, exit,
    ...), `presetID` and `presetSubtype`.
    """

    _tag_seq = (
        "p:stCondLst",
        "p:endCondLst",
        "p:endSync",
        "p:iterate",
        "p:childTnLst",
        "p:subTnLst",
    )
    id: int | None = OptionalAttribute("id", XsdUnsignedInt)  # pyright: ignore
    presetID: int | None = OptionalAttribute("presetID", XsdInt)  # pyright: ignore
    presetClass: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "presetClass", ST_TLTimeNodePresetClassType
    )
    presetSubtype: int | None = OptionalAttribute("presetSubtype", XsdInt)  # pyright: ignore
    dur: int | str | None = OptionalAttribute("dur", ST_TLTime)  # pyright: ignore
    nodeType: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "nodeType", ST_TLTimeNodeType
    )
    grpId: int | None = OptionalAttribute("grpId", XsdUnsignedInt)  # pyright: ignore
    stCondLst: CT_TLTimeConditionList | None = ZeroOrOne(  # pyright: ignore
        "p:stCondLst", successors=_tag_seq[1:]
    )
    childTnLst: CT_TimeNodeList | None = ZeroOrOne(  # pyright: ignore
        "p:childTnLst", successors=_tag_seq[5:]
    )
    del _tag_seq

    @property
    def child_time_nodes(self) -> list[_Element]:
        """The time nodes in this node's `p:childTnLst`, empty when it has none."""
        childTnLst = self.childTnLst
        return [] if childTnLst is None else childTnLst.time_nodes


class CT_TLTimeConditionList(BaseOxmlElement):
    """`p:stCondLst`, `p:endCondLst`, `p:prevCondLst` or `p:nextCondLst` element."""

    cond_lst: list[CT_TLTimeCondition]

    cond = ZeroOrMore("p:cond")


class CT_TLTimeCondition(BaseOxmlElement):
    """`p:cond` element, a condition that starts or ends a time node.

    `delay` is in milliseconds, or `ST_TLTime.INDEFINITE` (`"indefinite"`); `evt` names an
    event, such as `"onClick"` or `"onNext"`, and `p:tgtEl` its target.
    """

    evt: str | None = OptionalAttribute("evt", XsdString)  # pyright: ignore
    delay: int | str | None = OptionalAttribute("delay", ST_TLTime)  # pyright: ignore
    tgtEl: CT_TLTimeTargetElement | None = ZeroOrOne(  # pyright: ignore
        "p:tgtEl", successors=("p:tn", "p:rtn")
    )


class CT_TLTimeTargetElement(BaseOxmlElement):
    """`p:tgtEl` element, the target of a behavior or condition.

    One of a slide (`p:sldTgt`), a sound (`p:sndTgt`), a shape or part of one (`p:spTgt`) or
    ink (`p:inkTgt`).
    """

    spTgt: CT_TLShapeTargetElement | None = ZeroOrOne(  # pyright: ignore
        "p:spTgt", successors=("p:inkTgt",)
    )


class CT_TLShapeTargetElement(BaseOxmlElement):
    """`p:spTgt` element, a shape, or the text in a range of its paragraphs (`p:txEl`)."""

    spid: int = RequiredAttribute("spid", ST_DrawingElementId)  # pyright: ignore
    txEl: CT_TLTextTargetElement | None = ZeroOrOne(  # pyright: ignore
        "p:txEl", successors=("p:graphicEl",)
    )


class CT_TLTextTargetElement(BaseOxmlElement):
    """`p:txEl` element, a range of characters (`p:charRg`) or paragraphs (`p:pRg`)."""

    charRg: CT_IndexRange | None = ZeroOrOne(  # pyright: ignore
        "p:charRg", successors=("p:pRg",)
    )
    pRg: CT_IndexRange | None = ZeroOrOne("p:pRg", successors=())  # pyright: ignore


class CT_IndexRange(BaseOxmlElement):
    """`p:pRg` or `p:charRg` element, an inclusive range of zero-based indexes."""

    st: int = RequiredAttribute("st", XsdUnsignedInt)  # pyright: ignore
    end: int = RequiredAttribute("end", XsdUnsignedInt)  # pyright: ignore


class CT_TLBehavior(BaseOxmlElement):
    """A behavior element: `p:set`, `p:anim`, `p:animEffect`, `p:animMotion`, `p:animScale`,
    `p:animRot`, `p:animClr` or `p:cmd`.

    Each starts with `p:cBhvr`, which holds its timing and its target.
    """

    cBhvr: CT_TLCommonBehaviorData = OneAndOnlyOne("p:cBhvr")  # pyright: ignore


class CT_TLCommonBehaviorData(BaseOxmlElement):
    """`p:cBhvr` element, the timing (`p:cTn`) and target (`p:tgtEl`) of a behavior."""

    cTn: CT_TLCommonTimeNodeData = OneAndOnlyOne("p:cTn")  # pyright: ignore
    tgtEl: CT_TLTimeTargetElement = OneAndOnlyOne("p:tgtEl")  # pyright: ignore


class CT_TLMediaNodeVideo(BaseOxmlElement):
    """`p:video` element, specifying video media details."""

    _tag_seq = ("p:cMediaNode",)
    cMediaNode: CT_TLCommonMediaNodeData = OneAndOnlyOne("p:cMediaNode")  # pyright: ignore
    del _tag_seq


class CT_TLMediaNodeAudio(BaseOxmlElement):
    """`p:audio` element, specifying audio media details."""

    _tag_seq = ("p:cMediaNode",)
    cMediaNode: CT_TLCommonMediaNodeData = OneAndOnlyOne("p:cMediaNode")  # pyright: ignore
    del _tag_seq


class CT_TLCommonMediaNodeData(BaseOxmlElement):
    """`p:cMediaNode` element, the timing (`p:cTn`) and target (`p:tgtEl`) of a media node."""

    cTn: CT_TLCommonTimeNodeData = OneAndOnlyOne("p:cTn")  # pyright: ignore
    tgtEl: CT_TLTimeTargetElement = OneAndOnlyOne("p:tgtEl")  # pyright: ignore


class CT_BuildList(BaseOxmlElement):
    """`p:bldLst` element, how each animated shape's content is built up."""

    bldP_lst: list[CT_TLBuildParagraph]

    bldP = ZeroOrMore("p:bldP")


class CT_TLBuildParagraph(BaseOxmlElement):
    """`p:bldP` element, the text build of one shape, e.g. paragraph by paragraph.

    `grpId` matches the `grpId` of the effects that animate the shape `spid`; `build` is
    `"whole"` (the default), `"p"` (by paragraph), `"cust"` or `"allAtOnce"`.
    """

    spid: str = RequiredAttribute("spid", XsdString)  # pyright: ignore
    grpId: int = RequiredAttribute("grpId", XsdUnsignedInt)  # pyright: ignore
    build: str = OptionalAttribute("build", XsdString, default="whole")  # pyright: ignore
