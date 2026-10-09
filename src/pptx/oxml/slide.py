"""Slide-related custom element classes, including those for masters."""

from __future__ import annotations

import random
from typing import TYPE_CHECKING, Callable, cast

from pptx.oxml import parse_from_template, parse_xml
from pptx.oxml.dml.fill import CT_GradientFillProperties
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.simpletypes import XsdBoolean, XsdString
from pptx.oxml.tags import CT_CustomerDataList, CustDataLstOwnerMixin
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    Choice,
    OneAndOnlyOne,
    OptionalAttribute,
    OxmlElement,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
    ZeroOrOneChoice,
)

if TYPE_CHECKING:
    from pptx.oxml.shapes.groupshape import CT_GroupShape
    from pptx.oxml.text import CT_TextListStyle


# -- `p:ext/@uri` of the slide extension holding `p188:commentRel`, per [MS-PPTX] --
COMMENT_REL_EXT_URI = "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}"
# -- `p:ext/@uri` of the common-slide-data extension holding `p14:creationId`, per [MS-PPTX] --
CREATION_ID_EXT_URI = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"


class _BaseSlideElement(BaseOxmlElement):
    """Base class for the six slide types, providing common methods."""

    cSld: CT_CommonSlideData

    @property
    def spTree(self) -> CT_GroupShape:
        """Return required `p:cSld/p:spTree` grandchild."""
        return self.cSld.spTree


class CT_Background(BaseOxmlElement):
    """`p:bg` element."""

    _insert_bgPr: Callable[[CT_BackgroundProperties], None]

    # ---these two are actually a choice, not a sequence, but simpler for
    # ---present purposes this way.
    _tag_seq = ("p:bgPr", "p:bgRef")
    bgPr: CT_BackgroundProperties | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:bgPr", successors=()
    )
    bgRef = ZeroOrOne("p:bgRef", successors=())
    del _tag_seq

    def add_noFill_bgPr(self):
        """Return a new `p:bgPr` element with noFill properties."""
        xml = "<p:bgPr %s>\n" "  <a:noFill/>\n" "  <a:effectLst/>\n" "</p:bgPr>" % nsdecls("a", "p")
        bgPr = cast(CT_BackgroundProperties, parse_xml(xml))
        self._insert_bgPr(bgPr)
        return bgPr


class CT_BackgroundProperties(BaseOxmlElement):
    """`p:bgPr` element."""

    _tag_seq = (
        "a:noFill",
        "a:solidFill",
        "a:gradFill",
        "a:blipFill",
        "a:pattFill",
        "a:grpFill",
        "a:effectLst",
        "a:effectDag",
        "a:extLst",
    )
    eg_fillProperties = ZeroOrOneChoice(
        (
            Choice("a:noFill"),
            Choice("a:solidFill"),
            Choice("a:gradFill"),
            Choice("a:blipFill"),
            Choice("a:pattFill"),
            Choice("a:grpFill"),
        ),
        successors=_tag_seq[6:],
    )
    del _tag_seq

    def _new_gradFill(self):
        """Override default to add default gradient subtree."""
        return CT_GradientFillProperties.new_gradFill()


class CT_CommonSlideData(CustDataLstOwnerMixin, BaseOxmlElement):
    """`p:cSld` element."""

    _remove_bg: Callable[[], None]
    get_or_add_bg: Callable[[], CT_Background]
    get_or_add_extLst: Callable[[], BaseOxmlElement]

    _tag_seq = ("p:bg", "p:spTree", "p:custDataLst", "p:controls", "p:extLst")
    bg: CT_Background | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:bg", successors=_tag_seq[1:]
    )
    spTree: CT_GroupShape = OneAndOnlyOne("p:spTree")  # pyright: ignore[reportAssignmentType]
    custDataLst: CT_CustomerDataList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:custDataLst", successors=_tag_seq[3:]
    )
    extLst: BaseOxmlElement | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:extLst", successors=()
    )
    del _tag_seq
    name: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "name", XsdString, default=""
    )

    @property
    def creation_id(self) -> int | None:
        """Value of the `p14:creationId` extension identifying this slide, or None if absent."""
        vals = self.xpath(
            "./p:extLst/p:ext[@uri='%s']/p14:creationId/@val" % CREATION_ID_EXT_URI
        )
        return int(vals[0]) if vals else None

    def renew_creation_id(self) -> None:
        """Give a slide that has a `p14:creationId` a new random one, as a copied slide needs."""
        vals = self.xpath(
            "./p:extLst/p:ext[@uri='%s']/p14:creationId" % CREATION_ID_EXT_URI
        )
        for creationId in vals:
            creationId.set("val", str(random.randint(1, 0xFFFFFFFF)))

    def get_or_add_creation_id(self) -> int:
        """Return the `p14:creationId` value, first adding a random one when there is none.

        PowerPoint gives every slide a creation id; modern comments name their slide by it.
        """
        creation_id = self.creation_id
        if creation_id is not None:
            return creation_id
        creation_id = random.randint(1, 0xFFFFFFFF)
        ext = OxmlElement("p:ext")
        ext.set("uri", CREATION_ID_EXT_URI)
        creationId = OxmlElement("p14:creationId")
        creationId.set("val", str(creation_id))
        ext.append(creationId)
        self.get_or_add_extLst().append(ext)
        return creation_id

    def get_or_add_bgPr(self) -> CT_BackgroundProperties:
        """Return `p:bg/p:bgPr` grandchild.

        If no such grandchild is present, any existing `p:bg` child is first removed and a new
        default `p:bg` with noFill settings is added.
        """
        bg = self.bg
        if bg is None or bg.bgPr is None:
            bg = self._change_to_noFill_bg()
        return cast(CT_BackgroundProperties, bg.bgPr)

    def _change_to_noFill_bg(self) -> CT_Background:
        """Establish a `p:bg` child with no-fill settings.

        Any existing `p:bg` child is first removed.
        """
        self._remove_bg()
        bg = self.get_or_add_bg()
        bg.add_noFill_bgPr()
        return bg


class CT_HeaderFooter(BaseOxmlElement):
    """`p:hf` element, specifying header/footer visibility on slides.

    Controls visibility of date/time, footer, and slide number placeholders.
    """

    sldNum: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "sldNum", XsdBoolean
    )
    hdr: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "hdr", XsdBoolean
    )
    ftr: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "ftr", XsdBoolean
    )
    dt: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dt", XsdBoolean
    )


class CT_NotesMaster(_BaseSlideElement):
    """`p:notesMaster` element, root of a notes master part."""

    get_or_add_hf: Callable[[], CT_HeaderFooter]

    _tag_seq = ("p:cSld", "p:clrMap", "p:hf", "p:notesStyle", "p:extLst")
    cSld: CT_CommonSlideData = OneAndOnlyOne("p:cSld")  # pyright: ignore[reportAssignmentType]
    hf: CT_HeaderFooter | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:hf", successors=_tag_seq[3:]
    )
    del _tag_seq

    @classmethod
    def new_default(cls) -> CT_NotesMaster:
        """Return a new `p:notesMaster` element based on the built-in default template."""
        return cast(CT_NotesMaster, parse_from_template("notesMaster"))


class CT_NotesSlide(_BaseSlideElement):
    """`p:notes` element, root of a notes slide part."""

    _tag_seq = ("p:cSld", "p:clrMapOvr", "p:extLst")
    cSld: CT_CommonSlideData = OneAndOnlyOne("p:cSld")  # pyright: ignore[reportAssignmentType]
    del _tag_seq

    @classmethod
    def new(cls) -> CT_NotesSlide:
        """Return a new ``<p:notes>`` element based on the default template.

        Note that the template does not include placeholders, which must be subsequently cloned
        from the notes master.
        """
        return cast(CT_NotesSlide, parse_from_template("notes"))


class CT_Slide(_BaseSlideElement):
    """`p:sld` element, root element of a slide part (XML document)."""

    get_or_add_extLst: Callable[[], BaseOxmlElement]
    get_or_add_hf: Callable[[], CT_HeaderFooter]

    _tag_seq = ("p:cSld", "p:clrMapOvr", "p:transition", "p:timing", "p:hf", "p:extLst")
    cSld: CT_CommonSlideData = OneAndOnlyOne("p:cSld")  # pyright: ignore[reportAssignmentType]
    clrMapOvr = ZeroOrOne("p:clrMapOvr", successors=_tag_seq[2:])
    transition = ZeroOrOne("p:transition", successors=_tag_seq[3:])
    timing = ZeroOrOne("p:timing", successors=_tag_seq[4:])
    hf: CT_HeaderFooter | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:hf", successors=_tag_seq[5:]
    )
    extLst: BaseOxmlElement | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:extLst", successors=()
    )
    del _tag_seq

    @classmethod
    def new(cls) -> CT_Slide:
        """Return new `p:sld` element configured as base slide shape."""
        return cast(CT_Slide, parse_xml(cls._sld_xml()))

    @property
    def comment_rel_rId(self) -> str | None:
        """`r:id` of the `p188:commentRel` extension naming the modern comments part, or None."""
        rIds = self.xpath(
            "./p:extLst/p:ext[@uri='%s']/p188:commentRel/@r:id" % COMMENT_REL_EXT_URI
        )
        return str(rIds[0]) if rIds else None

    def set_comment_rel(self, rId: str) -> None:
        """Point the `p188:commentRel` extension at relationship `rId`, adding it if needed."""
        self.remove_comment_rel()
        ext = OxmlElement("p:ext")
        ext.set("uri", COMMENT_REL_EXT_URI)
        commentRel = OxmlElement("p188:commentRel")
        commentRel.set(qn("r:id"), rId)
        ext.append(commentRel)
        self.get_or_add_extLst().append(ext)

    def remove_comment_rel(self) -> None:
        """Remove the `p188:commentRel` extension, and `p:extLst` if that leaves it empty."""
        for ext in self.xpath("./p:extLst/p:ext[@uri='%s']" % COMMENT_REL_EXT_URI):
            extLst = ext.getparent()
            extLst.remove(ext)
            if len(extLst) == 0:
                self.remove(extLst)

    @property
    def bg(self):
        """Return `p:bg` grandchild or None if not present."""
        return self.cSld.bg

    def get_or_add_childTnLst(self):
        """Return parent element for a new `p:video` child element.

        The `p:video` element causes play controls to appear under a video
        shape (pic shape containing video). There can be more than one video
        shape on a slide, which causes the precondition to vary. It needs to
        handle the case when there is no `p:sld/p:timing` element and when
        that element already exists. If the case isn't simple, it just nukes
        what's there and adds a fresh one. This could theoretically remove
        desired existing timing information, but there isn't any evidence
        available to me one way or the other, so I've taken the simple
        approach.
        """
        childTnLst = self._childTnLst
        if childTnLst is None:
            childTnLst = self._add_childTnLst()
        return childTnLst

    def _add_childTnLst(self):
        """Add `./p:timing/p:tnLst/p:par/p:cTn/p:childTnLst` descendant.

        Any existing `p:timing` child element is ruthlessly removed and
        replaced.
        """
        self.remove(self.get_or_add_timing())
        timing = parse_xml(self._childTnLst_timing_xml())
        self._insert_timing(timing)
        return timing.xpath("./p:tnLst/p:par/p:cTn/p:childTnLst")[0]

    @property
    def _childTnLst(self):
        """Return `./p:timing/p:tnLst/p:par/p:cTn/p:childTnLst` descendant.

        Return None if that element is not present.
        """
        childTnLsts = self.xpath("./p:timing/p:tnLst/p:par/p:cTn/p:childTnLst")
        if not childTnLsts:
            return None
        return childTnLsts[0]

    @staticmethod
    def _childTnLst_timing_xml():
        return (
            "<p:timing %s>\n"
            "  <p:tnLst>\n"
            "    <p:par>\n"
            '      <p:cTn id="1" dur="indefinite" restart="never" nodeType="'
            'tmRoot">\n'
            "        <p:childTnLst/>\n"
            "      </p:cTn>\n"
            "    </p:par>\n"
            "  </p:tnLst>\n"
            "</p:timing>" % nsdecls("p")
        )

    @staticmethod
    def _sld_xml():
        return (
            "<p:sld %s>\n"
            "  <p:cSld>\n"
            "    <p:spTree>\n"
            "      <p:nvGrpSpPr>\n"
            '        <p:cNvPr id="1" name=""/>\n'
            "        <p:cNvGrpSpPr/>\n"
            "        <p:nvPr/>\n"
            "      </p:nvGrpSpPr>\n"
            "      <p:grpSpPr/>\n"
            "    </p:spTree>\n"
            "  </p:cSld>\n"
            "  <p:clrMapOvr>\n"
            "    <a:masterClrMapping/>\n"
            "  </p:clrMapOvr>\n"
            "</p:sld>" % nsdecls("a", "p", "r")
        )


class CT_SlideLayout(_BaseSlideElement):
    """`p:sldLayout` element, root of a slide layout part."""

    get_or_add_hf: Callable[[], CT_HeaderFooter]

    _tag_seq = ("p:cSld", "p:clrMapOvr", "p:transition", "p:timing", "p:hf", "p:extLst")
    cSld: CT_CommonSlideData = OneAndOnlyOne("p:cSld")  # pyright: ignore[reportAssignmentType]
    transition = ZeroOrOne("p:transition", successors=_tag_seq[3:])
    hf: CT_HeaderFooter | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:hf", successors=_tag_seq[5:]
    )
    del _tag_seq


class CT_SlideLayoutIdList(BaseOxmlElement):
    """`p:sldLayoutIdLst` element, child of `p:sldMaster`.

    Contains references to the slide layouts that inherit from the slide master.
    """

    sldLayoutId_lst: list[CT_SlideLayoutIdListEntry]
    _add_sldLayoutId: Callable[..., CT_SlideLayoutIdListEntry]

    sldLayoutId = ZeroOrMore("p:sldLayoutId")

    def add_sldLayoutId(self, rId: str) -> CT_SlideLayoutIdListEntry:
        """Create and return a new `p:sldLayoutId` child element with `rId`."""
        return self._add_sldLayoutId(rId=rId)


class CT_SlideLayoutIdListEntry(BaseOxmlElement):
    """`p:sldLayoutId` element, child of `p:sldLayoutIdLst`.

    Contains a reference to a slide layout.
    """

    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_SlideMaster(_BaseSlideElement):
    """`p:sldMaster` element, root of a slide master part."""

    get_or_add_hf: Callable[[], CT_HeaderFooter]
    get_or_add_sldLayoutIdLst: Callable[[], CT_SlideLayoutIdList]
    get_or_add_txStyles: Callable[[], CT_SlideMasterTextStyles]

    _tag_seq = (
        "p:cSld",
        "p:clrMap",
        "p:sldLayoutIdLst",
        "p:transition",
        "p:timing",
        "p:hf",
        "p:txStyles",
        "p:extLst",
    )
    cSld: CT_CommonSlideData = OneAndOnlyOne("p:cSld")  # pyright: ignore[reportAssignmentType]
    sldLayoutIdLst: CT_SlideLayoutIdList = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:sldLayoutIdLst", successors=_tag_seq[3:]
    )
    hf: CT_HeaderFooter | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:hf", successors=_tag_seq[6:]
    )
    txStyles: CT_SlideMasterTextStyles | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:txStyles", successors=_tag_seq[7:]
    )
    del _tag_seq


class CT_SlideMasterTextStyles(BaseOxmlElement):
    """`p:txStyles` element, child of `p:sldMaster`.

    Holds the master's default formatting for title, body, and other text, each a
    `CT_TextListStyle` with per-level paragraph properties.
    """

    get_or_add_titleStyle: Callable[[], CT_TextListStyle]
    get_or_add_bodyStyle: Callable[[], CT_TextListStyle]
    get_or_add_otherStyle: Callable[[], CT_TextListStyle]

    _tag_seq = ("p:titleStyle", "p:bodyStyle", "p:otherStyle", "p:extLst")
    titleStyle: CT_TextListStyle | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:titleStyle", successors=_tag_seq[1:]
    )
    bodyStyle: CT_TextListStyle | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:bodyStyle", successors=_tag_seq[2:]
    )
    otherStyle: CT_TextListStyle | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:otherStyle", successors=_tag_seq[3:]
    )
    del _tag_seq
