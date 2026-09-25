"""Custom element classes for SmartArt (DrawingML diagram) elements.

A SmartArt graphic frame on a slide references four parts through its `dgm:relIds` element:
the *data* part (`dgm:dataModel`, the content), plus the layout, quick-style and colors
definitions. PowerPoint also writes a fifth, *drawing* part (`dsp:drawing`) holding the
shapes it last laid out, which it uses as a display cache. Only the data model is given
element classes here; the definitions and the drawing are carried as opaque XML.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Iterator, cast

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import XsdBoolean, XsdString, XsdUnsignedInt
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from pptx.oxml.text import CT_TextBody

# -- `a:ext/@uri` of the extension that points a data model at its cached drawing part --
DATA_MODEL_EXT_URI = "http://schemas.microsoft.com/office/drawing/2008/diagram"

# -- `dgm:pt/@type` values of points that carry user content (as opposed to presentation and
# -- transition points); an omitted `type` means "node" --
CONTENT_POINT_TYPES = ("node", "asst")


class CT_DiagramRelIds(BaseOxmlElement):
    """`dgm:relIds` element, the payload of a SmartArt `a:graphicData`.

    Each attribute is the rId of a relationship from the slide part to one diagram part.
    """

    dm: str = RequiredAttribute("r:dm", XsdString)  # pyright: ignore[reportAssignmentType]
    lo: str = RequiredAttribute("r:lo", XsdString)  # pyright: ignore[reportAssignmentType]
    qs: str = RequiredAttribute("r:qs", XsdString)  # pyright: ignore[reportAssignmentType]
    cs: str = RequiredAttribute("r:cs", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_DataModel(BaseOxmlElement):
    """`dgm:dataModel` element, root of a diagram-data part.

    Holds every point (`dgm:pt`) of the diagram and the connections (`dgm:cxn`) between them.
    Content points form a tree rooted at the single "doc" point, linked by "parOf" connections.
    """

    _tag_seq = ("dgm:ptLst", "dgm:cxnLst", "dgm:bg", "dgm:whole", "dgm:extLst")
    ptLst: CT_PtList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "dgm:ptLst", successors=_tag_seq[1:]
    )
    cxnLst: CT_CxnList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "dgm:cxnLst", successors=_tag_seq[2:]
    )
    del _tag_seq

    @property
    def doc_pt(self) -> CT_Pt | None:
        """The "doc" `dgm:pt` element at the root of the content tree, or |None| if missing."""
        pts = cast("list[CT_Pt]", self.xpath("./dgm:ptLst/dgm:pt[@type='doc']"))
        return pts[0] if pts else None

    @property
    def drawing_rId(self) -> str | None:
        """rId (relative to the slide part) of this diagram's cached drawing part, if any."""
        relIds = cast(
            "list[str]",
            self.xpath(
                "./dgm:extLst/a:ext[@uri='%s']/dsp:dataModelExt/@relId" % DATA_MODEL_EXT_URI
            ),
        )
        return relIds[0] if relIds else None

    def child_pts(self, modelId: str) -> list[CT_Pt]:
        """Content points that are children of the point identified by `modelId`, in order.

        Children are the destinations of "parOf" connections from `modelId` (a connection with
        no `type` is "parOf"), sorted by the connection's `srcOrd`. Presentation and transition
        points are excluded.
        """
        pts_by_id = {pt.modelId: pt for pt in self.iter_pts()}
        cxns = sorted(
            (
                cxn
                for cxn in self.iter_cxns()
                if cxn.type == "parOf" and cxn.srcId == modelId and cxn.destId in pts_by_id
            ),
            key=lambda cxn: cxn.srcOrd,
        )
        return [
            pts_by_id[cxn.destId]
            for cxn in cxns
            if pts_by_id[cxn.destId].type in CONTENT_POINT_TYPES
        ]

    def iter_cxns(self) -> Iterator[CT_Cxn]:
        """Generate each `dgm:cxn` element, in document order."""
        cxnLst = self.cxnLst
        if cxnLst is None:
            return
        yield from cxnLst.cxn_lst

    def iter_pts(self) -> Iterator[CT_Pt]:
        """Generate each `dgm:pt` element, in document order."""
        ptLst = self.ptLst
        if ptLst is None:
            return
        yield from ptLst.pt_lst

    def remove_drawing_ref(self) -> None:
        """Remove the extension that points this data model at its cached drawing part."""
        for ext in self.xpath("./dgm:extLst/a:ext[@uri='%s']" % DATA_MODEL_EXT_URI):
            extLst = ext.getparent()
            extLst.remove(ext)
            if len(extLst) == 0:
                self.remove(extLst)


class CT_PtList(BaseOxmlElement):
    """`dgm:ptLst` element, the list of points in a data model."""

    pt_lst: list[CT_Pt]

    pt = ZeroOrMore("dgm:pt")


class CT_Pt(BaseOxmlElement):
    """`dgm:pt` element, one point (node) of a SmartArt data model.

    Content points ("node", "asst") carry the user's text in an optional `dgm:t` child, a
    DrawingML text body.
    """

    get_or_add_t: Callable[[], CT_TextBody]

    _tag_seq = ("dgm:prSet", "dgm:spPr", "dgm:t", "dgm:extLst")
    prSet: CT_ElemPropSet | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "dgm:prSet", successors=_tag_seq[1:]
    )
    t: CT_TextBody | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "dgm:t", successors=_tag_seq[3:]
    )
    del _tag_seq

    modelId: str = RequiredAttribute("modelId", XsdString)  # pyright: ignore[reportAssignmentType]
    type: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "type", XsdString, default="node"
    )

    def _new_t(self) -> CT_TextBody:
        return cast(
            "CT_TextBody",
            parse_xml("<dgm:t %s><a:bodyPr/><a:lstStyle/><a:p/></dgm:t>" % nsdecls("dgm", "a")),
        )


class CT_ElemPropSet(BaseOxmlElement):
    """`dgm:prSet` element, presentation properties of a point.

    `phldr` marks a point whose text is placeholder (prompt) text rather than user content.
    """

    phldr: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "phldr", XsdBoolean
    )


class CT_CxnList(BaseOxmlElement):
    """`dgm:cxnLst` element, the list of connections in a data model."""

    cxn_lst: list[CT_Cxn]

    cxn = ZeroOrMore("dgm:cxn")


class CT_Cxn(BaseOxmlElement):
    """`dgm:cxn` element, a directed connection between two points of a data model."""

    modelId: str = RequiredAttribute("modelId", XsdString)  # pyright: ignore[reportAssignmentType]
    type: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "type", XsdString, default="parOf"
    )
    srcId: str = RequiredAttribute("srcId", XsdString)  # pyright: ignore[reportAssignmentType]
    destId: str = RequiredAttribute("destId", XsdString)  # pyright: ignore[reportAssignmentType]
    srcOrd: int = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "srcOrd", XsdUnsignedInt, default=0
    )
    destOrd: int = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "destOrd", XsdUnsignedInt, default=0
    )
