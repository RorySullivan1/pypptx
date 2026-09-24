"""Custom element classes for presentation-related XML elements."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, cast

from pptx.oxml.simpletypes import (
    ST_SlideId,
    ST_SlideSizeCoordinate,
    XsdInt,
    XsdString,
    XsdUnsignedInt,
)
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

from pptx.oxml.ns import qn

if TYPE_CHECKING:
    from pptx.oxml.text import CT_TextListStyle
    from pptx.util import Length


class CT_Presentation(BaseOxmlElement):
    """`p:presentation` element, root of the Presentation part stored as `/ppt/presentation.xml`."""

    get_or_add_sldSz: Callable[[], CT_SlideSize]
    get_or_add_notesSz: Callable[[], CT_NotesSize]
    get_or_add_sldIdLst: Callable[[], CT_SlideIdList]
    get_or_add_sldMasterIdLst: Callable[[], CT_SlideMasterIdList]
    get_or_add_defaultTextStyle: Callable[[], CT_TextListStyle]
    get_or_add_custShowLst: Callable[[], CT_CustomShowList]

    # -- NOTE: `p:smartTags`, `p:embeddedFontLst`, `p:photoAlbum`, `p:custDataLst` and
    # -- `p:modifyVerifier` are valid schema positions between `p:notesSz` and `p:custShowLst`
    # -- and between `p:custShowLst` and `p:kinsoku` respectively, but are not yet modeled here.
    _tag_seq = (
        "p:sldMasterIdLst",
        "p:notesMasterIdLst",
        "p:handoutMasterIdLst",
        "p:sldIdLst",
        "p:sldSz",
        "p:notesSz",
        "p:custShowLst",
        "p:kinsoku",
        "p:defaultTextStyle",
        "p:extLst",
    )
    sldMasterIdLst: CT_SlideMasterIdList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:sldMasterIdLst", successors=_tag_seq[1:]
    )
    notesMasterIdLst = ZeroOrOne("p:notesMasterIdLst", successors=_tag_seq[2:])
    handoutMasterIdLst = ZeroOrOne("p:handoutMasterIdLst", successors=_tag_seq[3:])
    sldIdLst: CT_SlideIdList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:sldIdLst", successors=_tag_seq[4:]
    )
    sldSz: CT_SlideSize | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:sldSz", successors=_tag_seq[5:]
    )
    notesSz: CT_NotesSize | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:notesSz", successors=_tag_seq[6:]
    )
    custShowLst: CT_CustomShowList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:custShowLst", successors=_tag_seq[7:]
    )
    defaultTextStyle: CT_TextListStyle | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:defaultTextStyle", successors=_tag_seq[9:]
    )
    del _tag_seq

    firstSlideNum: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "firstSlideNum", XsdInt
    )

    _SECTION_EXT_URI = "{521415D9-36F7-43E2-AB2F-B2CE04A55DE4}"

    @property
    def sectionLst(self):
        """Return the `p14:sectionLst` element or None if not present."""
        from pptx.oxml.section import CT_SectionList

        results = self.xpath(
            "p:extLst/p:ext/p14:sectionLst",
        )
        return results[0] if results else None

    def get_or_add_sectionLst(self):
        """Return `p14:sectionLst`, creating the extension structure if needed."""
        from pptx.oxml.section import CT_SectionList

        sectionLst = self.sectionLst
        if sectionLst is not None:
            return sectionLst

        from lxml import etree

        extLst = self.find(qn("p:extLst"))
        if extLst is None:
            extLst = etree.SubElement(self, qn("p:extLst"))
        ext = etree.SubElement(extLst, qn("p:ext"))
        ext.set("uri", self._SECTION_EXT_URI)
        sectionLst = etree.SubElement(ext, qn("p14:sectionLst"))
        return sectionLst


class CT_SlideId(BaseOxmlElement):
    """`p:sldId` element.

    Direct child of `p:sldIdLst` that contains an `rId` reference to a slide in the presentation.
    """

    id: int = RequiredAttribute("id", ST_SlideId)  # pyright: ignore[reportAssignmentType]
    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_SlideIdList(BaseOxmlElement):
    """`p:sldIdLst` element.

    Direct child of <p:presentation> that contains a list of the slide parts in the presentation.
    """

    sldId_lst: list[CT_SlideId]

    _add_sldId: Callable[..., CT_SlideId]
    sldId = ZeroOrMore("p:sldId")

    def add_sldId(self, rId: str) -> CT_SlideId:
        """Create and return a reference to a new `p:sldId` child element.

        The new `p:sldId` element has its r:id attribute set to `rId`.
        """
        return self._add_sldId(id=self._next_id, rId=rId)

    @property
    def _next_id(self) -> int:
        """The next available slide ID as an `int`.

        Valid slide IDs start at 256. The next integer value greater than the max value in use is
        chosen, which minimizes that chance of reusing the id of a deleted slide.
        """
        MIN_SLIDE_ID = 256
        MAX_SLIDE_ID = 2147483647

        used_ids = [int(s) for s in cast("list[str]", self.xpath("./p:sldId/@id"))]
        simple_next = max([MIN_SLIDE_ID - 1] + used_ids) + 1
        if simple_next <= MAX_SLIDE_ID:
            return simple_next

        # -- fall back to search for next unused from bottom --
        valid_used_ids = sorted(id for id in used_ids if (MIN_SLIDE_ID <= id <= MAX_SLIDE_ID))
        return (
            next(
                candidate_id
                for candidate_id, used_id in enumerate(valid_used_ids, start=MIN_SLIDE_ID)
                if candidate_id != used_id
            )
            if valid_used_ids
            else 256
        )


class CT_SlideMasterIdList(BaseOxmlElement):
    """`p:sldMasterIdLst` element.

    Child of `p:presentation` containing references to the slide masters that belong to the
    presentation.
    """

    sldMasterId_lst: list[CT_SlideMasterIdListEntry]
    _add_sldMasterId: Callable[..., CT_SlideMasterIdListEntry]

    sldMasterId = ZeroOrMore("p:sldMasterId")

    def add_sldMasterId(self, rId: str) -> CT_SlideMasterIdListEntry:
        """Create and return a new `p:sldMasterId` child element with `rId`."""
        return self._add_sldMasterId(rId=rId)


class CT_SlideMasterIdListEntry(BaseOxmlElement):
    """
    ``<p:sldMasterId>`` element, child of ``<p:sldMasterIdLst>`` containing
    a reference to a slide master.
    """

    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_SlideSize(BaseOxmlElement):
    """`p:sldSz` element.

    Direct child of <p:presentation> that contains the width and height of slides in the
    presentation.
    """

    cx: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "cx", ST_SlideSizeCoordinate
    )
    cy: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "cy", ST_SlideSizeCoordinate
    )


class CT_NotesSize(BaseOxmlElement):
    """`p:notesSz` element.

    Direct child of <p:presentation> that contains the width and height of notes pages in the
    presentation.
    """

    cx: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "cx", ST_SlideSizeCoordinate
    )
    cy: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "cy", ST_SlideSizeCoordinate
    )


class CT_CustomShowSlideList(BaseOxmlElement):
    """`p:sldLst` element, child of `p:custShow` listing its member slides in order.

    NOTE: PowerPoint's schema reuses the `p:sld` tag for the `r:id`-only reference elements
    contained here -- the *same* tag used for a slide part's root element (`CT_Slide`).
    Because pypptx's element-class lookup is global by tag name, these children are parsed as
    whatever class is registered for `p:sld` (`CT_Slide`). To avoid depending on that
    class's unrelated shape, only the generic, tag-agnostic `r:id` attribute is read or
    written here rather than declaring a dedicated element class.
    """

    def sld_rIds(self) -> list[str]:
        """List of `r:id` values of the `p:sld` children, in document order."""
        return cast("list[str]", self.xpath("./p:sld/@r:id"))

    def add_sld(self, rId: str) -> BaseOxmlElement:
        """Create and return a new `p:sld` child element referencing `rId`."""
        from lxml import etree

        sld = etree.SubElement(self, qn("p:sld"))
        sld.set(qn("r:id"), rId)
        return sld

    def remove_sld_with_rId(self, rId: str) -> bool:
        """Remove the `p:sld` child referencing `rId`, if present.

        Returns |True| if a matching child was found and removed, |False| otherwise.
        """
        for sld in self.findall(qn("p:sld")):
            if sld.get(qn("r:id")) == rId:
                self.remove(sld)
                return True
        return False

    def clear_slds(self) -> None:
        """Remove all `p:sld` children."""
        for sld in self.findall(qn("p:sld")):
            self.remove(sld)


class CT_CustomShow(BaseOxmlElement):
    """`p:custShow` element, a named, ordered subset of the presentation's slides."""

    get_or_add_sldLst: Callable[[], CT_CustomShowSlideList]

    _tag_seq = ("p:sldLst", "p:extLst")
    sldLst: CT_CustomShowSlideList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:sldLst", successors=_tag_seq[1:]
    )
    del _tag_seq

    name: str = RequiredAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]
    id: int = RequiredAttribute("id", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]


class CT_CustomShowList(BaseOxmlElement):
    """`p:custShowLst` element, child of `p:presentation` containing the custom shows."""

    custShow_lst: list[CT_CustomShow]

    _add_custShow: Callable[..., CT_CustomShow]
    custShow = ZeroOrMore("p:custShow")

    def add_custShow(self, name: str, id: int) -> CT_CustomShow:
        """Create and return a new `p:custShow` child element with `name` and `id`."""
        return self._add_custShow(name=name, id=id)
