"""Custom element classes for embedded-font related XML elements.

These implement `p:embeddedFontLst` (`CT_EmbeddedFontList`), the direct child of
`p:presentation` that lists the fonts embedded in the package, one `p:embeddedFont`
(`CT_EmbeddedFontListEntry`) per typeface. Each entry identifies its typeface via a
`p:font` (`CT_Font`) child and references the font-data parts for whichever of its
regular/bold/italic/bold-italic style variants are embedded via `p:regular`, `p:bold`,
`p:italic`, and `p:boldItalic` (`CT_EmbeddedFontDataId`) children, each an `r:id`
reference to a font-data part.
"""

from __future__ import annotations

from typing import Callable, cast

from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import XsdInt, XsdString
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    OxmlElement,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


class CT_Font(BaseOxmlElement):
    """`p:font` element, child of `p:embeddedFont` identifying the embedded typeface.

    Same attribute set as `a:CT_TextFont` (typeface, panose, pitchFamily, charset) but
    a distinct element class since it occurs on a different tag (`p:font` rather than
    `a:latin` etc.).
    """

    typeface: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "typeface", XsdString
    )
    # -- panose is a hex-encoded 10-byte value; kept as opaque string, not parsed --
    panose: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "panose", XsdString
    )
    pitchFamily: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "pitchFamily", XsdInt
    )
    charset: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "charset", XsdInt
    )


class CT_EmbeddedFontDataId(BaseOxmlElement):
    """`p:regular`, `p:bold`, `p:italic`, or `p:boldItalic` element.

    Each is a bare `r:id` reference to the font-data part embedding that style variant
    of the containing `p:embeddedFont` typeface.
    """

    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_EmbeddedFontListEntry(BaseOxmlElement):
    """`p:embeddedFont` element, entry in `p:embeddedFontLst` for one typeface."""

    font: CT_Font = OneAndOnlyOne("p:font")  # pyright: ignore[reportAssignmentType]

    _tag_seq = ("p:font", "p:regular", "p:bold", "p:italic", "p:boldItalic")
    regular: CT_EmbeddedFontDataId | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:regular", successors=_tag_seq[2:]
    )
    bold: CT_EmbeddedFontDataId | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:bold", successors=_tag_seq[3:]
    )
    italic: CT_EmbeddedFontDataId | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:italic", successors=_tag_seq[4:]
    )
    boldItalic: CT_EmbeddedFontDataId | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:boldItalic", successors=()
    )
    del _tag_seq

    _add_regular: Callable[..., CT_EmbeddedFontDataId]
    _add_bold: Callable[..., CT_EmbeddedFontDataId]
    _add_italic: Callable[..., CT_EmbeddedFontDataId]
    _add_boldItalic: Callable[..., CT_EmbeddedFontDataId]

    _STYLE_ADD_METHOD_NAMES = {
        "regular": "_add_regular",
        "bold": "_add_bold",
        "italic": "_add_italic",
        "bold_italic": "_add_boldItalic",
    }

    @classmethod
    def new(cls, typeface: str) -> CT_EmbeddedFontListEntry:
        """Return a new "loose" `p:embeddedFont` element for `typeface`.

        The new element has no style-variant (regular/bold/italic/boldItalic) children;
        those are added separately with `.add_style()`.
        """
        entry = cast("CT_EmbeddedFontListEntry", OxmlElement("p:embeddedFont"))
        font = cast("CT_Font", OxmlElement("p:font"))
        font.typeface = typeface
        entry.append(font)
        return entry

    def add_style(self, style: str, rId: str) -> CT_EmbeddedFontDataId:
        """Add the font-data relationship for `style`.

        `style` is one of "regular", "bold", "italic", or "bold_italic". `rId` is the
        relationship id of the related font-data part.
        """
        add_method = getattr(self, self._STYLE_ADD_METHOD_NAMES[style])
        return add_method(rId=rId)

    def iter_style_rIds(self):
        """Generate (style, rId) pairs for each style-variant present on this entry."""
        for style, tag in (
            ("regular", "p:regular"),
            ("bold", "p:bold"),
            ("italic", "p:italic"),
            ("bold_italic", "p:boldItalic"),
        ):
            child = self.find(qn(tag))
            if child is not None:
                yield style, cast("CT_EmbeddedFontDataId", child).rId


class CT_EmbeddedFontList(BaseOxmlElement):
    """`p:embeddedFontLst` element, direct child of `p:presentation`.

    Lists the fonts embedded in the presentation package, one `p:embeddedFont` entry
    per typeface.
    """

    embeddedFont_lst: list[CT_EmbeddedFontListEntry]

    embeddedFont = ZeroOrMore("p:embeddedFont")

    def add_embeddedFont(self, typeface: str) -> CT_EmbeddedFontListEntry:
        """Return a newly added `p:embeddedFont` entry for `typeface`."""
        entry = CT_EmbeddedFontListEntry.new(typeface)
        self.append(entry)
        return entry
