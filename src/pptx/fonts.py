"""Embedded-fonts API for a |Presentation| (issue #61).

Provides read access to the fonts embedded in a presentation package (`p:embeddedFontLst`
in `presentation.xml`) and support for removing an embedded font, e.g. to shrink a file
or drop a licensed font. Embedding *new* fonts is out of scope.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, Sequence, Union, overload

if TYPE_CHECKING:
    from pptx.oxml.embeddedfont import CT_EmbeddedFontListEntry
    from pptx.oxml.presentation import CT_Presentation
    from pptx.parts.presentation import PresentationPart


class EmbeddedFonts(Sequence["EmbeddedFont"]):
    """Sequence of |EmbeddedFont| objects embedded in a presentation.

    Supports `len()`, iteration, and indexed (including sliced) access. Access this
    collection via `Presentation.embedded_fonts`; it is not intended to be constructed
    directly.
    """

    def __init__(self, prs_elm: CT_Presentation, part: PresentationPart):
        self._prs_elm = prs_elm
        self._part = part

    def __len__(self) -> int:
        return len(self._entries)

    def __iter__(self) -> Iterator[EmbeddedFont]:
        return (EmbeddedFont(entry, self._part) for entry in self._entries)

    @overload
    def __getitem__(self, idx: int) -> EmbeddedFont: ...
    @overload
    def __getitem__(self, idx: slice) -> list[EmbeddedFont]: ...

    def __getitem__(
        self, idx: Union[int, slice]
    ) -> Union[EmbeddedFont, list[EmbeddedFont]]:
        entries = self._entries
        if isinstance(idx, slice):
            return [EmbeddedFont(entry, self._part) for entry in entries[idx]]
        return EmbeddedFont(entries[idx], self._part)

    def remove(self, font: EmbeddedFont) -> None:
        """Remove `font` from this presentation.

        Equivalent to `font.remove()`.
        """
        font.remove()

    @property
    def _entries(self) -> list[CT_EmbeddedFontListEntry]:
        """list of `p:embeddedFont` elements, in document order."""
        embeddedFontLst = self._prs_elm.embeddedFontLst
        if embeddedFontLst is None:
            return []
        return embeddedFontLst.embeddedFont_lst


class EmbeddedFont:
    """A single embedded typeface (`p:embeddedFont`) in a presentation.

    Not intended to be constructed directly; access via `Presentation.embedded_fonts`.
    """

    def __init__(self, entry: CT_EmbeddedFontListEntry, part: PresentationPart):
        self._entry = entry
        self._part = part

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, EmbeddedFont):
            return NotImplemented
        return self._entry is other._entry

    def __ne__(self, other: object) -> bool:
        result = self.__eq__(other)
        return result if result is NotImplemented else not result

    def __repr__(self) -> str:
        return f"<EmbeddedFont typeface={self.typeface!r} styles={self.styles!r}>"

    @property
    def typeface(self) -> str:
        """Name of the embedded typeface, e.g. `"Calibri"`."""
        return self._entry.font.typeface

    @property
    def styles(self) -> tuple[str, ...]:
        """Tuple of style-names embedded for this typeface.

        Each item is one of `"regular"`, `"bold"`, `"italic"`, or `"bold_italic"`; only
        style-variants actually embedded for this typeface are included, in that order.
        """
        return tuple(style for style, _rId in self._entry.iter_style_rIds())

    @property
    def has_regular(self) -> bool:
        """True when the regular (non-bold, non-italic) style is embedded."""
        return self._entry.regular is not None

    @property
    def has_bold(self) -> bool:
        """True when the bold style is embedded."""
        return self._entry.bold is not None

    @property
    def has_italic(self) -> bool:
        """True when the italic style is embedded."""
        return self._entry.italic is not None

    @property
    def has_bold_italic(self) -> bool:
        """True when the bold-italic style is embedded."""
        return self._entry.boldItalic is not None

    def remove(self) -> None:
        """Remove this embedded font from the presentation.

        Drops the relationship for each embedded style-variant (regular, bold, italic,
        bold-italic); the underlying font-data part is no longer reachable from the
        package and so is not written the next time the presentation is saved (parts
        are discovered by walking relationships starting from the package root).

        When this was the last remaining embedded font, the now-empty
        `p:embeddedFontLst` element is also removed and `embedTrueTypeFonts` is turned
        off on the presentation, so PowerPoint does not re-embed fonts the next time
        the file is saved from within the application.
        """
        part = self._part
        entry = self._entry

        for _style, rId in entry.iter_style_rIds():
            part.drop_rel(rId)

        embeddedFontLst = entry.getparent()
        embeddedFontLst.remove(entry)

        if len(embeddedFontLst.embeddedFont_lst) == 0:
            prs_elm = embeddedFontLst.getparent()
            prs_elm.remove(embeddedFontLst)
            prs_elm.embedTrueTypeFonts = False
