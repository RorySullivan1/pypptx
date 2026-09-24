"""Text-style objects: the per-level paragraph and font defaults of masters and presentations."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Iterator, cast

from pptx.text.text import Font

if TYPE_CHECKING:
    from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
    from pptx.oxml.slide import CT_SlideMaster
    from pptx.oxml.text import CT_TextListStyle, CT_TextParagraphProperties
    from pptx.util import Length


class MasterTextStyles:
    """The title, body, and other text styles of a slide master (`p:txStyles`).

    Accessed via :attr:`.SlideMaster.text_styles`. Text in slides, layouts, and placeholders
    that doesn't override a property inherits it from one of these three styles: title
    placeholders from :attr:`title`, body and other content placeholders from :attr:`body`,
    and text in non-placeholder shapes such as text boxes from :attr:`other`.
    """

    def __init__(self, sldMaster: CT_SlideMaster):
        self._sldMaster = sldMaster

    @property
    def body(self) -> TextListStyle:
        """|TextListStyle| for body placeholders (`p:bodyStyle`)."""
        return self._style("bodyStyle")

    @property
    def other(self) -> TextListStyle:
        """|TextListStyle| for text outside placeholders (`p:otherStyle`)."""
        return self._style("otherStyle")

    @property
    def title(self) -> TextListStyle:
        """|TextListStyle| for title placeholders (`p:titleStyle`)."""
        return self._style("titleStyle")

    def _style(self, tagname: str) -> TextListStyle:
        sldMaster = self._sldMaster

        def get_lstStyle() -> CT_TextListStyle | None:
            txStyles = sldMaster.txStyles
            return None if txStyles is None else getattr(txStyles, tagname)

        def get_or_add_lstStyle() -> CT_TextListStyle:
            return getattr(sldMaster.get_or_add_txStyles(), "get_or_add_%s" % tagname)()

        return TextListStyle(get_lstStyle, get_or_add_lstStyle)


class TextListStyle:
    """Nine levels of paragraph and font defaults, e.g. a master's body style.

    Indexed by zero-based paragraph level, matching :attr:`._Paragraph.level`, so
    ``style[0]`` is the first (outermost) level, stored as `a:lvl1pPr`. Supports ``len()``
    (always 9) and iteration. Reading an unset property adds no XML; writing one, or
    accessing a level's :attr:`~TextLevelStyle.font`, adds the elements it needs.
    """

    def __init__(
        self,
        get_lstStyle: Callable[[], CT_TextListStyle | None],
        get_or_add_lstStyle: Callable[[], CT_TextListStyle],
    ):
        self._get_lstStyle = get_lstStyle
        self._get_or_add_lstStyle = get_or_add_lstStyle

    def __getitem__(self, level: int) -> TextLevelStyle:
        """|TextLevelStyle| for zero-based paragraph `level` (0-8)."""
        if not isinstance(level, int) or isinstance(level, bool) or not 0 <= level <= 8:
            raise IndexError("text style level must be an int in range 0-8, got %r" % (level,))

        def get_pPr() -> CT_TextParagraphProperties | None:
            lstStyle = self._get_lstStyle()
            return None if lstStyle is None else lstStyle.lvl_pPr(level)

        return TextLevelStyle(
            get_pPr, lambda: self._get_or_add_lstStyle().get_or_add_lvl_pPr(level)
        )

    def __iter__(self) -> Iterator[TextLevelStyle]:
        for level in range(9):
            yield self[level]

    def __len__(self) -> int:
        return 9

    @property
    def default(self) -> TextLevelStyle:
        """|TextLevelStyle| for the level-independent defaults (`a:defPPr`)."""

        def get_pPr() -> CT_TextParagraphProperties | None:
            lstStyle = self._get_lstStyle()
            return None if lstStyle is None else lstStyle.defPPr

        return TextLevelStyle(get_pPr, lambda: self._get_or_add_lstStyle().get_or_add_defPPr())


class TextLevelStyle:
    """Paragraph and font defaults for one level of a |TextListStyle|.

    Wraps an `a:lvlNpPr` (or `a:defPPr`) element. Read-only access to a property that isn't
    set returns |None|, meaning the value is inherited from further up the style hierarchy.
    """

    def __init__(
        self,
        get_pPr: Callable[[], CT_TextParagraphProperties | None],
        get_or_add_pPr: Callable[[], CT_TextParagraphProperties],
    ):
        self._get_pPr = get_pPr
        self._get_or_add_pPr = get_or_add_pPr

    @property
    def alignment(self) -> PP_PARAGRAPH_ALIGNMENT | None:
        """Horizontal alignment for paragraphs at this level; |None| when inherited."""
        pPr = self._get_pPr()
        return None if pPr is None else pPr.algn

    @alignment.setter
    def alignment(self, value: PP_PARAGRAPH_ALIGNMENT | None) -> None:
        self._get_or_add_pPr().algn = value

    @property
    def font(self) -> Font:
        """|Font| holding the default run formatting for this level (`a:defRPr`)."""
        return Font(self._get_or_add_pPr().get_or_add_defRPr())

    @property
    def indent(self) -> Length | None:
        """First-line indent relative to :attr:`margin_left`; |None| when inherited."""
        pPr = self._get_pPr()
        return None if pPr is None else cast("Length | None", pPr.indent)

    @indent.setter
    def indent(self, value: Length | None) -> None:
        self._get_or_add_pPr().indent = value

    @property
    def line_spacing(self) -> float | Length | None:
        """Line spacing as a number of lines (float) or fixed |Length|; |None| when inherited."""
        pPr = self._get_pPr()
        return None if pPr is None else pPr.line_spacing

    @line_spacing.setter
    def line_spacing(self, value: float | Length | None) -> None:
        self._get_or_add_pPr().line_spacing = value

    @property
    def margin_left(self) -> Length | None:
        """Left margin of paragraphs at this level; |None| when inherited."""
        pPr = self._get_pPr()
        return None if pPr is None else cast("Length | None", pPr.marL)

    @margin_left.setter
    def margin_left(self, value: Length | None) -> None:
        self._get_or_add_pPr().marL = value

    @property
    def space_after(self) -> Length | None:
        """Fixed spacing after paragraphs at this level; |None| when inherited or percentage."""
        pPr = self._get_pPr()
        return None if pPr is None else pPr.space_after

    @space_after.setter
    def space_after(self, value: Length | None) -> None:
        self._get_or_add_pPr().space_after = value

    @property
    def space_before(self) -> Length | None:
        """Fixed spacing before paragraphs at this level; |None| when inherited or percentage."""
        pPr = self._get_pPr()
        return None if pPr is None else pPr.space_before

    @space_before.setter
    def space_before(self, value: Length | None) -> None:
        self._get_or_add_pPr().space_before = value
