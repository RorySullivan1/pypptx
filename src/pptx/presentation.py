"""Main presentation object."""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, Iterator, Mapping, cast

from pptx.custom_show import CustomShows
from pptx.enum.pres import PP_SLIDE_SHOW_TYPE
from pptx.exc import InvalidValueError
from pptx.shared import PartElementProxy
from pptx.slide import SlideMasters, Slides
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.fonts import EmbeddedFonts
    from pptx.dml.color import RGBColor
    from pptx.oxml.dml.color import CT_SRgbColor
    from pptx.oxml.presentation import CT_Presentation, CT_SlideId
    from pptx.oxml.presprops import CT_ShowProperties
    from pptx.oxml.section import CT_Section, CT_SectionList
    from pptx.oxml.xmlchemy import BaseOxmlElement
    from pptx.parts.coreprops import CorePropertiesPart
    from pptx.parts.custprops import CustomPropertiesPart
    from pptx.parts.presentation import PresentationPart
    from pptx.parts.tablestyles import TableStylesPart
    from pptx.slide import HandoutMaster, NotesMaster, SlideLayouts, SlideMaster
    from pptx.text.styles import TextListStyle
    from pptx.theme import Theme
    from pptx.util import Length


class Presentation(PartElementProxy):
    """PresentationML (PML) presentation.

    Not intended to be constructed directly. Use :func:`pptx.Presentation` to open or
    create a presentation.
    """

    _element: CT_Presentation
    part: PresentationPart  # pyright: ignore[reportIncompatibleMethodOverride]

    @property
    def first_slide_number(self) -> int:
        """Starting slide number for this presentation.

        Read/write. Defaults to 1 when not explicitly set.
        """
        val = self._element.firstSlideNum
        return val if val is not None else 1

    @first_slide_number.setter
    def first_slide_number(self, value: int) -> None:
        self._element.firstSlideNum = value

    @property
    def core_properties(self) -> CorePropertiesPart:
        """|CoreProperties| instance for this presentation.

        Provides read/write access to the Dublin Core document properties for the presentation.
        """
        return self.part.core_properties

    @property
    def embedded_fonts(self) -> EmbeddedFonts:
        """|EmbeddedFonts| collection of fonts embedded in this presentation.

        Supports `len()`, iteration, and indexed access; see `EmbeddedFont.remove()` (or
        `EmbeddedFonts.remove()`) to remove one, e.g. to shrink the file or drop a
        licensed font. Embedding new fonts is not supported.
        """
        from pptx.fonts import EmbeddedFonts

        return EmbeddedFonts(self._element, self.part)

    @property
    def custom_properties(self) -> CustomPropertiesPart:
        """|CustomPropertiesPart| providing dict-like access to custom document properties.

        Supports ``custom_properties[name]``, ``custom_properties[name] = value``,
        ``del custom_properties[name]``, ``name in custom_properties``, and iteration.
        Values can be str, int, float, or bool.
        """
        return self.part.package.custom_properties

    @property
    def handout_master(self) -> HandoutMaster | None:
        """The |HandoutMaster| of this presentation, or |None| when it has none. Read only.

        Unlike `notes_master`, this never creates a handout master.
        """
        handout_master_part = self.part.handout_master_part
        return None if handout_master_part is None else handout_master_part.handout_master

    @property
    def notes_master(self) -> NotesMaster:
        """Instance of |NotesMaster| for this presentation.

        If the presentation does not have a notes master, one is created from a default template
        and returned. The same single instance is returned on each call.
        """
        return self.part.notes_master

    def save(self, file: str | IO[bytes]) -> None:
        """Writes this presentation to `file`.

        `file` can be either a file-path or a file-like object open for writing bytes.
        """
        self.part.save(file)

    @property
    def slide_height(self) -> Length | None:
        """Height of slides in this presentation, in English Metric Units (EMU).

        Returns |None| if no slide width is defined. Read/write.
        """
        sldSz = self._element.sldSz
        if sldSz is None:
            return None
        return sldSz.cy

    @slide_height.setter
    def slide_height(self, height: Length) -> None:
        sldSz = self._element.get_or_add_sldSz()
        sldSz.cy = height

    @property
    def notes_height(self) -> Length | None:
        """Height of notes pages in this presentation, in English Metric Units (EMU).

        Returns |None| if no notes size is defined. Read/write.
        """
        notesSz = self._element.notesSz
        if notesSz is None:
            return None
        return notesSz.cy

    @notes_height.setter
    def notes_height(self, height: Length) -> None:
        notesSz = self._element.get_or_add_notesSz()
        notesSz.cy = height

    @property
    def notes_width(self) -> Length | None:
        """Width of notes pages in this presentation, in English Metric Units (EMU).

        Returns |None| if no notes size is defined. Read/write.
        """
        notesSz = self._element.notesSz
        if notesSz is None:
            return None
        return notesSz.cx

    @notes_width.setter
    def notes_width(self, width: Length) -> None:
        notesSz = self._element.get_or_add_notesSz()
        notesSz.cx = width

    @property
    def slide_layouts(self) -> SlideLayouts:
        """|SlideLayouts| collection belonging to the first |SlideMaster| of this presentation.

        A presentation can have more than one slide master and each master will have its own set
        of layouts. This property is a convenience for the common case where the presentation has
        only a single slide master.
        """
        return self.slide_masters[0].slide_layouts

    @property
    def slide_master(self) -> SlideMaster:
        """
        First |SlideMaster| object belonging to this presentation. Typically,
        presentations have only a single slide master. This property provides
        simpler access in that common case.
        """
        return self.slide_masters[0]

    @property
    def default_text_style(self) -> TextListStyle:
        """|TextListStyle| holding the presentation-wide text defaults (`p:defaultTextStyle`).

        These are the lowest-priority text defaults, applying where neither the text itself nor
        a master text style sets a value.
        """
        from pptx.text.styles import TextListStyle

        prs = self._element
        return TextListStyle(
            lambda: prs.defaultTextStyle, lambda: prs.get_or_add_defaultTextStyle()
        )

    @property
    def theme(self) -> Theme | None:
        """A |Theme| object for this presentation's first slide master theme.

        Returns None if no theme is available. This is a convenience for
        the common case of a single slide master.
        """
        return self.slide_master.theme

    @lazyproperty
    def slide_masters(self) -> SlideMasters:
        """|SlideMasters| collection of slide-masters belonging to this presentation."""
        return SlideMasters(self._element.get_or_add_sldMasterIdLst(), self)

    @property
    def slide_width(self) -> Length | None:
        """
        Width of slides in this presentation, in English Metric Units (EMU).
        Returns |None| if no slide width is defined. Read/write.
        """
        sldSz = self._element.sldSz
        if sldSz is None:
            return None
        return sldSz.cx

    @slide_width.setter
    def slide_width(self, width: Length) -> None:
        sldSz = self._element.get_or_add_sldSz()
        sldSz.cx = width

    @property
    def sections(self) -> Sections:
        """|Sections| object providing access to named slide groups."""
        return Sections(self._element)

    @property
    def custom_shows(self) -> CustomShows:
        """|CustomShows| collection of named, ordered slide subsets in this presentation."""
        return CustomShows(self._element, self)

    @property
    def slide_show_settings(self) -> SlideShowSettings:
        """|SlideShowSettings| object for this presentation's `ppt/presProps.xml` part.

        Reading these settings never creates `presProps.xml` if it isn't already present
        (defaults are returned instead); writing any setting creates the part (and its
        relationship from the presentation part) on first use.
        """
        return SlideShowSettings(self.part)

    @property
    def table_styles(self) -> TableStyles:
        """|TableStyles| read-only mapping of table-style id to display name.

        Reflects the table styles defined in `ppt/tableStyles.xml`. Empty when the
        presentation has no `tableStyles.xml` part.
        """
        return TableStyles(self.part.table_styles_part)

    @lazyproperty
    def slides(self) -> Slides:
        """|Slides| object containing the slides in this presentation."""
        sldIdLst = self._element.get_or_add_sldIdLst()
        self.part.rename_slide_parts([cast("CT_SlideId", sldId).rId for sldId in sldIdLst])
        return Slides(sldIdLst, self)


class Sections:
    """Sequence of |Section| objects representing named slide groups.

    Supports iteration, indexed access, and len(). Sections are defined in the
    presentation extensions as ``p14:sectionLst``.
    """

    def __init__(self, prs_elm: CT_Presentation):
        self._prs_elm = prs_elm

    def __getitem__(self, idx: int) -> Section:
        sections = self._section_elms
        if idx < 0 or idx >= len(sections):
            raise IndexError("section index out of range")
        return Section(sections[idx])

    def __iter__(self) -> Iterator[Section]:
        for section_elm in self._section_elms:
            yield Section(section_elm)

    def __len__(self) -> int:
        return len(self._section_elms)

    def add(self, name: str) -> Section:
        """Add a new section with the given `name` and return it."""
        from lxml import etree

        from pptx.oxml.ns import qn

        sectionLst = self._prs_elm.get_or_add_sectionLst()
        section = etree.SubElement(sectionLst, qn("p14:section"))
        section.set("name", name)
        return Section(section)

    @property
    def _section_elms(self) -> list[CT_Section]:
        sectionLst = self._prs_elm.sectionLst
        if sectionLst is None:
            return []
        return sectionLst.section_lst


class Section:
    """A named group of slides in the presentation."""

    def __init__(self, section_elm: CT_Section):
        self._section_elm = section_elm

    @property
    def name(self) -> str | None:
        """Name of this section. Read/write."""
        return self._section_elm.name

    @name.setter
    def name(self, value: str) -> None:
        self._section_elm.name = value

    @property
    def slide_ids(self) -> tuple[int, ...]:
        """Tuple of slide IDs belonging to this section."""
        return tuple(entry.id for entry in self._section_elm.sldId_lst)

    def add_slide_id(self, slide_id: int) -> None:
        """Add the slide whose ID is `slide_id` to the end of this section (`p14:sldId`)."""
        from lxml import etree

        from pptx.oxml.ns import qn

        sldId = etree.SubElement(self._section_elm, qn("p14:sldId"))
        sldId.set("id", str(slide_id))

    def remove(self) -> None:
        """Remove this section from the section list."""
        self._section_elm.getparent().remove(self._section_elm)


class SlideShowSettings:
    """Read/write access to slide-show settings, `p:showPr` in `ppt/presProps.xml`.

    Obtained via `Presentation.slide_show_settings`. Not intended to be constructed
    directly.

    Reading a property never creates `presProps.xml` or `p:showPr`; absent settings
    read back as their ECMA-376 default value. Writing any property creates
    `presProps.xml` (and its relationship from the presentation part) and `p:showPr`
    on first use, if not already present.

    The slide range shown is exposed as two mutually-exclusive properties mirroring
    the `p:sldRg` / `p:custShow` choice: `slide_range` (a 1-based `(start, end)`
    tuple) and `custom_show_id`. Setting one clears the other (and `p:sldAll`).
    Setting both to `None` reverts to showing all slides.
    """

    def __init__(self, presentation_part: PresentationPart) -> None:
        self._presentation_part = presentation_part

    @property
    def loop(self) -> bool:
        """`True` if the show restarts after the last slide instead of ending. Read/write."""
        showPr = self._showPr
        return showPr.loop if showPr is not None else False

    @loop.setter
    def loop(self, value: bool) -> None:
        self._get_or_add_showPr().loop = bool(value)

    @property
    def show_narration(self) -> bool:
        """`True` if slide narrations are played during the show. Read/write."""
        showPr = self._showPr
        return showPr.showNarration if showPr is not None else False

    @show_narration.setter
    def show_narration(self, value: bool) -> None:
        self._get_or_add_showPr().showNarration = bool(value)

    @property
    def show_animation(self) -> bool:
        """`True` if animations are played during the show. Read/write."""
        showPr = self._showPr
        return showPr.showAnimation if showPr is not None else True

    @show_animation.setter
    def show_animation(self, value: bool) -> None:
        self._get_or_add_showPr().showAnimation = bool(value)

    @property
    def use_timings(self) -> bool:
        """`True` if slide/animation timings are used to advance the show. Read/write."""
        showPr = self._showPr
        return showPr.useTimings if showPr is not None else True

    @use_timings.setter
    def use_timings(self, value: bool) -> None:
        self._get_or_add_showPr().useTimings = bool(value)

    @property
    def show_type(self) -> PP_SLIDE_SHOW_TYPE:
        """Member of :ref:`PpSlideShowType` specifying how the show is displayed. Read/write.

        Defaults to `PP_SLIDE_SHOW_TYPE.SPEAKER` (`p:present`, PowerPoint's default) when
        not explicitly set.
        """
        showPr = self._showPr
        choice = showPr.showTypeChoice if showPr is not None else None
        if choice is None:
            return PP_SLIDE_SHOW_TYPE.SPEAKER
        return _SHOW_TYPE_TAG_MAP[cast("BaseOxmlElement", choice).tag]

    @show_type.setter
    def show_type(self, value: PP_SLIDE_SHOW_TYPE) -> None:
        showPr = self._get_or_add_showPr()
        if value == PP_SLIDE_SHOW_TYPE.SPEAKER:
            showPr.get_or_change_to_present()
        elif value == PP_SLIDE_SHOW_TYPE.BROWSE:
            showPr.get_or_change_to_browse()
        elif value == PP_SLIDE_SHOW_TYPE.KIOSK:
            showPr.get_or_change_to_kiosk()
        else:  # pragma: no cover - defensive, PP_SLIDE_SHOW_TYPE has only 3 members
            raise InvalidValueError("unknown PP_SLIDE_SHOW_TYPE member: %r" % (value,))

    @property
    def pen_color(self) -> RGBColor | None:
        """|RGBColor| of the presenter's pen/laser, or None if not set. Read/write.

        Only the RGB (`a:srgbClr`) color form is supported for read and write; a
        `p:penClr` set to a theme or system color reads back as `None`.
        """
        from pptx.dml.color import RGBColor
        from pptx.oxml.ns import qn

        showPr = self._showPr
        penClr = showPr.penClr if showPr is not None else None
        if penClr is None:
            return None
        xClr = cast("BaseOxmlElement | None", penClr.eg_colorChoice)
        if xClr is None or xClr.tag != qn("a:srgbClr"):
            return None
        return RGBColor.from_string(cast("CT_SRgbColor", xClr).val)

    @pen_color.setter
    def pen_color(self, value: RGBColor | None) -> None:
        from pptx.oxml.ns import qn

        if value is None:
            showPr = self._showPr
            if showPr is not None:
                showPr._remove_penClr()  # pyright: ignore[reportPrivateUsage]
            return
        penClr = self._get_or_add_showPr().get_or_add_penClr()
        xClr = cast("BaseOxmlElement | None", penClr.eg_colorChoice)
        srgbClr = (
            cast("CT_SRgbColor", xClr)
            if xClr is not None and xClr.tag == qn("a:srgbClr")
            else penClr.get_or_change_to_srgbClr()
        )
        srgbClr.val = str(value)  # pyright: ignore[reportAttributeAccessIssue]

    @property
    def slide_range(self) -> tuple[int, int] | None:
        """1-based `(start, end)` inclusive slide range shown, or None. Read/write.

        None means either "all slides" (the default) or that a custom show is
        selected instead -- see `custom_show_id`.
        """
        showPr = self._showPr
        sldRg = showPr.sldRg if showPr is not None else None
        if sldRg is None:
            return None
        return (sldRg.st, sldRg.end)

    @slide_range.setter
    def slide_range(self, value: tuple[int, int] | None) -> None:
        if value is None:
            showPr = self._showPr
            if showPr is not None:
                showPr._remove_slideRangeChoice()  # pyright: ignore[reportPrivateUsage]
            return
        start, end = value
        sldRg = self._get_or_add_showPr().get_or_change_to_sldRg()
        sldRg.st = start
        sldRg.end = end

    @property
    def custom_show_id(self) -> int | None:
        """Id of the `p:custShow` selected to be shown, or None. Read/write.

        None means either "all slides" or that a `slide_range` is selected instead.
        """
        showPr = self._showPr
        custShow = showPr.custShow if showPr is not None else None
        return custShow.id if custShow is not None else None

    @custom_show_id.setter
    def custom_show_id(self, value: int | None) -> None:
        if value is None:
            showPr = self._showPr
            if showPr is not None:
                showPr._remove_slideRangeChoice()  # pyright: ignore[reportPrivateUsage]
            return
        custShow = self._get_or_add_showPr().get_or_change_to_custShow()
        custShow.id = value

    @property
    def _showPr(self) -> CT_ShowProperties | None:
        """The `p:showPr` element, or None if `presProps.xml` or `p:showPr` is absent."""
        pres_props_part = self._presentation_part.pres_props_part
        if pres_props_part is None:
            return None
        return pres_props_part._element.showPr  # pyright: ignore[reportPrivateUsage]

    def _get_or_add_showPr(self) -> CT_ShowProperties:
        """The `p:showPr` element, creating `presProps.xml` and/or `p:showPr` as needed."""
        pres_props_part = self._presentation_part.get_or_add_pres_props_part()
        return pres_props_part._element.get_or_add_showPr()  # pyright: ignore[reportPrivateUsage]


class TableStyles(Mapping[str, str]):
    """Read-only mapping of table-style id (a GUID) to display name.

    Obtained via `Presentation.table_styles`. Reflects the styles defined in
    `ppt/tableStyles.xml`; empty when the presentation has no such part. The id of
    the style used by default for new tables is available as `default_id`.
    """

    def __init__(self, table_styles_part: TableStylesPart | None) -> None:
        self._table_styles_part = table_styles_part

    def __getitem__(self, key: str) -> str:
        for styleId, styleName in self._items:
            if styleId == key:
                return styleName
        raise KeyError(key)

    def __iter__(self) -> Iterator[str]:
        for styleId, _ in self._items:
            yield styleId

    def __len__(self) -> int:
        return len(self._items)

    @property
    def default_id(self) -> str | None:
        """Id (GUID) of the table style PowerPoint applies to a new table by default.

        None when the presentation has no `tableStyles.xml` part, or that part
        defines no default.
        """
        part = self._table_styles_part
        if part is None:
            return None
        return part._element.def_  # pyright: ignore[reportPrivateUsage]

    @property
    def _items(self) -> list[tuple[str, str]]:
        part = self._table_styles_part
        if part is None:
            return []
        return [(s.styleId, s.styleName) for s in part._element.tblStyle_lst]


def _show_type_tag_map() -> dict[str, PP_SLIDE_SHOW_TYPE]:
    from pptx.oxml.ns import qn

    return {
        qn("p:present"): PP_SLIDE_SHOW_TYPE.SPEAKER,
        qn("p:browse"): PP_SLIDE_SHOW_TYPE.BROWSE,
        qn("p:kiosk"): PP_SLIDE_SHOW_TYPE.KIOSK,
    }


_SHOW_TYPE_TAG_MAP: dict[str, PP_SLIDE_SHOW_TYPE] = _show_type_tag_map()
