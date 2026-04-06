"""Main presentation object."""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, Iterator, cast

from pptx.shared import PartElementProxy
from pptx.slide import SlideMasters, Slides
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.presentation import CT_Presentation, CT_SlideId
    from pptx.oxml.section import CT_Section, CT_SectionList
    from pptx.parts.coreprops import CorePropertiesPart
    from pptx.parts.custprops import CustomPropertiesPart
    from pptx.parts.presentation import PresentationPart
    from pptx.slide import NotesMaster, SlideLayouts, SlideMaster
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
    def custom_properties(self) -> CustomPropertiesPart:
        """|CustomPropertiesPart| providing dict-like access to custom document properties.

        Supports ``custom_properties[name]``, ``custom_properties[name] = value``,
        ``del custom_properties[name]``, ``name in custom_properties``, and iteration.
        Values can be str, int, float, or bool.
        """
        return self.part.package.custom_properties

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
        """Add a slide (by its slide ID) to this section."""
        from lxml import etree

        from pptx.oxml.ns import qn

        sldId = etree.SubElement(self._section_elm, qn("p14:sldId"))
        sldId.set("id", str(slide_id))

    def remove(self) -> None:
        """Remove this section from the section list."""
        self._section_elm.getparent().remove(self._section_elm)
