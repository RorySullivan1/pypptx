"""Slide and related objects."""

from __future__ import annotations

from copy import deepcopy
from io import BytesIO
from typing import IO, TYPE_CHECKING, cast

from pptx.enum.shapes import PROG_ID
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.slide import CT_NotesMaster, CT_NotesSlide, CT_Slide, CT_SlideLayout
from pptx.oxml.theme import CT_OfficeStyleSheet
from pptx.parts.chart import ChartPart
from pptx.parts.comments import CommentsPart, ModernCommentsPart
from pptx.parts.embeddedpackage import EmbeddedPackagePart
from pptx.parts.tags import TagsPart
from pptx.slide import HandoutMaster, NotesMaster, NotesSlide, Slide, SlideLayout, SlideMaster
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.chart.data import CategoryChartData, ChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.media import Audio, Video
    from pptx.opc.package import Package
    from pptx.oxml.tags import CustDataLstOwnerMixin
    from pptx.parts.image import Image, ImagePart
    from pptx.parts.presentation import PresentationPart


class BaseSlidePart(XmlPart):
    """Base class for slide parts.

    This includes slide, slide-layout, and slide-master parts, but also notes-slide,
    notes-master, and handout-master parts.
    """

    _element: CT_Slide

    def get_image(self, rId: str) -> Image:
        """Return an |Image| object containing the image related to this slide by *rId*.

        Raises |KeyError| if no image is related by that id, which would generally indicate a
        corrupted .pptx file.
        """
        return cast("ImagePart", self.related_part(rId)).image

    def get_or_add_image_part(self, image_file: str | IO[bytes]) -> tuple[ImagePart, str]:
        """Return `(image_part, rId)` pair corresponding to `image_file`.

        The returned |ImagePart| object contains the image in `image_file` and is
        related to this slide with the key `rId`. If either the image part or
        relationship already exists, they are reused, otherwise they are newly created.
        """
        image_part = self._package.get_or_add_image_part(image_file)
        rId = self.relate_to(image_part, RT.IMAGE)
        return image_part, rId

    def get_or_add_tags_part(self, owner: CustDataLstOwnerMixin) -> TagsPart:
        """The |TagsPart| that `owner` refers to, created and related when it has none.

        `owner` is the `p:nvPr` of a shape in this part, or this part's `p:cSld`; either holds
        the `p:custDataLst/p:tags` reference to its own tags part, as PowerPoint writes it.
        """
        rId = owner.tags_rId
        if rId is not None:
            return cast("TagsPart", self.related_part(rId))
        tags_part = TagsPart.new(self._package)
        owner.set_tags_rId(self.relate_to(tags_part, RT.TAGS))
        return tags_part

    @property
    def name(self) -> str:
        """Internal name of this slide."""
        return self._element.cSld.name


class HandoutMasterPart(BaseSlidePart):
    """Handout master part, the layout of printed handouts.

    Corresponds to package file `ppt/handoutMasters/handoutMaster1.xml`.
    """

    @lazyproperty
    def handout_master(self) -> HandoutMaster:
        """The |HandoutMaster| object that proxies this handout master part."""
        return HandoutMaster(self._element, self)


class NotesMasterPart(BaseSlidePart):
    """Notes master part.

    Corresponds to package file `ppt/notesMasters/notesMaster1.xml`.
    """

    @classmethod
    def create_default(cls, package: Package) -> NotesMasterPart:
        """
        Create and return a default notes master part, including creating the
        new theme it requires.
        """
        notes_master_part = cls._new(package)
        theme_part = cls._new_theme_part(package)
        notes_master_part.relate_to(theme_part, RT.THEME)
        return notes_master_part

    @lazyproperty
    def notes_master(self) -> NotesMaster:
        """
        Return the |NotesMaster| object that proxies this notes master part.
        """
        return NotesMaster(self._element, self)

    @classmethod
    def _new(cls, package: Package) -> NotesMasterPart:
        """
        Create and return a standalone, default notes master part based on
        the built-in template (without any related parts, such as theme).
        """
        return NotesMasterPart(
            PackURI("/ppt/notesMasters/notesMaster1.xml"),
            CT.PML_NOTES_MASTER,
            package,
            CT_NotesMaster.new_default(),
        )

    @classmethod
    def _new_theme_part(cls, package: Package) -> XmlPart:
        """Return new default theme-part suitable for use with a notes master."""
        return XmlPart(
            package.next_partname("/ppt/theme/theme%d.xml"),
            CT.OFC_THEME,
            package,
            CT_OfficeStyleSheet.new_default(),
        )


class NotesSlidePart(BaseSlidePart):
    """Notes slide part.

    Contains the slide notes content and the layout for the slide handout page.
    Corresponds to package file `ppt/notesSlides/notesSlide[1-9][0-9]*.xml`.
    """

    @classmethod
    def new(cls, package: Package, slide_part: SlidePart) -> NotesSlidePart:
        """Return new |NotesSlidePart| for the slide in `slide_part`.

        The new notes-slide part is based on the (singleton) notes master and related to
        both the notes-master part and `slide_part`. If no notes-master is present,
        one is created based on the default template.
        """
        notes_master_part = package.presentation_part.notes_master_part
        notes_slide_part = cls._add_notes_slide_part(package, slide_part, notes_master_part)
        notes_slide = notes_slide_part.notes_slide
        notes_slide.clone_master_placeholders(notes_master_part.notes_master)
        return notes_slide_part

    @lazyproperty
    def notes_master(self) -> NotesMaster:
        """Return the |NotesMaster| object this notes slide inherits from."""
        notes_master_part = self.part_related_by(RT.NOTES_MASTER)
        return notes_master_part.notes_master

    @lazyproperty
    def notes_slide(self) -> NotesSlide:
        """Return the |NotesSlide| object that proxies this notes slide part."""
        return NotesSlide(self._element, self)

    @classmethod
    def _add_notes_slide_part(
        cls, package: Package, slide_part: SlidePart, notes_master_part: NotesMasterPart
    ) -> NotesSlidePart:
        """Create and return a new notes-slide part.

        The return part is fully related, but has no shape content (i.e. placeholders
        not cloned).
        """
        notes_slide_part = NotesSlidePart(
            package.next_partname("/ppt/notesSlides/notesSlide%d.xml"),
            CT.PML_NOTES_SLIDE,
            package,
            CT_NotesSlide.new(),
        )
        notes_slide_part.relate_to(notes_master_part, RT.NOTES_MASTER)
        notes_slide_part.relate_to(slide_part, RT.SLIDE)
        return notes_slide_part


class SlidePart(BaseSlidePart):
    """Slide part. Corresponds to package files ppt/slides/slide[1-9][0-9]*.xml."""

    @classmethod
    def new(cls, partname: PackURI, package: Package, slide_layout_part: SlideLayoutPart) -> SlidePart:
        """Return newly-created blank slide part.

        The new slide-part has `partname` and a relationship to `slide_layout_part`.
        """
        slide_part = cls(partname, CT.PML_SLIDE, package, CT_Slide.new())
        slide_part.relate_to(slide_layout_part, RT.SLIDE_LAYOUT)
        return slide_part

    def add_chart_part(self, chart_type: XL_CHART_TYPE, chart_data: ChartData) -> str:
        """Return str rId of new |ChartPart| object containing chart of `chart_type`.

        The chart depicts `chart_data` and is related to the slide contained in this
        part by `rId`.
        """
        return self.relate_to(ChartPart.new(chart_type, chart_data, self._package), RT.CHART)

    def add_chartex_part(
        self, chart_type: XL_CHART_TYPE, chart_data: CategoryChartData
    ) -> str:
        """Return str rId of a new |ChartExPart| containing a chartex chart of `chart_type`.

        The chart depicts `chart_data` and is related to the slide contained in this part by
        `rId`.
        """
        from pptx.parts.chartex import ChartExPart

        return self.relate_to(ChartExPart.new(chart_type, chart_data, self._package), RT.CHART_EX)

    def add_embedded_ole_object_part(
        self, prog_id: PROG_ID | str, ole_object_file: str | IO[bytes]
    ) -> str:
        """Return rId of newly-added OLE-object part formed from `ole_object_file`."""
        relationship_type = RT.PACKAGE if isinstance(prog_id, PROG_ID) else RT.OLE_OBJECT
        return self.relate_to(
            EmbeddedPackagePart.factory(
                prog_id, self._blob_from_file(ole_object_file), self._package
            ),
            relationship_type,
        )

    def get_or_add_video_media_part(self, video: Video) -> tuple[str, str]:
        """Return rIds for media and video relationships to media part.

        A new |MediaPart| object is created if it does not already exist
        (such as would occur if the same video appeared more than once in
         a presentation). Two relationships to the media part are created,
        one each with MEDIA and VIDEO relationship types. The need for two
        appears to be for legacy support for an earlier (pre-Office 2010)
        PowerPoint media embedding strategy.
        """
        media_part = self._package.get_or_add_media_part(video)
        media_rId = self.relate_to(media_part, RT.MEDIA)
        video_rId = self.relate_to(media_part, RT.VIDEO)
        return media_rId, video_rId

    def get_or_add_audio_media_part(self, audio: Audio) -> tuple[str, str]:
        """Return rIds for media and audio relationships to media part.

        A new |MediaPart| object is created if it does not already exist (such as would
        occur if the same audio clip appeared more than once in a presentation). Two
        relationships to the media part are created, one each with MEDIA and AUDIO
        relationship types, mirroring the legacy dual-relationship strategy used for video.
        """
        media_part = self._package.get_or_add_media_part(audio)
        media_rId = self.relate_to(media_part, RT.MEDIA)
        audio_rId = self.relate_to(media_part, RT.AUDIO)
        return media_rId, audio_rId

    @property
    def has_notes_slide(self) -> bool:
        """
        Return True if this slide has a notes slide, False otherwise. A notes
        slide is created by the :attr:`notes_slide` property when one doesn't
        exist; use this property to test for a notes slide without the
        possible side-effect of creating one.
        """
        try:
            self.part_related_by(RT.NOTES_SLIDE)
        except KeyError:
            return False
        return True

    @property
    def comments_part(self) -> CommentsPart:
        """The |CommentsPart| for this slide, creating one if not present."""
        try:
            return cast("CommentsPart", self.part_related_by(RT.COMMENTS))
        except KeyError:
            # Derive index from slide partname (e.g. slide1.xml → comment1.xml)
            idx = self.partname.idx
            partname = "/ppt/comments/comment%d.xml" % idx
            comments_part = CommentsPart.default(self._package, partname)
            self.relate_to(comments_part, RT.COMMENTS)
            return comments_part

    @property
    def modern_comments_part(self) -> ModernCommentsPart | None:
        """The |ModernCommentsPart| holding this slide's comment threads, or None if absent."""
        try:
            return cast(ModernCommentsPart, self.part_related_by(RT.MODERN_COMMENTS))
        except KeyError:
            return None

    def get_or_add_modern_comments_part(self) -> ModernCommentsPart:
        """The |ModernCommentsPart| for this slide, created and related when not present.

        A new part is named after the slide as PowerPoint names it,
        ``modernComment_<slide-id>_<creation-id>.xml`` in hex, and the slide's
        `p188:commentRel` extension is pointed at the new relationship.
        """
        comments_part = self.modern_comments_part
        if comments_part is not None:
            return comments_part

        creation_id = self._element.cSld.get_or_add_creation_id()
        partname = PackURI(
            "/ppt/comments/modernComment_%X_%X.xml" % (self.slide_id, creation_id)
        )
        if partname in {part.partname for part in self._package.iter_parts()}:
            partname = self._package.next_partname("/ppt/comments/modernComment_%d.xml")
        comments_part = ModernCommentsPart.new(partname, self._package)
        self._element.set_comment_rel(self.relate_to(comments_part, RT.MODERN_COMMENTS))
        return comments_part

    @property
    def has_comments(self) -> bool:
        """True if this slide has a comments part."""
        try:
            self.part_related_by(RT.COMMENTS)
            return True
        except KeyError:
            return False

    @property
    def tags_part(self) -> TagsPart:
        """The |TagsPart| for this slide, creating one if not present.

        It is the part `p:cSld/p:custDataLst/p:tags` refers to. The slide part also relates the
        tags part of each tagged shape, so a tags relationship alone does not identify the
        slide's; one no shape refers to (as pypptx wrote before shape tags) is adopted.
        """
        cSld = self._element.cSld
        if cSld.tags_rId is None:
            shape_rIds = set(cSld.spTree.xpath(".//p:custDataLst/p:tags/@r:id"))
            for rId, rel in self.rels.items():
                if rel.reltype == RT.TAGS and not rel.is_external and rId not in shape_rIds:
                    cSld.set_tags_rId(rId)
                    break
        return self.get_or_add_tags_part(cSld)

    @lazyproperty
    def notes_slide(self) -> NotesSlide:
        """The |NotesSlide| instance associated with this slide.

        If the slide does not have a notes slide, a new one is created. The same single instance
        is returned on each call.
        """
        try:
            notes_slide_part = self.part_related_by(RT.NOTES_SLIDE)
        except KeyError:
            notes_slide_part = self._add_notes_slide_part()
        return notes_slide_part.notes_slide

    @lazyproperty
    def slide(self) -> Slide:
        """
        The |Slide| object representing this slide part.
        """
        return Slide(self._element, self)

    @property
    def slide_id(self) -> int:
        """Return the slide identifier stored in the presentation part for this slide part."""
        presentation_part = self.package.presentation_part
        return presentation_part.slide_id(self)

    @property
    def slide_layout(self) -> SlideLayout:
        """|SlideLayout| object the slide in this part inherits appearance from."""
        slide_layout_part = self.part_related_by(RT.SLIDE_LAYOUT)
        return slide_layout_part.slide_layout

    def _add_notes_slide_part(self) -> NotesSlidePart:
        """
        Return a newly created |NotesSlidePart| object related to this slide
        part. Caller is responsible for ensuring this slide doesn't already
        have a notes slide part.
        """
        notes_slide_part = NotesSlidePart.new(self.package, self)
        self.relate_to(notes_slide_part, RT.NOTES_SLIDE)
        return notes_slide_part


class SlideLayoutPart(BaseSlidePart):
    """Slide layout part.

    Corresponds to package files ``ppt/slideLayouts/slideLayout[1-9][0-9]*.xml``.
    """

    @lazyproperty
    def slide_layout(self) -> SlideLayout:
        """
        The |SlideLayout| object representing this part.
        """
        return SlideLayout(self._element, self)

    @property
    def slide_master(self) -> SlideMaster:
        """Slide master from which this slide layout inherits properties."""
        return self.part_related_by(RT.SLIDE_MASTER).slide_master


class SlideMasterPart(BaseSlidePart):
    """Slide master part.

    Corresponds to package files ppt/slideMasters/slideMaster[1-9][0-9]*.xml.
    """

    def related_slide_layout(self, rId: str) -> SlideLayout:
        """Return |SlideLayout| related to this slide-master by key `rId`."""
        return self.related_part(rId).slide_layout

    def add_slide_layout_part(self, name: str) -> SlideLayoutPart:
        """Add and return a new, empty slide-layout part named `name` under this master.

        The layout is related to this master and back, and listed last in its
        `p:sldLayoutIdLst` with a fresh id.
        """
        package = self.package
        layout_part = SlideLayoutPart(
            package.next_partname("/ppt/slideLayouts/slideLayout%d.xml"),
            CT.PML_SLIDE_LAYOUT,
            package,
            CT_SlideLayout.new(name),
        )
        layout_part.relate_to(self, RT.SLIDE_MASTER)
        self._list_layout(layout_part)
        return layout_part

    def duplicate_slide_layout_part(
        self, source_part: SlideLayoutPart, name: str
    ) -> SlideLayoutPart:
        """Add and return a copy of `source_part`, one of this master's layouts, named `name`.

        The copy is listed directly after its source. Like a duplicated slide it shares the
        source's images and media, and gets its own copy of each chart and tags part.
        """
        from pptx.parts.presentation import _PartCopier, _remap_rIds

        package = self.package
        new_element = deepcopy(source_part._element)
        layout_part = SlideLayoutPart(
            package.next_partname("/ppt/slideLayouts/slideLayout%d.xml"),
            CT.PML_SLIDE_LAYOUT,
            package,
            new_element,
        )
        copier = _PartCopier(package, dedup=False)
        rId_map: dict[str, str] = {}
        for rId, rel in source_part.rels.items():
            if rel.is_external:
                new_rId = layout_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            elif rel.reltype in (RT.CHART, RT.TAGS):
                new_rId = layout_part.relate_to(copier.copy(rel.target_part), rel.reltype)
            else:
                new_rId = layout_part.relate_to(rel.target_part, rel.reltype)
            rId_map[rId] = new_rId
        _remap_rIds(new_element, rId_map)
        new_element.cSld.name = name
        new_element.cSld.renew_creation_id()

        source_entry = next(
            entry
            for entry in self._element.get_or_add_sldLayoutIdLst().sldLayoutId_lst
            if self.related_part(entry.rId) is source_part
        )
        self._list_layout(layout_part, after=source_entry)
        return layout_part

    @lazyproperty
    def slide_master(self) -> SlideMaster:
        """
        The |SlideMaster| object representing this part.
        """
        return SlideMaster(self._element, self)

    def _list_layout(self, layout_part: SlideLayoutPart, after=None) -> None:
        """Relate `layout_part` and list it in `p:sldLayoutIdLst`, last or after `after`."""
        presentation_part = cast("PresentationPart", self.package.presentation_part)
        rId = self.relate_to(layout_part, RT.SLIDE_LAYOUT)
        self._element.get_or_add_sldLayoutIdLst().add_sldLayoutId(
            rId, id=presentation_part.next_master_or_layout_id(), after=after
        )

    @property
    def theme_part(self) -> XmlPart | None:
        """The |XmlPart| containing the theme for this slide master.

        Returns None if no theme relationship exists.
        """
        try:
            return self.part_related_by(RT.THEME)
        except KeyError:
            return None

    def apply_theme(self, source_theme_part: XmlPart) -> XmlPart:
        """Relate this master to a new theme part copied from `source_theme_part`.

        `source_theme_part` can belong to another package. Parts it relates to, such as images
        used by fill styles, are copied too; slide-master and other structural relationships of
        a `.thmx` theme are not. The previous theme part is dropped from the package unless
        something else still relates to it. The presentation part's theme relationship follows
        this master when no other master uses the previous theme. Returns the new theme part.
        """
        from pptx.parts.image import ImagePart
        from pptx.parts.presentation import _STRUCTURAL_RELTYPES, _PartCopier, _remap_rIds

        package = self.package
        old_theme_part = self.theme_part
        for rId in [rId for rId, rel in self.rels.items() if rel.reltype == RT.THEME]:
            self._rels.pop(rId)

        new_theme_part = XmlPart(
            package.next_partname("/ppt/theme/theme%d.xml"),
            CT.OFC_THEME,
            package,
            deepcopy(source_theme_part._element),
        )
        self.relate_to(new_theme_part, RT.THEME)

        copier = _PartCopier(package, dedup=True)
        rId_map: dict[str, str] = {}
        for rId, rel in source_theme_part.rels.items():
            if rel.is_external:
                rId_map[rId] = new_theme_part.relate_to(
                    rel.target_ref, rel.reltype, is_external=True
                )
            elif rel.reltype in _STRUCTURAL_RELTYPES:
                continue
            elif isinstance(rel.target_part, ImagePart):
                image_part = package.get_or_add_image_part(BytesIO(rel.target_part.blob))
                rId_map[rId] = new_theme_part.relate_to(image_part, rel.reltype)
            else:
                rId_map[rId] = new_theme_part.relate_to(copier.copy(rel.target_part), rel.reltype)
        _remap_rIds(new_theme_part._element, rId_map)

        if old_theme_part is not None:
            self._repoint_presentation_theme(old_theme_part, new_theme_part)
        return new_theme_part

    def _repoint_presentation_theme(self, old_theme_part: XmlPart, new_theme_part: XmlPart):
        """Move the presentation part's theme relationship from the old theme to the new one.

        Only done when no other slide master still uses `old_theme_part`.
        """
        presentation_part = cast("PresentationPart", self.package.presentation_part)
        for rel in presentation_part.rels.values():
            if rel.reltype != RT.SLIDE_MASTER or rel.target_part is self:
                continue
            if getattr(rel.target_part, "theme_part", None) is old_theme_part:
                return
        for rId, rel in list(presentation_part.rels.items()):
            if rel.reltype == RT.THEME and rel.target_part is old_theme_part:
                presentation_part._rels.pop(rId)
                presentation_part.relate_to(new_theme_part, RT.THEME)
