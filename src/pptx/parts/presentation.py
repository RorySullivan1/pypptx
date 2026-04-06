"""Presentation part, the main part in a .pptx package."""

from __future__ import annotations

from copy import deepcopy
from typing import IO, TYPE_CHECKING, Iterable

from pptx.exc import SlideError
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.parts.slide import NotesMasterPart, SlidePart
from pptx.presentation import Presentation
from pptx.util import lazyproperty

if TYPE_CHECKING:
    from pptx.oxml.xmlchemy import BaseOxmlElement
    from pptx.parts.coreprops import CorePropertiesPart
    from pptx.slide import NotesMaster, Slide, SlideLayout, SlideMaster


class PresentationPart(XmlPart):
    """Top level class in object model.

    Represents the contents of the /ppt directory of a .pptx file.
    """

    def add_slide(self, slide_layout: SlideLayout):
        """Return (rId, slide) pair of a newly created blank slide.

        New slide inherits appearance from `slide_layout`.
        """
        partname = self._next_slide_partname
        slide_layout_part = slide_layout.part
        slide_part = SlidePart.new(partname, self.package, slide_layout_part)
        rId = self.relate_to(slide_part, RT.SLIDE)
        return rId, slide_part.slide

    def duplicate_slide(self, slide_part: SlidePart) -> tuple[str, Slide]:
        """Return (rId, slide) pair for a newly created duplicate of `slide_part`.

        The new slide is a deep clone of the source slide with all relationships
        (images, charts, media, layout, etc.) re-established on the new part.
        Notes slide and comments are not carried over to the duplicate.
        """
        from pptx.opc.constants import CONTENT_TYPE as CT

        partname = self.package.next_partname("/ppt/slides/slide%d.xml")
        new_element = deepcopy(slide_part._element)
        new_slide_part = SlidePart(partname, CT.PML_SLIDE, self.package, new_element)

        # --- per-slide parts that should NOT be shared with the duplicate ---
        _skip_reltypes = {RT.NOTES_SLIDE, RT.COMMENTS, RT.TAGS}

        # --- build rId mapping from old slide to new slide ---
        rId_map: dict[str, str] = {}
        for rId, rel in slide_part.rels.items():
            if rel.reltype in _skip_reltypes:
                continue
            if rel.is_external:
                new_rId = new_slide_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            else:
                new_rId = new_slide_part.relate_to(rel.target_part, rel.reltype)
            rId_map[rId] = new_rId

        # --- update rId references in the cloned XML (r:id, r:embed, r:link) ---
        _remap_rIds(new_element, rId_map)

        rId = self.relate_to(new_slide_part, RT.SLIDE)
        return rId, new_slide_part.slide

    @property
    def core_properties(self) -> CorePropertiesPart:
        """A |CoreProperties| object for the presentation.

        Provides read/write access to the Dublin Core properties of this presentation.
        """
        return self.package.core_properties

    def get_slide(self, slide_id: int) -> Slide | None:
        """Return optional related |Slide| object identified by `slide_id`.

        Returns |None| if no slide with `slide_id` is related to this presentation.
        """
        for sldId in self._element.sldIdLst:
            if sldId.id == slide_id:
                return self.related_part(sldId.rId).slide
        return None

    @lazyproperty
    def notes_master(self) -> NotesMaster:
        """
        Return the |NotesMaster| object for this presentation. If the
        presentation does not have a notes master, one is created from
        a default template. The same single instance is returned on each
        call.
        """
        return self.notes_master_part.notes_master

    @lazyproperty
    def notes_master_part(self) -> NotesMasterPart:
        """Return the |NotesMasterPart| object for this presentation.

        If the presentation does not have a notes master, one is created from a default template.
        The same single instance is returned on each call.
        """
        try:
            return self.part_related_by(RT.NOTES_MASTER)
        except KeyError:
            notes_master_part = NotesMasterPart.create_default(self.package)
            self.relate_to(notes_master_part, RT.NOTES_MASTER)
            return notes_master_part

    @lazyproperty
    def presentation(self):
        """
        A |Presentation| object providing access to the content of this
        presentation.
        """
        return Presentation(self._element, self)

    def related_slide(self, rId: str) -> Slide:
        """Return |Slide| object for related |SlidePart| related by `rId`."""
        return self.related_part(rId).slide

    def related_slide_master(self, rId: str) -> SlideMaster:
        """Return |SlideMaster| object for |SlideMasterPart| related by `rId`."""
        return self.related_part(rId).slide_master

    def rename_slide_parts(self, rIds: Iterable[str]):
        """Assign incrementing partnames to the slide parts identified by `rIds`.

        Partnames are like `/ppt/slides/slide9.xml` and are assigned in the order their id appears
        in the `rIds` sequence. The name portion is always `slide`. The number part forms a
        continuous sequence starting at 1 (e.g. 1, 2, ... 10, ...). The extension is always
        `.xml`.
        """
        for idx, rId in enumerate(rIds):
            slide_part = self.related_part(rId)
            slide_part.partname = PackURI("/ppt/slides/slide%d.xml" % (idx + 1))

    def save(self, path_or_stream: str | IO[bytes]):
        """Save this presentation package to `path_or_stream`.

        `path_or_stream` can be either a path to a filesystem location (a string) or a
        file-like object.
        """
        self.package.save(path_or_stream)

    def slide_id(self, slide_part):
        """Return the slide-id associated with `slide_part`."""
        for sldId in self._element.sldIdLst:
            if self.related_part(sldId.rId) is slide_part:
                return sldId.id
        raise SlideError("matching slide_part not found")

    @property
    def _next_slide_partname(self):
        """Return |PackURI| instance containing next available slide partname."""
        sldIdLst = self._element.get_or_add_sldIdLst()
        partname_str = "/ppt/slides/slide%d.xml" % (len(sldIdLst) + 1)
        return PackURI(partname_str)


def _remap_rIds(element: BaseOxmlElement, rId_map: dict[str, str]) -> None:
    """Replace rId references in `element` XML tree per `rId_map`.

    Handles all relationship attribute forms: ``r:id``, ``r:embed``, and ``r:link``.
    """
    from pptx.oxml.ns import qn

    r_attrib_names = (qn("r:id"), qn("r:embed"), qn("r:link"))
    for descendant in element.iter():
        for attr_name in r_attrib_names:
            old_rId = descendant.get(attr_name)
            if old_rId is not None and old_rId in rId_map:
                new_rId = rId_map[old_rId]
                if old_rId != new_rId:
                    descendant.set(attr_name, new_rId)
