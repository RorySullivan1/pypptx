"""Presentation part, the main part in a .pptx package."""

from __future__ import annotations

from copy import deepcopy
from typing import IO, TYPE_CHECKING, Iterable

from pptx.exc import SlideError
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part, XmlPart
from pptx.opc.packuri import PackURI
from pptx.parts.slide import NotesMasterPart, SlideLayoutPart, SlideMasterPart, SlidePart
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
        _skip_reltypes = {RT.NOTES_SLIDE, RT.COMMENTS, RT.MODERN_COMMENTS, RT.TAGS}

        # --- a chart holds editable data, so each slide gets its own copy; images and media
        # --- are immutable blobs and stay shared ---
        copier = _PartCopier(self.package, dedup=False)

        # --- build rId mapping from old slide to new slide ---
        rId_map: dict[str, str] = {}
        for rId, rel in slide_part.rels.items():
            if rel.reltype in _skip_reltypes:
                continue
            if rel.is_external:
                new_rId = new_slide_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            elif rel.reltype == RT.CHART:
                new_rId = new_slide_part.relate_to(copier.copy(rel.target_part), rel.reltype)
            else:
                new_rId = new_slide_part.relate_to(rel.target_part, rel.reltype)
            rId_map[rId] = new_rId

        # --- update rId references in the cloned XML (r:id, r:embed, r:link) ---
        _remap_rIds(new_element, rId_map)
        # --- comments aren't carried over, so drop the extension that points at them ---
        new_element.remove_comment_rel()

        rId = self.relate_to(new_slide_part, RT.SLIDE)
        return rId, new_slide_part.slide

    def import_slide(self, source_slide_part: SlidePart) -> tuple[str, Slide]:
        """Return (rId, slide) pair for a slide imported from another presentation.

        The source slide's XML is deep-copied and a new SlidePart is created in this
        package. The source slide's layout is matched by name in this presentation;
        if no match is found, the layout (and its master/theme if needed) are imported.
        All relationships (images, charts, media) are re-established in this package.
        Notes, comments, and tags are not imported.
        """
        from pptx.opc.constants import CONTENT_TYPE as CT

        partname = self.package.next_partname("/ppt/slides/slide%d.xml")
        new_element = deepcopy(source_slide_part._element)
        new_slide_part = SlidePart(partname, CT.PML_SLIDE, self.package, new_element)

        # --- find or import the slide layout ---
        source_layout_part = source_slide_part.part_related_by(RT.SLIDE_LAYOUT)
        target_layout_part = self._find_or_import_layout(source_layout_part)

        # --- per-slide parts that should NOT be imported ---
        _skip_reltypes = {
            RT.NOTES_SLIDE,
            RT.COMMENTS,
            RT.MODERN_COMMENTS,
            RT.TAGS,
            RT.SLIDE_LAYOUT,
        }

        # --- content parts (charts, images, media, OLE) come with their own relationships, e.g.
        # --- a chart's embedded workbook; structural parts keep the flat import ---
        copier = _PartCopier(self.package, dedup=True)

        # --- build rId mapping, re-establishing relationships in target package ---
        rId_map: dict[str, str] = {}
        for rId, rel in source_slide_part.rels.items():
            if rel.reltype in _skip_reltypes:
                continue
            if rel.is_external:
                new_rId = new_slide_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            elif rel.reltype in _STRUCTURAL_RELTYPES:
                new_rId = new_slide_part.relate_to(self._import_part(rel.target_part), rel.reltype)
            else:
                new_rId = new_slide_part.relate_to(copier.copy(rel.target_part), rel.reltype)
            rId_map[rId] = new_rId

        # --- add the layout relationship separately (uses target layout, not source) ---
        source_layout_rId = None
        for rId, rel in source_slide_part.rels.items():
            if rel.reltype == RT.SLIDE_LAYOUT:
                source_layout_rId = rId
                break
        if source_layout_rId is not None:
            new_layout_rId = new_slide_part.relate_to(target_layout_part, RT.SLIDE_LAYOUT)
            rId_map[source_layout_rId] = new_layout_rId

        # --- update rId references in the cloned XML ---
        _remap_rIds(new_element, rId_map)
        # --- comments aren't carried over, so drop the extension that points at them ---
        new_element.remove_comment_rel()

        rId = self.relate_to(new_slide_part, RT.SLIDE)
        return rId, new_slide_part.slide

    def _find_or_import_layout(self, source_layout_part: SlideLayoutPart) -> SlideLayoutPart:
        """Return a layout part in this presentation matching the source, importing if needed."""
        source_name = source_layout_part.slide_layout.name

        # --- search all masters in this presentation for a matching layout name ---
        for sldMasterId in self._element.get_or_add_sldMasterIdLst():
            master_part = self.related_part(sldMasterId.rId)
            for layout in master_part.slide_master.slide_layouts:
                if layout.name == source_name:
                    return layout.part

        # --- no match found — import the layout and its master ---
        return self._import_layout_and_master(source_layout_part)

    def _import_layout_and_master(self, source_layout_part: SlideLayoutPart) -> SlideLayoutPart:
        """Import a slide layout and its master into this presentation."""
        from pptx.opc.constants import CONTENT_TYPE as CT

        source_master_part = source_layout_part.part_related_by(RT.SLIDE_MASTER)

        # --- import the master ---
        master_partname = self.package.next_partname("/ppt/slideMasters/slideMaster%d.xml")
        new_master_element = deepcopy(source_master_part._element)
        new_master_part = SlideMasterPart(
            master_partname, CT.PML_SLIDE_MASTER, self.package, new_master_element
        )

        # --- import master's theme if present ---
        try:
            source_theme_part = source_master_part.part_related_by(RT.THEME)
            target_theme_part = self._import_part(source_theme_part)
            new_master_part.relate_to(target_theme_part, RT.THEME)
        except KeyError:
            pass

        # --- import master's non-layout relationships (images, etc.) ---
        _skip_master_reltypes = {RT.SLIDE_LAYOUT, RT.THEME}
        master_rId_map: dict[str, str] = {}
        for rId, rel in source_master_part.rels.items():
            if rel.reltype in _skip_master_reltypes:
                continue
            if rel.is_external:
                new_rId = new_master_part.relate_to(
                    rel.target_ref, rel.reltype, is_external=True
                )
            else:
                target_part = self._import_part(rel.target_part)
                new_rId = new_master_part.relate_to(target_part, rel.reltype)
            master_rId_map[rId] = new_rId
        _remap_rIds(new_master_element, master_rId_map)

        # --- register the master in presentation.xml ---
        master_rId = self.relate_to(new_master_part, RT.SLIDE_MASTER)
        sldMasterIdLst = self._element.get_or_add_sldMasterIdLst()
        sldMasterIdLst.add_sldMasterId(master_rId)

        # --- import the layout ---
        layout_partname = self.package.next_partname("/ppt/slideLayouts/slideLayout%d.xml")
        new_layout_element = deepcopy(source_layout_part._element)
        new_layout_part = SlideLayoutPart(
            layout_partname, CT.PML_SLIDE_LAYOUT, self.package, new_layout_element
        )
        new_layout_part.relate_to(new_master_part, RT.SLIDE_MASTER)

        # --- import layout's non-master relationships ---
        layout_rId_map: dict[str, str] = {}
        for rId, rel in source_layout_part.rels.items():
            if rel.reltype == RT.SLIDE_MASTER:
                continue
            if rel.is_external:
                new_rId = new_layout_part.relate_to(
                    rel.target_ref, rel.reltype, is_external=True
                )
            else:
                target_part = self._import_part(rel.target_part)
                new_rId = new_layout_part.relate_to(target_part, rel.reltype)
            layout_rId_map[rId] = new_rId
        _remap_rIds(new_layout_element, layout_rId_map)

        # --- register layout under the new master ---
        layout_rId = new_master_part.relate_to(new_layout_part, RT.SLIDE_LAYOUT)
        new_master_part._element.get_or_add_sldLayoutIdLst().add_sldLayoutId(layout_rId)

        return new_layout_part

    def _import_part(self, source_part: Part) -> Part:
        """Import a part from another package into this one, reusing if blob matches."""
        # --- for image/media parts, check if an identical part already exists ---
        source_blob = source_part.blob
        for existing_part in self.package.iter_parts():
            if (
                existing_part.content_type == source_part.content_type
                and existing_part.blob == source_blob
            ):
                return existing_part

        # --- create a new part with appropriate partname ---
        partname_tmpl = self._partname_template(source_part.partname)
        if partname_tmpl:
            new_partname = self.package.next_partname(partname_tmpl)
        else:
            new_partname = source_part.partname

        # --- for XmlPart subtypes, deep-copy the element ---
        if hasattr(source_part, "_element"):
            new_element = deepcopy(source_part._element)
            new_part = type(source_part)(
                new_partname, source_part.content_type, self.package, new_element
            )
        else:
            new_part = type(source_part)(
                new_partname, source_part.content_type, self.package, source_blob
            )
        return new_part

    @staticmethod
    def _partname_template(partname: PackURI) -> str | None:
        """Return a printf-style template for generating new partnames like the given one."""
        import re

        name = str(partname)
        match = re.match(r"^(.*?)(\d+)(\.\w+)$", name)
        if match:
            prefix, _, ext = match.groups()
            return f"{prefix}%d{ext}"
        return None

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


# --- relationship types that link a part into the slide/master hierarchy rather than to its own
# --- content; the part copier never follows these ---
_STRUCTURAL_RELTYPES = frozenset(
    (
        RT.SLIDE,
        RT.SLIDE_LAYOUT,
        RT.SLIDE_MASTER,
        RT.NOTES_SLIDE,
        RT.NOTES_MASTER,
        RT.HANDOUT_MASTER,
        RT.THEME,
        RT.COMMENTS,
        RT.MODERN_COMMENTS,
        RT.TAGS,
    )
)


class _PartCopier:
    """Copies parts into `package` together with the parts they relate to.

    A chart part, for example, is copied along with its embedded workbook, and the `r:id`
    references in the copied XML are remapped to the new relationships. With `dedup`, a part
    whose content type and blob match a part already in `package` is reused rather than
    copied (the behavior wanted when importing from another package, not when duplicating).

    Partnames assigned by one copier are tracked because `Package.next_partname()` only sees
    parts already reachable from the package, and new parts aren't reachable until the caller
    relates them.
    """

    def __init__(self, package, dedup: bool):
        self._package = package
        self._dedup = dedup
        self._copies: dict[Part, Part] = {}
        self._taken: set[str] | None = None

    def copy(self, source_part: Part) -> Part:
        """Return the copy of `source_part` in the target package, creating it if needed."""
        if source_part in self._copies:
            return self._copies[source_part]

        if self._dedup:
            existing = self._find_identical(source_part)
            if existing is not None:
                self._copies[source_part] = existing
                return existing

        new_part = self._new_part(source_part)
        self._copies[source_part] = new_part

        rId_map: dict[str, str] = {}
        for rId, rel in source_part.rels.items():
            if rel.is_external:
                rId_map[rId] = new_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            elif rel.reltype not in _STRUCTURAL_RELTYPES:
                rId_map[rId] = new_part.relate_to(self.copy(rel.target_part), rel.reltype)

        if isinstance(new_part, XmlPart):
            _remap_rIds(new_part._element, rId_map)
        return new_part

    def _find_identical(self, source_part: Part) -> Part | None:
        source_blob = source_part.blob
        for existing_part in self._package.iter_parts():
            if (
                existing_part.content_type == source_part.content_type
                and existing_part.blob == source_blob
            ):
                return existing_part
        return None

    def _new_part(self, source_part: Part) -> Part:
        partname = self._next_partname(source_part.partname)
        if isinstance(source_part, XmlPart):
            content = deepcopy(source_part._element)
        else:
            content = source_part.blob
        return type(source_part)(partname, source_part.content_type, self._package, content)

    def _next_partname(self, source_partname: PackURI) -> PackURI:
        if self._taken is None:
            self._taken = {str(p.partname) for p in self._package.iter_parts()}
        tmpl = PresentationPart._partname_template(source_partname)
        if tmpl is None:
            partname = str(source_partname)
        else:
            n = 1
            while tmpl % n in self._taken:
                n += 1
            partname = tmpl % n
        self._taken.add(partname)
        return PackURI(partname)


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
