"""Initialization module for pypptx package."""

from __future__ import annotations

import sys
from typing import TYPE_CHECKING

import pptx.exc as exceptions
from pptx.api import Presentation
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import Part as _BlobPart
from pptx.opc.package import PartFactory, XmlPart
from pptx.parts.chart import ChartPart
from pptx.parts.chartex import ChartExPart
from pptx.parts.comments import (
    AuthorsPart,
    CommentAuthorsPart,
    CommentsPart,
    ModernCommentsPart,
)
from pptx.parts.coreprops import CorePropertiesPart
from pptx.parts.custprops import CustomPropertiesPart
from pptx.parts.diagram import (
    DiagramColorsPart,
    DiagramDataPart,
    DiagramDrawingPart,
    DiagramLayoutPart,
    DiagramStylePart,
)
from pptx.parts.image import ImagePart
from pptx.parts.media import MediaPart
from pptx.parts.presentation import PresentationPart
from pptx.parts.presprops import PresPropsPart
from pptx.parts.slide import (
    NotesMasterPart,
    NotesSlidePart,
    SlideLayoutPart,
    SlideMasterPart,
    SlidePart,
)
from pptx.parts.tablestyles import TableStylesPart
from pptx.parts.tags import TagsPart

if TYPE_CHECKING:
    from pptx.opc.package import Part

__version__ = "0.1.0"

sys.modules["pptx.exceptions"] = exceptions
del sys

__all__ = ["Presentation"]

content_type_to_part_class_map: dict[str, type[Part]] = {
    CT.PML_PRESENTATION_MAIN: PresentationPart,
    CT.PML_PRES_MACRO_MAIN: PresentationPart,
    CT.PML_TEMPLATE_MAIN: PresentationPart,
    CT.PML_SLIDESHOW_MAIN: PresentationPart,
    CT.OPC_CORE_PROPERTIES: CorePropertiesPart,
    CT.PML_NOTES_MASTER: NotesMasterPart,
    CT.PML_NOTES_SLIDE: NotesSlidePart,
    CT.PML_SLIDE: SlidePart,
    CT.PML_SLIDE_LAYOUT: SlideLayoutPart,
    CT.PML_SLIDE_MASTER: SlideMasterPart,
    CT.DML_CHART: ChartPart,
    CT.OFC_CHART_EX: ChartExPart,
    CT.DML_DIAGRAM_COLORS: DiagramColorsPart,
    CT.DML_DIAGRAM_DATA: DiagramDataPart,
    CT.DML_DIAGRAM_DRAWING: DiagramDrawingPart,
    CT.DML_DIAGRAM_LAYOUT: DiagramLayoutPart,
    CT.DML_DIAGRAM_STYLE: DiagramStylePart,
    CT.BMP: ImagePart,
    CT.GIF: ImagePart,
    CT.JPEG: ImagePart,
    CT.MS_PHOTO: ImagePart,
    CT.PNG: ImagePart,
    CT.TIFF: ImagePart,
    CT.X_EMF: ImagePart,
    CT.X_WMF: ImagePart,
    CT.ASF: MediaPart,
    CT.AVI: MediaPart,
    CT.M4A: MediaPart,
    CT.MOV: MediaPart,
    CT.MP3: MediaPart,
    CT.MP4: MediaPart,
    CT.MPG: MediaPart,
    CT.MS_VIDEO: MediaPart,
    CT.SWF: MediaPart,
    CT.VIDEO: MediaPart,
    CT.WAV: MediaPart,
    CT.WMV: MediaPart,
    CT.X_M4A: MediaPart,
    CT.X_MS_VIDEO: MediaPart,
    CT.X_WAV: MediaPart,
    CT.OFC_CUSTOM_PROPERTIES: CustomPropertiesPart,
    CT.PML_COMMENTS: CommentsPart,
    CT.PML_COMMENT_AUTHORS: CommentAuthorsPart,
    CT.PML_AUTHORS: AuthorsPart,
    CT.PML_MODERN_COMMENTS: ModernCommentsPart,
    CT.PML_TAGS: TagsPart,
    CT.PML_PRES_PROPS: PresPropsPart,
    CT.PML_TABLE_STYLES: TableStylesPart,
    CT.OFC_THEME: XmlPart,
    CT.OFC_THEME_OVERRIDE: XmlPart,
    # -- embedded font data is opaque; held and written back as a blob --
    CT.X_FONTDATA: _BlobPart,
    CT.X_FONT_TTF: _BlobPart,
    # -- accommodate "image/jpg" as an alias for "image/jpeg" --
    "image/jpg": ImagePart,
}

PartFactory.part_type_for.update(content_type_to_part_class_map)

del (
    AuthorsPart,
    ChartPart,
    ChartExPart,
    CommentAuthorsPart,
    CommentsPart,
    ModernCommentsPart,
    CorePropertiesPart,
    CustomPropertiesPart,
    DiagramColorsPart,
    DiagramDataPart,
    DiagramDrawingPart,
    DiagramLayoutPart,
    DiagramStylePart,
    ImagePart,
    MediaPart,
    SlidePart,
    SlideLayoutPart,
    SlideMasterPart,
    PresentationPart,
    PresPropsPart,
    TableStylesPart,
    TagsPart,
    XmlPart,
    _BlobPart,
    CT,
    PartFactory,
)
