# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.parts.diagram` module."""

from __future__ import annotations

import pytest

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import PartFactory
from pptx.opc.packuri import PackURI
from pptx.oxml.diagram import CT_DataModel
from pptx.parts.diagram import (
    DiagramColorsPart,
    DiagramDataPart,
    DiagramDrawingPart,
    DiagramLayoutPart,
    DiagramStylePart,
)

from ..unitutil.mock import function_mock, instance_mock

DATA_BLOB = (
    b'<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
    b'<dgm:ptLst><dgm:pt modelId="0" type="doc"/></dgm:ptLst></dgm:dataModel>'
)


class DescribeDiagramParts:
    """Unit-test suite for the `pptx.parts.diagram` part classes."""

    @pytest.mark.parametrize(
        ("content_type", "expected_cls"),
        [
            (CT.DML_DIAGRAM_DATA, DiagramDataPart),
            (CT.DML_DIAGRAM_LAYOUT, DiagramLayoutPart),
            (CT.DML_DIAGRAM_STYLE, DiagramStylePart),
            (CT.DML_DIAGRAM_COLORS, DiagramColorsPart),
            (CT.DML_DIAGRAM_DRAWING, DiagramDrawingPart),
        ],
    )
    def it_is_registered_for_its_content_type(self, content_type: str, expected_cls: type):
        assert PartFactory.part_type_for[content_type] is expected_cls

    def it_never_parses_when_only_the_blob_is_read_back(self, request):
        parse_xml_ = function_mock(request, "pptx.opc.package.parse_xml")
        package_ = instance_mock(request, "pptx.package.Package")

        part = DiagramDataPart.load(
            PackURI("/ppt/diagrams/data1.xml"), CT.DML_DIAGRAM_DATA, package_, DATA_BLOB
        )

        assert part.blob == DATA_BLOB
        parse_xml_.assert_not_called()


class DescribeDiagramDataPart:
    """Unit-test suite for `pptx.parts.diagram.DiagramDataPart` objects."""

    def it_provides_access_to_its_data_model(self, request):
        package_ = instance_mock(request, "pptx.package.Package")
        part = DiagramDataPart.load(
            PackURI("/ppt/diagrams/data1.xml"), CT.DML_DIAGRAM_DATA, package_, DATA_BLOB
        )

        data_model = part.data_model

        assert isinstance(data_model, CT_DataModel)
        assert data_model.doc_pt is not None
