# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.parts.tablestyles` module."""

from __future__ import annotations

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.packuri import PackURI
from pptx.parts.tablestyles import TableStylesPart

from ..unitutil.mock import function_mock, instance_mock


class DescribeTableStylesPart:
    """Unit-test suite for `pptx.parts.tablestyles.TableStylesPart`."""

    def it_can_create_a_new_empty_tableStyles_part(self, request):
        package_ = instance_mock(request, "pptx.package.Package")

        table_styles_part = TableStylesPart.new(package_)

        assert table_styles_part.partname == PackURI("/ppt/tableStyles.xml")
        assert table_styles_part.content_type == CT.PML_TABLE_STYLES
        assert table_styles_part._element.tblStyle_lst == []
        assert table_styles_part._element.def_ is None

    def it_never_parses_when_only_the_blob_is_read_back(self, request):
        from pptx.oxml import parse_xml as real_parse_xml

        parse_xml_ = function_mock(request, "pptx.opc.package.parse_xml")
        package_ = instance_mock(request, "pptx.package.Package")
        blob = real_parse_xml(
            '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>'
        ).xml.encode("utf-8")

        part = TableStylesPart.load(
            PackURI("/ppt/tableStyles.xml"), CT.PML_TABLE_STYLES, package_, blob
        )

        assert part.blob == blob
        parse_xml_.assert_not_called()
