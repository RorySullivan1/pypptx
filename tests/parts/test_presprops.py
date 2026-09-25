# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.parts.presprops` module."""

from __future__ import annotations

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.packuri import PackURI
from pptx.parts.presprops import PresPropsPart

from ..unitutil.mock import instance_mock


class DescribePresPropsPart:
    """Unit-test suite for `pptx.parts.presprops.PresPropsPart`."""

    def it_can_create_a_new_minimal_presProps_part(self, request):
        package_ = instance_mock(request, "pptx.package.Package")

        pres_props_part = PresPropsPart.new(package_)

        assert pres_props_part.partname == PackURI("/ppt/presProps.xml")
        assert pres_props_part.content_type == CT.PML_PRES_PROPS
        assert pres_props_part._element.showPr is None

    def it_reads_defaults_when_showPr_is_absent(self, request):
        package_ = instance_mock(request, "pptx.package.Package")
        pres_props_part = PresPropsPart.new(package_)

        assert pres_props_part._element.showPr is None

    def it_never_parses_when_only_the_blob_is_read_back(self, request):
        """Loading from a blob and reading `.blob` back must not parse the XML."""
        from pptx.oxml import parse_xml as real_parse_xml

        from ..unitutil.mock import function_mock

        parse_xml_ = function_mock(request, "pptx.opc.package.parse_xml")
        package_ = instance_mock(request, "pptx.package.Package")
        blob = real_parse_xml(
            '<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/'
            '2006/main"/>'
        ).xml.encode("utf-8")

        part = PresPropsPart.load(PackURI("/ppt/presProps.xml"), CT.PML_PRES_PROPS, package_, blob)

        assert part.blob == blob
        parse_xml_.assert_not_called()
