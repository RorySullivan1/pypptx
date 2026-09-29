"""Unit-test suite for `pptx.oxml.tags` module."""

from __future__ import annotations

import re

import pytest

from pptx.oxml.tags import CT_CustomerDataList, CT_TagsData

from ..unitutil.cxml import element, xml


class DescribeCT_CustomerDataList:
    """Unit-test suite for `pptx.oxml.tags.CT_CustomerDataList` objects."""

    def it_provides_access_to_its_tags_reference(self):
        custDataLst = element("p:custDataLst/(p:custData{r:id=rId1},p:tags{r:id=rId2})")

        assert isinstance(custDataLst, CT_CustomerDataList)
        assert isinstance(custDataLst.tags, CT_TagsData)
        assert custDataLst.tags.rId == "rId2"

    def it_adds_the_tags_reference_after_the_custom_data(self):
        custDataLst = element("p:custDataLst/p:custData{r:id=rId1}")

        custDataLst.get_or_add_tags().rId = "rId2"

        assert custDataLst.xml == xml("p:custDataLst/(p:custData{r:id=rId1},p:tags{r:id=rId2})")


class DescribeCustDataLstOwnerMixin:
    """Unit-test suite for the tags access shared by `p:nvPr` and `p:cSld`."""

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("p:nvPr/p:custDataLst/p:tags{r:id=rId4}", "rId4"),
            ("p:nvPr/p:custDataLst/p:custData{r:id=rId1}", None),
            ("p:nvPr", None),
            ("p:cSld/(p:spTree,p:custDataLst/p:tags{r:id=rId7})", "rId7"),
            ("p:cSld/p:spTree", None),
        ],
    )
    def it_knows_its_tags_rId(self, cxml: str, expected_value: str | None):
        assert element(cxml).tags_rId == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_children"),
        [
            ("p:nvPr/(p:ph,p:extLst)", ["ph", "custDataLst", "extLst"]),
            ("p:nvPr/p:custDataLst/p:tags{r:id=rId1}", ["custDataLst"]),
            ("p:cSld/(p:spTree,p:extLst)", ["spTree", "custDataLst", "extLst"]),
        ],
    )
    def it_can_set_its_tags_rId(self, cxml: str, expected_children: list[str]):
        owner = element(cxml)

        owner.set_tags_rId("rId9")

        assert [child.tag.split("}")[1] for child in owner] == expected_children
        assert owner.tags_rId == "rId9"
        assert len(owner.xpath("./p:custDataLst/p:tags")) == 1

    @pytest.mark.parametrize(
        ("cxml", "expected_cxml"),
        [
            ("p:nvPr/(p:ph,p:custDataLst/p:tags{r:id=rId1})", "p:nvPr/p:ph"),
            (
                "p:nvPr/p:custDataLst/(p:custData{r:id=rId1},p:tags{r:id=rId2})",
                "p:nvPr/p:custDataLst/p:custData{r:id=rId1}",
            ),
            ("p:cSld/(p:spTree,p:custDataLst/p:tags{r:id=rId1})", "p:cSld/p:spTree"),
            ("p:cSld/p:spTree", "p:cSld/p:spTree"),
        ],
    )
    def it_can_remove_its_tags_reference(self, cxml: str, expected_cxml: str):
        owner = element(cxml)

        owner.remove_tags_rId()

        assert _strip_ns_decls(owner.xml) == _strip_ns_decls(xml(expected_cxml))


def _strip_ns_decls(xml_text: str) -> str:
    """`xml_text` without its namespace declarations, which cxml adds only where used."""
    return re.sub(r' xmlns:\w+="[^"]*"', "", xml_text)
