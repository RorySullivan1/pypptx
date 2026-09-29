"""Custom element classes for tag-related XML elements."""

from __future__ import annotations

from typing import Callable

from pptx.oxml.simpletypes import XsdString
from pptx.oxml.xmlchemy import BaseOxmlElement, RequiredAttribute, ZeroOrMore, ZeroOrOne


class CT_StringTag(BaseOxmlElement):
    """`p:tag` element, a single name-value string pair."""

    name: str = RequiredAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]
    val: str = RequiredAttribute("val", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_TagList(BaseOxmlElement):
    """`p:tagLst` element, container for string tags on a shape or slide."""

    _add_tag: Callable[..., CT_StringTag]
    tag_lst: list[CT_StringTag]

    tag = ZeroOrMore("p:tag")


class CT_TagsData(BaseOxmlElement):
    """`p:tags` element, the reference from a shape or slide to its tags part."""

    rId: str = RequiredAttribute("r:id", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_CustomerDataList(BaseOxmlElement):
    """`p:custDataLst` element, customer data of a shape (in `p:nvPr`) or slide (in `p:cSld`).

    Holds `p:custData` references to custom XML parts, then at most one `p:tags` reference to
    the tags part.
    """

    get_or_add_tags: Callable[[], CT_TagsData]
    _remove_tags: Callable[[], None]

    tags: CT_TagsData | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:tags", successors=()
    )


class CustDataLstOwnerMixin:
    """Tags access shared by the elements holding a `p:custDataLst`: `p:nvPr` and `p:cSld`."""

    custDataLst: CT_CustomerDataList | None
    get_or_add_custDataLst: Callable[[], CT_CustomerDataList]
    _remove_custDataLst: Callable[[], None]

    @property
    def tags_rId(self) -> str | None:
        """`r:id` of the `p:custDataLst/p:tags` reference to the tags part, None when absent."""
        custDataLst = self.custDataLst
        if custDataLst is None or custDataLst.tags is None:
            return None
        return custDataLst.tags.rId

    def set_tags_rId(self, rId: str) -> None:
        """Point the `p:custDataLst/p:tags` reference at relationship `rId`, adding it if needed."""
        self.get_or_add_custDataLst().get_or_add_tags().rId = rId

    def remove_tags_rId(self) -> None:
        """Remove the `p:tags` reference, and the `p:custDataLst` too when that leaves it empty."""
        custDataLst = self.custDataLst
        if custDataLst is None:
            return
        custDataLst._remove_tags()  # pyright: ignore[reportPrivateUsage]
        if len(custDataLst) == 0:
            self._remove_custDataLst()
