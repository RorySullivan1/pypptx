"""Custom element classes for slide comment elements."""

from __future__ import annotations

from pptx.oxml.simpletypes import XsdString, XsdUnsignedInt
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


class CT_CommentAuthor(BaseOxmlElement):
    """`p:cmAuthor` element, identifying a comment author."""

    id: int = RequiredAttribute("id", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]
    name: str = RequiredAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]
    initials: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "initials", XsdString
    )
    lastIdx: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "lastIdx", XsdUnsignedInt
    )
    clrIdx: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "clrIdx", XsdUnsignedInt
    )


class CT_CommentAuthorList(BaseOxmlElement):
    """`p:cmAuthorLst` element, container for comment authors."""

    cmAuthor_lst: list[CT_CommentAuthor]

    cmAuthor = ZeroOrMore("p:cmAuthor")


class CT_Comment(BaseOxmlElement):
    """`p:cm` element, a single slide comment."""

    _tag_seq = ("p:pos", "p:text", "p:extLst")
    pos = ZeroOrOne("p:pos", successors=_tag_seq[1:])
    text = ZeroOrOne("p:text", successors=_tag_seq[2:])
    del _tag_seq

    authorId: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "authorId", XsdUnsignedInt
    )
    dt: str | None = OptionalAttribute("dt", XsdString)  # pyright: ignore[reportAssignmentType]
    idx: int = RequiredAttribute("idx", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]


class CT_CommentList(BaseOxmlElement):
    """`p:cmLst` element, container for slide comments."""

    cm_lst: list[CT_Comment]

    cm = ZeroOrMore("p:cm")
