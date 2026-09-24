"""Custom element classes for slide comment elements.

Two generations of comments exist. Legacy comments (`p:cmLst`, `p:cmAuthorLst`) are defined by
ECMA-376. Modern threaded comments (`p188:cmLst`, `p188:authorLst`) are defined by [MS-PPTX]
section 2.16.3 and are what PowerPoint 365 writes; they carry replies and a resolved state.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.oxml.simpletypes import ST_CommentStatus, XsdString, XsdUnsignedInt
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from pptx.oxml.shapes.shared import CT_Point2D
    from pptx.oxml.text import CT_TextBody


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


# -- modern (threaded) comments, [MS-PPTX] 2.16.3 ---------------------------------------------


class CT_ModernAuthor(BaseOxmlElement):
    """`p188:author` element, one author of modern comments, identified by a GUID."""

    id: str = RequiredAttribute("id", XsdString)  # pyright: ignore[reportAssignmentType]
    name: str = RequiredAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]
    initials: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "initials", XsdString
    )
    userId: str = RequiredAttribute("userId", XsdString)  # pyright: ignore[reportAssignmentType]
    providerId: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "providerId", XsdString
    )


class CT_ModernAuthorList(BaseOxmlElement):
    """`p188:authorLst` element, root of the authors part."""

    author_lst: list[CT_ModernAuthor]

    author = ZeroOrMore("p188:author")

    def get_author(self, author_id: str) -> CT_ModernAuthor | None:
        """Return the `p188:author` child whose `id` is `author_id`, or None."""
        for author in self.author_lst:
            if author.id == author_id:
                return author
        return None


class CT_SlideMoniker(BaseOxmlElement):
    """`pc:sldMk` element, identifies the slide a modern comment is anchored to."""

    cId: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "cId", XsdUnsignedInt
    )
    sldId: int = RequiredAttribute("sldId", XsdUnsignedInt)  # pyright: ignore[reportAssignmentType]


class CT_SlideMonikerList(BaseOxmlElement):
    """`pc:sldMkLst` element, the slide anchor of a modern comment."""

    sldMk: CT_SlideMoniker | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "pc:sldMk"
    )


class _BaseModernComment(BaseOxmlElement):
    """Attributes and text shared by `p188:cm` and `p188:reply` (`AG_CommentProperties`)."""

    txBody: CT_TextBody | None

    id: str = RequiredAttribute("id", XsdString)  # pyright: ignore[reportAssignmentType]
    authorId: str = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "authorId", XsdString
    )
    status: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "status", ST_CommentStatus, default=ST_CommentStatus.ACTIVE
    )
    created: str = RequiredAttribute("created", XsdString)  # pyright: ignore[reportAssignmentType]

    @property
    def text(self) -> str:
        """Text of the `p188:txBody` child, one line per paragraph; "" when there is none."""
        txBody = self.txBody
        if txBody is None:
            return ""
        return "\n".join(p.text for p in txBody.p_lst)


class CT_ModernCommentReply(_BaseModernComment):
    """`p188:reply` element, one reply in a comment thread."""

    txBody: CT_TextBody | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p188:txBody", successors=("p188:extLst",)
    )


class CT_ModernCommentReplyList(BaseOxmlElement):
    """`p188:replyLst` element, the replies of a comment thread in the order they were made."""

    reply_lst: list[CT_ModernCommentReply]

    reply = ZeroOrMore("p188:reply")


class CT_ModernComment(_BaseModernComment):
    """`p188:cm` element, the first comment of a thread; its replies are in `p188:replyLst`."""

    _tag_seq = (
        "pc:sldMkLst",
        "ac:deMkLst",
        "ac:txMkLst",
        "p188:unknownAnchor",
        "p188:pos",
        "p188:replyLst",
        "p188:txBody",
        "p188:extLst",
    )
    sldMkLst: CT_SlideMonikerList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "pc:sldMkLst", successors=_tag_seq[1:]
    )
    pos: CT_Point2D | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p188:pos", successors=_tag_seq[5:]
    )
    replyLst: CT_ModernCommentReplyList | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p188:replyLst", successors=_tag_seq[6:]
    )
    txBody: CT_TextBody | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p188:txBody", successors=_tag_seq[7:]
    )
    del _tag_seq

    title: str = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "title", XsdString, default=""
    )

    @property
    def anchor_tag(self) -> str | None:
        """Namespace-prefixed tag of the anchor child, e.g. "pc:sldMkLst", or None if absent."""
        from pptx.oxml.ns import NamespacePrefixedTag

        for child in self.iterchildren():
            if not isinstance(child.tag, str):
                continue
            tag = NamespacePrefixedTag.from_clark_name(child.tag)
            if tag in ("pc:sldMkLst", "ac:deMkLst", "ac:txMkLst", "p188:unknownAnchor"):
                return tag
        return None

    @property
    def reply_lst(self) -> list[CT_ModernCommentReply]:
        """The `p188:reply` grandchildren, in document order."""
        replyLst = self.replyLst
        return [] if replyLst is None else replyLst.reply_lst


class CT_ModernCommentList(BaseOxmlElement):
    """`p188:cmLst` element, root of a modern comments part holding one slide's threads."""

    cm_lst: list[CT_ModernComment]

    cm = ZeroOrMore("p188:cm")
