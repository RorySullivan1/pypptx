"""Custom element classes for slide comment elements.

Two generations of comments exist. Legacy comments (`p:cmLst`, `p:cmAuthorLst`) are defined by
ECMA-376. Modern threaded comments (`p188:cmLst`, `p188:authorLst`) are defined by [MS-PPTX]
section 2.16.3 and are what PowerPoint 365 writes; they carry replies and a resolved state.
"""

from __future__ import annotations

import datetime as dt
import re
import uuid
from typing import TYPE_CHECKING, Callable

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
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

    _add_author: Callable[..., CT_ModernAuthor]
    author_lst: list[CT_ModernAuthor]

    author = ZeroOrMore("p188:author")

    @classmethod
    def new(cls) -> CT_ModernAuthorList:
        """Return a new, empty `p188:authorLst` element."""
        return parse_xml("<p188:authorLst %s/>" % nsdecls("p188", "a", "r"))

    def add_author(self, name: str, initials: str | None) -> CT_ModernAuthor:
        """Append and return a new `p188:author` with a fresh GUID.

        `userId` is `name` and `providerId` is "None", which is how PowerPoint records an author
        who isn't signed in to an identity provider.
        """
        return self._add_author(
            id=new_guid(), name=name, initials=initials, userId=name, providerId="None"
        )

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

    _add_txBody: Callable[[], CT_TextBody]
    _remove_txBody: Callable[[], None]
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
    def created_datetime(self) -> dt.datetime | None:
        """`created` as a |datetime|, or None when it isn't a valid `xsd:dateTime`.

        The result is timezone-aware when the value carries a `Z` or `+hh:mm` suffix and naive
        (local time, which is what PowerPoint writes) when it doesn't.
        """
        return _parse_xsd_datetime(self.created)

    @property
    def text(self) -> str:
        """Text of the `p188:txBody` child, one line per paragraph; "" when there is none."""
        txBody = self.txBody
        if txBody is None:
            return ""
        return "\n".join(p.text for p in txBody.p_lst)

    @text.setter
    def text(self, value: str) -> None:
        """Replace the text with `value`, one `a:p` per line of it."""
        from pptx.oxml.xmlchemy import OxmlElement

        self._remove_txBody()
        txBody = self._add_txBody()
        txBody.append(OxmlElement("a:bodyPr"))
        txBody.append(OxmlElement("a:lstStyle"))
        for line in value.split("\n"):
            p = txBody.add_p()
            if line:
                p.add_r(line)


class CT_ModernCommentReply(_BaseModernComment):
    """`p188:reply` element, one reply in a comment thread."""

    txBody: CT_TextBody | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p188:txBody", successors=("p188:extLst",)
    )


class CT_ModernCommentReplyList(BaseOxmlElement):
    """`p188:replyLst` element, the replies of a comment thread in the order they were made."""

    _add_reply: Callable[..., CT_ModernCommentReply]
    reply_lst: list[CT_ModernCommentReply]

    reply = ZeroOrMore("p188:reply")


class CT_ModernComment(_BaseModernComment):
    """`p188:cm` element, the first comment of a thread; its replies are in `p188:replyLst`."""

    get_or_add_replyLst: Callable[[], CT_ModernCommentReplyList]

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

    @classmethod
    def new(cls, author_id: str, created: str, sld_id: int, creation_id: int) -> CT_ModernComment:
        """Return a new `p188:cm` with a fresh GUID, anchored to the slide these ids identify."""
        return parse_xml(
            '<p188:cm %s id="%s" authorId="%s" created="%s">'
            '<pc:sldMkLst><pc:docMk/><pc:sldMk cId="%d" sldId="%d"/></pc:sldMkLst>'
            "</p188:cm>"
            % (nsdecls("p188", "pc", "a"), new_guid(), author_id, created, creation_id, sld_id)
        )

    def add_reply(self, author_id: str, created: str, text: str) -> CT_ModernCommentReply:
        """Append and return a new `p188:reply` with a fresh GUID, after any existing replies."""
        reply = self.get_or_add_replyLst()._add_reply(
            id=new_guid(), authorId=author_id, created=created
        )
        reply.text = text
        return reply

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

    _insert_cm: Callable[[CT_ModernComment], CT_ModernComment]
    cm_lst: list[CT_ModernComment]

    cm = ZeroOrMore("p188:cm")

    @classmethod
    def new(cls) -> CT_ModernCommentList:
        """Return a new, empty `p188:cmLst` element."""
        return parse_xml("<p188:cmLst %s/>" % nsdecls("p188", "a", "r"))

    def add_cm(
        self, author_id: str, created: str, sld_id: int, creation_id: int, text: str
    ) -> CT_ModernComment:
        """Append and return a new slide-anchored thread."""
        cm = self._insert_cm(CT_ModernComment.new(author_id, created, sld_id, creation_id))
        cm.text = text
        return cm


def new_guid() -> str:
    """Return a new random GUID in the braced upper-case form PowerPoint writes."""
    return "{%s}" % str(uuid.uuid4()).upper()


_XSD_DATETIME_RE = re.compile(
    r"^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})(?:\.(\d+))?(Z|[+-]\d{2}:\d{2})?$"
)


def _parse_xsd_datetime(value: str) -> dt.datetime | None:
    """Return the |datetime| an `xsd:dateTime` string names, or None when it isn't one.

    Fractional seconds of any length are accepted and truncated to microseconds.
    """
    match = _XSD_DATETIME_RE.match(value.strip())
    if match is None:
        return None
    seconds_part, fraction, zone = match.groups()
    try:
        result = dt.datetime.strptime(seconds_part, "%Y-%m-%dT%H:%M:%S")
    except ValueError:
        return None
    if fraction:
        result = result.replace(microsecond=int(fraction[:6].ljust(6, "0")))
    if zone == "Z":
        result = result.replace(tzinfo=dt.timezone.utc)
    elif zone:
        sign = -1 if zone[0] == "-" else 1
        offset = dt.timedelta(hours=int(zone[1:3]), minutes=int(zone[4:6]))
        result = result.replace(tzinfo=dt.timezone(sign * offset))
    return result
