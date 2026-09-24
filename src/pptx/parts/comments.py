"""Comments parts for slide comments and comment authors, legacy and modern."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Iterator

from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.package import XmlPart
from pptx.opc.packuri import PackURI
from pptx.oxml.comment import (
    CT_Comment,
    CT_CommentAuthor,
    CT_CommentAuthorList,
    CT_CommentList,
    CT_ModernAuthor,
    CT_ModernAuthorList,
    CT_ModernComment,
    CT_ModernCommentList,
)

if TYPE_CHECKING:
    from pptx.package import Package


class CommentAuthorsPart(XmlPart):
    """Corresponds to ``/ppt/commentAuthors.xml`` part.

    Contains the list of comment authors for this presentation.
    """

    _element: CT_CommentAuthorList

    @classmethod
    def default(cls, package: Package) -> CommentAuthorsPart:
        """Return a new empty |CommentAuthorsPart|."""
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls

        xml = "<p:cmAuthorLst %s/>" % nsdecls("p")
        element = parse_xml(xml)
        return cls(
            PackURI("/ppt/commentAuthors.xml"),
            CT.PML_COMMENT_AUTHORS,
            package,
            element,
        )

    def get_author(self, author_id: int) -> CT_CommentAuthor | None:
        """Return the author element with the given ID, or None."""
        for author in self._element.cmAuthor_lst:
            if author.id == author_id:
                return author
        return None

    def get_or_add_author(self, name: str, initials: str) -> CT_CommentAuthor:
        """Return existing author matching name/initials, or create a new one."""
        for author in self._element.cmAuthor_lst:
            if author.name == name and author.initials == initials:
                return author
        return self._add_author(name, initials)

    def _add_author(self, name: str, initials: str) -> CT_CommentAuthor:
        """Create and return a new author element."""
        from lxml import etree

        from pptx.oxml.ns import qn

        used_ids = [a.id for a in self._element.cmAuthor_lst]
        next_id = max(used_ids, default=-1) + 1
        used_clr = [a.clrIdx for a in self._element.cmAuthor_lst]
        next_clr = max(used_clr, default=-1) + 1

        author = etree.SubElement(self._element, qn("p:cmAuthor"))
        author.set("id", str(next_id))
        author.set("name", name)
        author.set("initials", initials)
        author.set("lastIdx", "0")
        author.set("clrIdx", str(next_clr))
        return author

    @property
    def authors(self) -> list[CT_CommentAuthor]:
        """List of all comment author elements."""
        return self._element.cmAuthor_lst

    def __len__(self) -> int:
        return len(self._element.cmAuthor_lst)

    def __iter__(self) -> Iterator[CT_CommentAuthor]:
        return iter(self._element.cmAuthor_lst)


class CommentsPart(XmlPart):
    """Corresponds to ``/ppt/comments/commentN.xml`` part.

    Contains the comments for a single slide.
    """

    _element: CT_CommentList

    @classmethod
    def default(cls, package: Package, partname: str) -> CommentsPart:
        """Return a new empty |CommentsPart|."""
        from pptx.oxml import parse_xml
        from pptx.oxml.ns import nsdecls

        xml = "<p:cmLst %s/>" % nsdecls("p")
        element = parse_xml(xml)
        return cls(PackURI(partname), CT.PML_COMMENTS, package, element)

    @property
    def comments(self) -> list[CT_Comment]:
        """List of all comment elements."""
        return self._element.cm_lst

    def __len__(self) -> int:
        return len(self._element.cm_lst)

    def __iter__(self) -> Iterator[CT_Comment]:
        return iter(self._element.cm_lst)


class AuthorsPart(XmlPart):
    """Corresponds to the ``/ppt/authors.xml`` part, the authors of modern comments.

    The presentation part relates to it with the `RT.AUTHORS` relationship.
    """

    _element: CT_ModernAuthorList

    def get_author(self, author_id: str) -> CT_ModernAuthor | None:
        """Return the author element whose GUID is `author_id`, or None."""
        return self._element.get_author(author_id)

    @property
    def authors(self) -> list[CT_ModernAuthor]:
        """List of all author elements."""
        return self._element.author_lst

    def __len__(self) -> int:
        return len(self._element.author_lst)

    def __iter__(self) -> Iterator[CT_ModernAuthor]:
        return iter(self._element.author_lst)


class ModernCommentsPart(XmlPart):
    """Corresponds to a ``/ppt/comments/modernComment_*.xml`` part, one slide's comment threads.

    A slide relates to it with the `RT.MODERN_COMMENTS` relationship and also names that
    relationship in a `p188:commentRel` extension of its `p:sld/p:extLst`.
    """

    _element: CT_ModernCommentList

    @property
    def comments(self) -> list[CT_ModernComment]:
        """List of all thread-starting comment elements."""
        return self._element.cm_lst

    def __len__(self) -> int:
        return len(self._element.cm_lst)

    def __iter__(self) -> Iterator[CT_ModernComment]:
        return iter(self._element.cm_lst)
