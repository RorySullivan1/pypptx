# pyright: reportPrivateUsage=false

"""Unit-test suite for slide comments API."""

from __future__ import annotations

import pytest

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.parts.comments import CommentAuthorsPart, CommentsPart
from pptx.slide import Comment, SlideComments


class DescribeCommentAuthorsPart:
    """Unit-test suite for `pptx.parts.comments.CommentAuthorsPart`."""

    def _part(self):
        xml = '<p:cmAuthorLst %s/>' % nsdecls("p")
        element = parse_xml(xml)
        return CommentAuthorsPart("/ppt/commentAuthors.xml", "application/test", None, element)

    def it_starts_empty(self):
        part = self._part()
        assert len(part) == 0
        assert list(part) == []

    def it_can_add_an_author(self):
        part = self._part()
        author = part.get_or_add_author("Alice", "A")
        assert author.name == "Alice"
        assert author.initials == "A"
        assert author.id == 0
        assert len(part) == 1

    def it_reuses_existing_authors(self):
        part = self._part()
        a1 = part.get_or_add_author("Alice", "A")
        a2 = part.get_or_add_author("Alice", "A")
        assert a1 is a2
        assert len(part) == 1

    def it_assigns_unique_ids(self):
        part = self._part()
        a1 = part.get_or_add_author("Alice", "A")
        a2 = part.get_or_add_author("Bob", "B")
        assert a1.id == 0
        assert a2.id == 1

    def it_can_find_an_author_by_id(self):
        part = self._part()
        part.get_or_add_author("Alice", "A")
        author = part.get_author(0)
        assert author is not None
        assert author.name == "Alice"
        assert part.get_author(99) is None


class DescribeCommentsPart:
    """Unit-test suite for `pptx.parts.comments.CommentsPart`."""

    def _part(self, n_comments=0):
        parts = ['<p:cmLst %s>' % nsdecls("p")]
        for i in range(n_comments):
            parts.append(
                f'<p:cm authorId="0" idx="{i}">'
                f'<p:pos x="100" y="200"/>'
                f'<p:text>Comment {i}</p:text>'
                f'</p:cm>'
            )
        parts.append('</p:cmLst>')
        element = parse_xml("".join(parts))
        return CommentsPart("/ppt/comments/comment1.xml", "application/test", None, element)

    def it_starts_empty(self):
        part = self._part()
        assert len(part) == 0

    def it_counts_comments(self):
        part = self._part(n_comments=3)
        assert len(part) == 3

    def it_iterates_comments(self):
        part = self._part(n_comments=2)
        comments = list(part)
        assert len(comments) == 2


class DescribeComment:
    """Unit-test suite for `pptx.slide.Comment`."""

    def _make_comment(self, author_name="Alice", text="Hello", x=100, y=200):
        # Build authors part
        authors_xml = (
            '<p:cmAuthorLst %s>'
            '<p:cmAuthor id="0" name="%s" initials="A" lastIdx="0" clrIdx="0"/>'
            '</p:cmAuthorLst>'
        ) % (nsdecls("p"), author_name)
        authors_element = parse_xml(authors_xml)
        authors_part = CommentAuthorsPart(
            "/ppt/commentAuthors.xml", "application/test", None, authors_element
        )

        # Build comment element
        cm_xml = (
            '<p:cm %s authorId="0" idx="0">'
            '<p:pos x="%d" y="%d"/>'
            '<p:text>%s</p:text>'
            '</p:cm>'
        ) % (nsdecls("p"), x, y, text)
        cm = parse_xml(cm_xml)
        return Comment(cm, authors_part)

    def it_knows_its_author(self):
        comment = self._make_comment(author_name="Alice")
        assert comment.author == "Alice"

    def it_knows_its_text(self):
        comment = self._make_comment(text="Hello world")
        assert comment.text == "Hello world"

    def it_can_change_its_text(self):
        comment = self._make_comment(text="Old")
        comment.text = "New"
        assert comment.text == "New"

    def it_knows_its_position(self):
        comment = self._make_comment(x=100, y=200)
        assert comment.position == (100, 200)

    def it_can_be_deleted(self):
        # Build a comment list with one comment
        xml = (
            '<p:cmLst %s>'
            '<p:cm authorId="0" idx="0">'
            '<p:pos x="0" y="0"/><p:text>Test</p:text>'
            '</p:cm>'
            '</p:cmLst>'
        ) % nsdecls("p")
        cmLst = parse_xml(xml)
        cm = cmLst[0]
        authors_xml = '<p:cmAuthorLst %s/>' % nsdecls("p")
        authors_part = CommentAuthorsPart(
            "/ppt/commentAuthors.xml", "application/test", None, parse_xml(authors_xml)
        )
        comment = Comment(cm, authors_part)
        comment.delete()
        assert len(cmLst) == 0
