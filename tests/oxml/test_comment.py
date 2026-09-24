"""Unit-test suite for `pptx.oxml.comment` modern (threaded) comment element classes."""

from __future__ import annotations

import pytest

from pptx.oxml.comment import (
    CT_ModernAuthor,
    CT_ModernAuthorList,
    CT_ModernComment,
    CT_ModernCommentList,
    CT_ModernCommentReply,
)

from ..unitutil.cxml import element


class DescribeCT_ModernAuthorList:
    """Unit-test suite for `pptx.oxml.comment.CT_ModernAuthorList` and its authors."""

    def it_provides_access_to_its_authors(self):
        authorLst = element(
            "p188:authorLst/(p188:author{id=a1,name=Alice,initials=AA,userId=alice,providerId=None}"
            ",p188:author{id=b2,name=Bob,userId=bob,providerId=AD})"
        )

        assert isinstance(authorLst, CT_ModernAuthorList)
        alice, bob = authorLst.author_lst
        assert isinstance(alice, CT_ModernAuthor)
        assert (alice.id, alice.name, alice.initials, alice.userId, alice.providerId) == (
            "a1",
            "Alice",
            "AA",
            "alice",
            "None",
        )
        assert bob.initials is None

    @pytest.mark.parametrize(("author_id", "expected_name"), [("b2", "Bob"), ("zz", None)])
    def it_can_find_an_author_by_id(self, author_id: str, expected_name: str | None):
        authorLst = element(
            "p188:authorLst/(p188:author{id=a1,name=Alice,userId=alice,providerId=None}"
            ",p188:author{id=b2,name=Bob,userId=bob,providerId=AD})"
        )

        author = authorLst.get_author(author_id)

        assert (author.name if author is not None else None) == expected_name


class DescribeCT_ModernComment:
    """Unit-test suite for `pptx.oxml.comment.CT_ModernComment` and `CT_ModernCommentReply`."""

    def it_knows_its_identity_attributes(self):
        cm = element("p188:cm{id=c1,authorId=a1,created=2024-05-01T09:30:00.123,title=Todo}")

        assert isinstance(cm, CT_ModernComment)
        assert (cm.id, cm.authorId, cm.created, cm.title) == (
            "c1",
            "a1",
            "2024-05-01T09:30:00.123",
            "Todo",
        )

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("p188:cm{id=c1,authorId=a1,created=x}", "active"),
            ("p188:cm{id=c1,authorId=a1,created=x,status=resolved}", "resolved"),
            ("p188:cm{id=c1,authorId=a1,created=x,status=closed}", "closed"),
        ],
    )
    def it_knows_its_status(self, cxml: str, expected_value: str):
        assert element(cxml).status == expected_value

    @pytest.mark.parametrize(
        ("cxml", "expected_value"),
        [
            ("p188:cm{id=c1,authorId=a1,created=x}", ""),
            ('p188:cm{id=c1,authorId=a1,created=x}/p188:txBody/(a:bodyPr,a:p/a:r/a:t"hi")', "hi"),
            (
                "p188:cm{id=c1,authorId=a1,created=x}/p188:txBody/"
                '(a:bodyPr,a:p/a:r/a:t"one",a:p/(a:r/a:t"tw",a:r/a:t"o"))',
                "one\ntwo",
            ),
        ],
    )
    def it_knows_its_text(self, cxml: str, expected_value: str):
        assert element(cxml).text == expected_value

    def it_provides_access_to_its_replies_in_order(self):
        cm = element(
            "p188:cm{id=c1,authorId=a1,created=x}/(pc:sldMkLst/pc:sldMk{sldId=256}"
            ',p188:replyLst/(p188:reply{id=r1,authorId=b2,created=y}/p188:txBody/(a:bodyPr,a:p/a:r/a:t"one")'
            ',p188:reply{id=r2,authorId=a1,created=z,status=resolved})'
            ',p188:txBody/(a:bodyPr,a:p))'
        )

        replies = cm.reply_lst

        assert [type(r) for r in replies] == [CT_ModernCommentReply, CT_ModernCommentReply]
        assert [(r.id, r.authorId, r.text, r.status) for r in replies] == [
            ("r1", "b2", "one", "active"),
            ("r2", "a1", "", "resolved"),
        ]

    def but_it_has_no_replies_when_it_has_no_replyLst(self):
        assert element("p188:cm{id=c1,authorId=a1,created=x}").reply_lst == []

    @pytest.mark.parametrize(
        ("anchor_cxml", "expected_value"),
        [
            ("/pc:sldMkLst/pc:sldMk{sldId=256}", "pc:sldMkLst"),
            ("/ac:deMkLst/pc:sldMk{sldId=256}", "ac:deMkLst"),
            ("/ac:txMkLst", "ac:txMkLst"),
            ("/p188:unknownAnchor", "p188:unknownAnchor"),
            ("", None),
        ],
    )
    def it_knows_which_kind_of_anchor_it_has(self, anchor_cxml: str, expected_value: str | None):
        cm = element("p188:cm{id=c1,authorId=a1,created=x}" + anchor_cxml)
        assert cm.anchor_tag == expected_value

    def it_provides_access_to_its_slide_moniker(self):
        cm = element("p188:cm{id=c1,authorId=a1,created=x}/pc:sldMkLst/pc:sldMk{cId=99,sldId=257}")

        sldMk = cm.sldMkLst.sldMk

        assert (sldMk.cId, sldMk.sldId) == (99, 257)


class DescribeCT_ModernCommentList:
    """Unit-test suite for `pptx.oxml.comment.CT_ModernCommentList`."""

    def it_provides_access_to_its_threads(self):
        cmLst = element(
            "p188:cmLst/(p188:cm{id=c1,authorId=a1,created=x},p188:cm{id=c2,authorId=a1,created=y})"
        )

        assert isinstance(cmLst, CT_ModernCommentList)
        assert [cm.id for cm in cmLst.cm_lst] == ["c1", "c2"]
