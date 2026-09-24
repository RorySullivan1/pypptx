# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.fonts` module."""

from __future__ import annotations

from typing import cast

import pytest

from pptx.fonts import EmbeddedFont, EmbeddedFonts
from pptx.oxml.presentation import CT_Presentation
from pptx.parts.presentation import PresentationPart

from .unitutil.cxml import element
from .unitutil.mock import instance_mock, method_mock


class DescribeEmbeddedFonts:
    """Unit-test suite for `pptx.fonts.EmbeddedFonts` objects."""

    def it_is_empty_when_no_embeddedFontLst_is_present(self, part_):
        prs_elm = cast(CT_Presentation, element("p:presentation"))
        embedded_fonts = EmbeddedFonts(prs_elm, part_)

        assert len(embedded_fonts) == 0
        assert list(embedded_fonts) == []

    def it_knows_how_many_fonts_are_embedded(self, part_):
        embedded_fonts = EmbeddedFonts(self._two_font_prs_elm(), part_)
        assert len(embedded_fonts) == 2

    def it_supports_iteration(self, part_):
        embedded_fonts = EmbeddedFonts(self._two_font_prs_elm(), part_)
        typefaces = [font.typeface for font in embedded_fonts]
        assert typefaces == ["Calibri", "Arial"]

    def it_supports_indexed_access(self, part_):
        embedded_fonts = EmbeddedFonts(self._two_font_prs_elm(), part_)
        assert embedded_fonts[0].typeface == "Calibri"
        assert embedded_fonts[1].typeface == "Arial"

    def it_supports_slicing(self, part_):
        embedded_fonts = EmbeddedFonts(self._two_font_prs_elm(), part_)
        fonts = embedded_fonts[0:2]
        assert [f.typeface for f in fonts] == ["Calibri", "Arial"]

    def it_can_remove_a_font_via_the_collection(self, request, part_):
        prs_elm = self._two_font_prs_elm()
        embedded_fonts = EmbeddedFonts(prs_elm, part_)
        font = embedded_fonts[0]
        remove_ = method_mock(request, EmbeddedFont, "remove")

        embedded_fonts.remove(font)

        remove_.assert_called_once_with(font)

    # fixtures -------------------------------------------------------

    def _two_font_prs_elm(self) -> CT_Presentation:
        return cast(
            CT_Presentation,
            element(
                "p:presentation/p:embeddedFontLst/("
                "p:embeddedFont/p:font{typeface=Calibri},"
                "p:embeddedFont/p:font{typeface=Arial}"
                ")"
            ),
        )

    @pytest.fixture
    def part_(self, request):
        return instance_mock(request, PresentationPart)


class DescribeEmbeddedFont:
    """Unit-test suite for `pptx.fonts.EmbeddedFont` objects."""

    def it_knows_its_typeface(self, part_):
        entry = self._entry_elm("Calibri")
        font = EmbeddedFont(entry, part_)
        assert font.typeface == "Calibri"

    def it_knows_which_styles_are_embedded(self, part_):
        entry = self._entry_elm(
            "Calibri", regular="rId1", bold="rId2", italic="rId3", bold_italic="rId4"
        )
        font = EmbeddedFont(entry, part_)

        assert font.styles == ("regular", "bold", "italic", "bold_italic")
        assert font.has_regular is True
        assert font.has_bold is True
        assert font.has_italic is True
        assert font.has_bold_italic is True

    def it_knows_when_a_style_is_not_embedded(self, part_):
        entry = self._entry_elm("Calibri", regular="rId1")
        font = EmbeddedFont(entry, part_)

        assert font.styles == ("regular",)
        assert font.has_bold is False
        assert font.has_italic is False
        assert font.has_bold_italic is False

    def it_supports_equality(self, part_):
        entry = self._entry_elm("Calibri", regular="rId1")
        other_entry = self._entry_elm("Arial", regular="rId2")

        assert EmbeddedFont(entry, part_) == EmbeddedFont(entry, part_)
        assert EmbeddedFont(entry, part_) != EmbeddedFont(other_entry, part_)
        assert EmbeddedFont(entry, part_) != "not-a-font"

    def it_drops_the_relationship_for_each_embedded_style_on_remove(self, part_):
        entry = self._entry_elm(
            "Calibri", regular="rId1", bold="rId2", italic="rId3", bold_italic="rId4"
        )
        font = EmbeddedFont(entry, part_)

        font.remove()

        part_.drop_rel.assert_any_call("rId1")
        part_.drop_rel.assert_any_call("rId2")
        part_.drop_rel.assert_any_call("rId3")
        part_.drop_rel.assert_any_call("rId4")
        assert part_.drop_rel.call_count == 4

    def it_removes_itself_from_the_embeddedFontLst(self, part_):
        prs_elm = cast(
            CT_Presentation,
            element(
                "p:presentation/p:embeddedFontLst/("
                "p:embeddedFont/p:font{typeface=Calibri},"
                "p:embeddedFont/p:font{typeface=Arial}"
                ")"
            ),
        )
        embeddedFontLst = prs_elm.embeddedFontLst
        assert embeddedFontLst is not None
        entry = embeddedFontLst.embeddedFont_lst[0]
        font = EmbeddedFont(entry, part_)

        font.remove()

        assert len(embeddedFontLst.embeddedFont_lst) == 1
        assert embeddedFontLst.embeddedFont_lst[0].font.typeface == "Arial"
        # -- list still has an entry, so embedTrueTypeFonts is left alone --
        assert prs_elm.embedTrueTypeFonts is None

    def it_removes_the_embeddedFontLst_when_the_last_font_is_removed(self, part_):
        prs_elm = cast(
            CT_Presentation,
            element(
                "p:presentation{embedTrueTypeFonts=1}/p:embeddedFontLst/"
                "p:embeddedFont/p:font{typeface=Calibri}"
            ),
        )
        embeddedFontLst = prs_elm.embeddedFontLst
        assert embeddedFontLst is not None
        entry = embeddedFontLst.embeddedFont_lst[0]
        font = EmbeddedFont(entry, part_)

        font.remove()

        assert prs_elm.embeddedFontLst is None
        assert prs_elm.embedTrueTypeFonts is False

    # fixtures -------------------------------------------------------

    def _entry_elm(self, typeface: str, **styles: str):
        cxml_parts = [f"p:font{{typeface={typeface}}}"]
        tag_for_style = {
            "regular": "p:regular",
            "bold": "p:bold",
            "italic": "p:italic",
            "bold_italic": "p:boldItalic",
        }
        for style, rId in styles.items():
            cxml_parts.append(f"{tag_for_style[style]}{{r:id={rId}}}")
        cxml = "p:presentation/p:embeddedFontLst/p:embeddedFont/(%s)" % ",".join(cxml_parts)
        prs_elm = cast(CT_Presentation, element(cxml))
        embeddedFontLst = prs_elm.embeddedFontLst
        assert embeddedFontLst is not None
        return embeddedFontLst.embeddedFont_lst[0]

    @pytest.fixture
    def part_(self, request):
        return instance_mock(request, PresentationPart)
