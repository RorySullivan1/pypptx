# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.oxml.embeddedfont` module."""

from __future__ import annotations

from typing import cast

from pptx.oxml.embeddedfont import (
    CT_EmbeddedFontDataId,
    CT_EmbeddedFontList,
    CT_EmbeddedFontListEntry,
    CT_Font,
)

from ..unitutil.cxml import element, xml


class DescribeCT_EmbeddedFontList:
    """Unit-test suite for `pptx.oxml.embeddedfont.CT_EmbeddedFontList` objects."""

    def it_can_add_an_embeddedFont_entry(self):
        embeddedFontLst = cast(CT_EmbeddedFontList, element("p:embeddedFontLst"))

        entry = embeddedFontLst.add_embeddedFont("Calibri")

        assert embeddedFontLst.xml == xml(
            'p:embeddedFontLst/p:embeddedFont/p:font{typeface=Calibri}'
        )
        assert entry.font.typeface == "Calibri"

    def it_provides_access_to_its_embeddedFont_entries(self):
        embeddedFontLst = cast(
            CT_EmbeddedFontList,
            element(
                "p:embeddedFontLst/("
                "p:embeddedFont/p:font{typeface=Calibri},"
                "p:embeddedFont/p:font{typeface=Arial}"
                ")"
            ),
        )

        typefaces = [entry.font.typeface for entry in embeddedFontLst.embeddedFont_lst]

        assert typefaces == ["Calibri", "Arial"]


class DescribeCT_EmbeddedFontListEntry:
    """Unit-test suite for `pptx.oxml.embeddedfont.CT_EmbeddedFontListEntry` objects."""

    def it_can_construct_a_new_entry(self):
        entry = CT_EmbeddedFontListEntry.new("Calibri")

        assert entry.xml == xml("p:embeddedFont/p:font{typeface=Calibri}")
        assert entry.font.typeface == "Calibri"
        assert entry.regular is None
        assert entry.bold is None
        assert entry.italic is None
        assert entry.boldItalic is None

    def it_can_add_a_style_variant_in_schema_order(self):
        # -- seed with an r:id-bearing child so the "r" namespace prefix is bound at the
        # -- root and later insertions reuse it rather than minting a fresh alias --
        entry = cast(
            CT_EmbeddedFontListEntry,
            element("p:embeddedFont/(p:font{typeface=Calibri},p:boldItalic{r:id=rId4})"),
        )
        entry._remove_boldItalic()

        entry.add_style("bold_italic", "rId4")
        entry.add_style("regular", "rId1")
        entry.add_style("bold", "rId2")
        entry.add_style("italic", "rId3")

        assert entry.xml == xml(
            "p:embeddedFont/("
            "p:font{typeface=Calibri},"
            "p:regular{r:id=rId1},"
            "p:bold{r:id=rId2},"
            "p:italic{r:id=rId3},"
            "p:boldItalic{r:id=rId4}"
            ")"
        )

    def it_can_iterate_its_style_rIds(self):
        entry = cast(
            CT_EmbeddedFontListEntry,
            element(
                "p:embeddedFont/("
                "p:font{typeface=Calibri},"
                "p:regular{r:id=rId1},"
                "p:italic{r:id=rId3}"
                ")"
            ),
        )

        assert list(entry.iter_style_rIds()) == [("regular", "rId1"), ("italic", "rId3")]


class DescribeCT_Font:
    """Unit-test suite for `pptx.oxml.embeddedfont.CT_Font` objects."""

    def it_provides_access_to_its_typeface(self):
        font = cast(CT_Font, element("p:font{typeface=Calibri}"))
        assert font.typeface == "Calibri"

    def it_provides_access_to_its_optional_attributes(self):
        font = cast(
            CT_Font,
            element("p:font{typeface=Calibri,panose=20020602040504020204,pitchFamily=34,charset=0}"),
        )

        assert font.panose == "20020602040504020204"
        assert font.pitchFamily == 34
        assert font.charset == 0

    def it_defaults_its_optional_attributes_to_None(self):
        font = cast(CT_Font, element("p:font{typeface=Calibri}"))
        assert font.panose is None
        assert font.pitchFamily is None
        assert font.charset is None


class DescribeCT_EmbeddedFontDataId:
    """Unit-test suite for `pptx.oxml.embeddedfont.CT_EmbeddedFontDataId` objects."""

    def it_provides_access_to_its_rId(self):
        regular = cast(CT_EmbeddedFontDataId, element("p:regular{r:id=rId1}"))
        assert regular.rId == "rId1"
