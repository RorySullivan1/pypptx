"""Unit-test suite for `pptx.custom_show` module."""

from __future__ import annotations

import pytest

from pptx.custom_show import CustomShow, CustomShows
from pptx.exc import SlideError
from pptx.parts.presentation import PresentationPart
from pptx.slide import Slide

from .unitutil.cxml import element, xml
from .unitutil.mock import Mock, instance_mock


class DescribeCustomShows:
    """Unit-test suite for `pptx.custom_show.CustomShows` objects."""

    def it_supports_len(self):
        custom_shows = CustomShows(
            element(
                "p:presentation/p:custShowLst/(p:custShow{name=A,id=0},"
                "p:custShow{name=B,id=1})"
            ),
            None,
        )
        assert len(custom_shows) == 2

    def it_supports_len_with_no_custom_shows(self):
        custom_shows = CustomShows(element("p:presentation"), None)
        assert len(custom_shows) == 0

    def it_supports_indexed_access(self):
        custom_shows = CustomShows(
            element(
                "p:presentation/p:custShowLst/(p:custShow{name=A,id=0},"
                "p:custShow{name=B,id=1})"
            ),
            None,
        )
        assert custom_shows[0].name == "A"
        assert custom_shows[1].name == "B"

    def it_raises_on_index_out_of_range(self):
        custom_shows = CustomShows(element("p:presentation"), None)
        with pytest.raises(IndexError):
            custom_shows[0]

    def it_supports_iteration(self):
        custom_shows = CustomShows(
            element(
                "p:presentation/p:custShowLst/(p:custShow{name=A,id=0},"
                "p:custShow{name=B,id=1})"
            ),
            None,
        )
        names = [cs.name for cs in custom_shows]
        assert names == ["A", "B"]

    def it_can_add_a_custom_show(self, add_fixture):
        custom_shows, prs_elm, slide_1, slide_3 = add_fixture

        custom_show = custom_shows.add("My Show", [slide_1, slide_3])

        assert len(custom_shows) == 1
        assert custom_show.name == "My Show"
        assert custom_show.id == 0
        assert custom_show.slides == [slide_1, slide_3]
        assert prs_elm.xpath("p:custShowLst/p:custShow/p:sldLst/p:sld/@r:id") == [
            "rId1",
            "rId3",
        ]

    def it_allocates_ids_starting_from_the_max_plus_one(self, add_fixture):
        custom_shows, _, slide_1, slide_3 = add_fixture
        custom_shows.add("First", [slide_1])

        second = custom_shows.add("Second", [slide_3])

        assert second.id == 1

    def it_raises_adding_a_show_with_a_foreign_slide(self, add_fixture):
        custom_shows, _, _, _ = add_fixture
        foreign_slide = Slide(element("p:sld"), None)

        with pytest.raises(SlideError):
            custom_shows.add("Bad", [foreign_slide])

    def it_can_get_a_custom_show_by_name(self, add_fixture):
        custom_shows, _, slide_1, _ = add_fixture
        custom_shows.add("My Show", [slide_1])

        found = custom_shows.get("My Show")
        missing = custom_shows.get("Nope", "default")

        assert found is not None and found.name == "My Show"
        assert missing == "default"

    def it_can_remove_a_custom_show(self, add_fixture):
        custom_shows, _, slide_1, _ = add_fixture
        custom_show = custom_shows.add("My Show", [slide_1])

        custom_shows.remove(custom_show)

        assert len(custom_shows) == 0

    def it_raises_removing_a_custom_show_not_in_the_collection(self, add_fixture):
        custom_shows, _, _, _ = add_fixture
        foreign_show = CustomShow(element("p:custShow{name=Foreign,id=9}"), None)

        with pytest.raises(SlideError):
            custom_shows.remove(foreign_show)

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def add_fixture(self, request, prs_part_):
        prs_ = Mock(part=prs_part_)
        prs_elm = element("p:presentation")
        custom_shows = CustomShows(prs_elm, prs_)

        slide_1_part, slide_3_part = object(), object()
        slide_1 = Slide(element("p:sld"), slide_1_part)
        slide_3 = Slide(element("p:sld"), slide_3_part)
        rIds = {slide_1_part: "rId1", slide_3_part: "rId3"}
        prs_part_.rId_for_slide.side_effect = lambda slide_part: rIds.get(slide_part)
        prs_part_.related_slide.side_effect = lambda rId: {
            "rId1": slide_1,
            "rId3": slide_3,
        }[rId]
        return custom_shows, prs_elm, slide_1, slide_3

    @pytest.fixture
    def prs_part_(self, request):
        return instance_mock(request, PresentationPart)


class DescribeCustomShow:
    """Unit-test suite for `pptx.custom_show.CustomShow` objects."""

    def it_knows_its_name(self):
        custom_show = CustomShow(element("p:custShow{name=Intro,id=0}"), None)
        assert custom_show.name == "Intro"

    def it_can_change_its_name(self):
        custShow_elm = element("p:custShow{name=Old,id=0}")
        custom_show = CustomShow(custShow_elm, None)
        custom_show.name = "New"
        assert custom_show.name == "New"
        assert custShow_elm.xml == xml("p:custShow{name=New,id=0}")

    def it_knows_its_id(self):
        custom_show = CustomShow(element("p:custShow{name=A,id=7}"), None)
        assert custom_show.id == 7

    def it_knows_its_slides_are_empty_with_no_sldLst(self):
        custom_show = CustomShow(element("p:custShow{name=A,id=0}"), None)
        assert custom_show.slides == []
