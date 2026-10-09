"""Unit-test suite for the `pptx.transition` module."""

from __future__ import annotations

import datetime as dt

import pytest

from pptx.enum.animation import PP_TRANSITION_SPEED, PP_TRANSITION_TYPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.slide import Slide, SlideLayout, SlideMaster
from pptx.transition import SlideTransition

from .unitutil.cxml import element


def _slide_with(content: str) -> Slide:
    sld = parse_xml(
        f"<p:sld {nsdecls('p', 'mc', 'p14', 'p15', 'p159')}>"
        f"<p:cSld><p:spTree/></p:cSld>{content}</p:sld>"
    )
    return Slide(sld, None)  # pyright: ignore[reportArgumentType]


class DescribeSlideTransition:
    """Unit-test suite for `pptx.transition.SlideTransition` objects."""

    def it_is_provided_by_a_slide_a_layout_and_a_master(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition/p:fade)")
        sldLayout = element("p:sldLayout/(p:cSld/p:spTree,p:transition/p:cut)")
        sldMaster = element("p:sldMaster/(p:cSld/p:spTree,p:transition/p:wipe)")

        slide = Slide(sld, None)  # pyright: ignore[reportArgumentType]
        layout = SlideLayout(sldLayout, None)  # pyright: ignore[reportArgumentType]
        master = SlideMaster(sldMaster, None)  # pyright: ignore[reportArgumentType]

        assert isinstance(slide.transition, SlideTransition)
        assert slide.transition.type == PP_TRANSITION_TYPE.FADE
        assert layout.transition.type == PP_TRANSITION_TYPE.CUT
        assert master.transition.type == PP_TRANSITION_TYPE.WIPE

    def it_reports_no_transition_when_there_is_none(self):
        transition = _slide_with("").transition

        assert transition.type == PP_TRANSITION_TYPE.NONE
        assert transition.advance_on_click is True
        assert transition.speed is None
        assert transition.duration is None
        assert transition.advance_after is None
        assert transition.direction is None
        assert transition.preset is None
        assert transition.has_sound is False

    def it_reads_a_plain_transition(self):
        transition = _slide_with(
            '<p:transition spd="slow" advClick="0" advTm="5000">'
            '<p:blinds dir="vert"/><p:sndAc><p:stSnd><p:snd r:embed="rId3" name="x"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
            "</p:stSnd></p:sndAc></p:transition>"
        ).transition

        assert transition.type == PP_TRANSITION_TYPE.BLINDS
        assert transition.speed == PP_TRANSITION_SPEED.SLOW
        assert transition.advance_on_click is False
        assert transition.advance_after == dt.timedelta(seconds=5)
        assert transition.duration is None
        assert transition.direction == "vert"
        assert transition.has_sound is True

    def it_reads_a_transition_that_has_timing_but_no_effect(self):
        transition = _slide_with('<p:transition spd="slow"/>').transition

        assert transition.type == PP_TRANSITION_TYPE.NONE
        assert transition.speed == PP_TRANSITION_SPEED.SLOW

    def it_reads_the_PowerPoint_2010_choice_of_an_alternate_content_transition(self):
        transition = _slide_with(
            '<mc:AlternateContent><mc:Choice Requires="p14">'
            '<p:transition spd="slow" p14:dur="1250"><p14:vortex dir="r"/></p:transition>'
            '</mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition>'
            "</mc:Fallback></mc:AlternateContent>"
        ).transition

        assert transition.type == PP_TRANSITION_TYPE.VORTEX
        assert transition.duration == dt.timedelta(milliseconds=1250)
        assert transition.direction == "r"

    @pytest.mark.parametrize(
        ("effect", "expected_type", "expected_preset"),
        [
            ('<p15:prstTrans prst="curtains"/>', PP_TRANSITION_TYPE.PRESET, "curtains"),
            ('<p159:morph option="byObject"/>', PP_TRANSITION_TYPE.MORPH, None),
        ],
    )
    def it_names_a_preset_transition(self, effect, expected_type, expected_preset):
        transition = _slide_with(f"<p:transition>{effect}</p:transition>").transition

        assert transition.type == expected_type
        assert transition.preset == expected_preset

    def it_does_not_change_the_xml_when_read(self):
        slide = _slide_with('<p:transition spd="med"><p:push dir="u"/></p:transition>')
        before = slide._element.xml

        transition = slide.transition
        _ = (
            transition.type,
            transition.speed,
            transition.duration,
            transition.advance_after,
            transition.advance_on_click,
            transition.direction,
            transition.preset,
            transition.has_sound,
        )

        assert slide._element.xml == before
