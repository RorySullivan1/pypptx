"""Unit-test suite for the `pptx.oxml.transition` module."""

from __future__ import annotations

import pytest

from pptx.enum.animation import PP_TRANSITION_SPEED, PP_TRANSITION_TYPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.transition import CT_SlideTransition, effective_transition

from ..unitutil.cxml import element


class DescribeCT_SlideTransition:
    """Unit-test suite for `pptx.oxml.transition.CT_SlideTransition` objects."""

    def it_is_the_element_class_for_a_transition(self):
        assert isinstance(element("p:transition"), CT_SlideTransition)

    @pytest.mark.parametrize(
        ("cxml", "spd", "advClick", "advTm", "dur"),
        [
            ("p:transition", PP_TRANSITION_SPEED.FAST, True, None, None),
            (
                "p:transition{spd=slow,advClick=0,advTm=3000}",
                PP_TRANSITION_SPEED.SLOW,
                False,
                3000,
                None,
            ),
            ("p:transition{spd=med}", PP_TRANSITION_SPEED.MEDIUM, True, None, None),
        ],
    )
    def it_knows_its_timing_attributes(self, cxml, spd, advClick, advTm, dur):
        transition = element(cxml)

        assert transition.spd == spd
        assert transition.advClick is advClick
        assert transition.advTm == advTm
        assert transition.dur == dur

    def it_knows_its_PowerPoint_2010_duration(self):
        transition = parse_xml(f'<p:transition {nsdecls("p", "p14")} p14:dur="2000"/>')

        assert transition.dur == 2000

    @pytest.mark.parametrize(
        ("cxml", "expected_type"),
        [
            ("p:transition", PP_TRANSITION_TYPE.NONE),
            ("p:transition/p:sndAc", PP_TRANSITION_TYPE.NONE),
            ("p:transition/p:fade", PP_TRANSITION_TYPE.FADE),
            ("p:transition/(p:randomBar{dir=vert},p:sndAc)", PP_TRANSITION_TYPE.RANDOM_BAR),
            ("p:transition/p14:vortex{dir=r}", PP_TRANSITION_TYPE.VORTEX),
            ("p:transition/p15:prstTrans{prst=curtains}", PP_TRANSITION_TYPE.PRESET),
            ("p:transition/p159:morph{option=byObject}", PP_TRANSITION_TYPE.MORPH),
        ],
    )
    def it_knows_its_effect(self, cxml, expected_type):
        transition = element(cxml)

        assert transition.effect_type == expected_type
        if expected_type == PP_TRANSITION_TYPE.NONE:
            assert transition.effect is None
        else:
            assert transition.effect is not None

    def it_provides_access_to_its_sound_action(self):
        assert element("p:transition/(p:fade,p:sndAc)").sndAc is not None
        assert element("p:transition/p:fade").sndAc is None


def _sld_with(content: str):
    return parse_xml(
        f"<p:sld {nsdecls('p', 'mc', 'p14', 'p159')}>"
        "<p:cSld><p:spTree/></p:cSld>"
        f"{content}"
        "</p:sld>"
    )


class Describe_effective_transition:
    """Unit-test suite for `pptx.oxml.transition.effective_transition()`."""

    def it_is_None_when_there_is_no_transition(self):
        assert effective_transition(_sld_with("")) is None

    def it_finds_a_plain_transition(self):
        sld = _sld_with('<p:transition spd="slow"><p:zoom/></p:transition>')

        transition = effective_transition(sld)

        assert transition is not None
        assert transition.effect_type == PP_TRANSITION_TYPE.ZOOM

    @pytest.mark.parametrize(
        ("requires", "effect", "expected_type", "expected_dur"),
        [
            ("p14", '<p14:vortex dir="r"/>', PP_TRANSITION_TYPE.VORTEX, 2000),
            ("p159", '<p159:morph option="byObject"/>', PP_TRANSITION_TYPE.MORPH, 2000),
        ],
    )
    def it_reads_the_choice_of_an_alternate_content_it_understands(
        self, requires, effect, expected_type, expected_dur
    ):
        sld = _sld_with(
            "<mc:AlternateContent>"
            f'<mc:Choice Requires="{requires}">'
            f'<p:transition spd="slow" p14:dur="2000">{effect}</p:transition>'
            "</mc:Choice>"
            '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback>'
            "</mc:AlternateContent>"
        )

        transition = effective_transition(sld)

        assert transition is not None
        assert transition.effect_type == expected_type
        assert transition.dur == expected_dur

    def but_it_reads_the_fallback_when_the_choice_requires_an_unknown_namespace(self):
        sld = parse_xml(
            f"<p:sld {nsdecls('p', 'mc')} xmlns:x9='http://example.com/future'>"
            "<p:cSld><p:spTree/></p:cSld>"
            "<mc:AlternateContent>"
            '<mc:Choice Requires="x9"><p:transition><x9:sparkle/></p:transition></mc:Choice>'
            '<mc:Fallback><p:transition spd="med"><p:fade/></p:transition></mc:Fallback>'
            "</mc:AlternateContent>"
            "</p:sld>"
        )

        transition = effective_transition(sld)

        assert transition is not None
        assert transition.effect_type == PP_TRANSITION_TYPE.FADE
        assert transition.spd == PP_TRANSITION_SPEED.MEDIUM

    def and_it_ignores_alternate_content_that_holds_no_transition(self):
        sld = _sld_with(
            '<mc:AlternateContent><mc:Choice Requires="p14"><p:timing/></mc:Choice>'
            "</mc:AlternateContent>"
            '<p:transition><p:wipe dir="d"/></p:transition>'
        )

        transition = effective_transition(sld)

        assert transition is not None
        assert transition.effect_type == PP_TRANSITION_TYPE.WIPE

    def it_works_for_a_layout_and_a_master_too(self):
        for root in ("p:sldLayout", "p:sldMaster"):
            owner = element(f"{root}/(p:cSld/p:spTree,p:transition/p:cut)")

            transition = effective_transition(owner)

            assert transition is not None
            assert transition.effect_type == PP_TRANSITION_TYPE.CUT
