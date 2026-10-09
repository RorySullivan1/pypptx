"""Custom element classes for slide transitions (`p:transition`)."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.enum.animation import PP_TRANSITION_SPEED, PP_TRANSITION_TYPE
from pptx.oxml.ns import namespaces, qn
from pptx.oxml.simpletypes import XsdBoolean, XsdUnsignedInt
from pptx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrOne

if TYPE_CHECKING:
    from lxml.etree import _Element

# -- the Clark-notation tag of each transition-effect element, mapped to its type --
_EFFECT_TYPES = {
    qn(member.xml_value): member for member in PP_TRANSITION_TYPE if member.xml_value
}

# -- an `mc:Choice` is read when every namespace it requires is one of these; pypptx reads the
# -- transition markup of PowerPoint 2010 (p14), 2013 (p15) and 2016+ Morph (p159) --
_READABLE_NAMESPACES = frozenset(namespaces("p14", "p15", "p159").values())


class CT_SlideTransition(BaseOxmlElement):
    """`p:transition` element, the transition into a slide.

    Its one optional effect child names the visual effect; `p:sndAc` (a sound) and `p:extLst`
    may follow it.
    """

    spd: PP_TRANSITION_SPEED = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "spd", PP_TRANSITION_SPEED, default=PP_TRANSITION_SPEED.FAST
    )
    advClick: bool = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "advClick", XsdBoolean, default=True
    )
    advTm: int | None = OptionalAttribute("advTm", XsdUnsignedInt)  # pyright: ignore
    dur: int | None = OptionalAttribute("p14:dur", XsdUnsignedInt)  # pyright: ignore
    sndAc: BaseOxmlElement | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "p:sndAc", successors=("p:extLst",)
    )

    @property
    def effect(self) -> _Element | None:
        """The effect child element, e.g. `p:fade` or `p159:morph`, or None when there is none."""
        for child in self.iterchildren():
            if child.tag in _EFFECT_TYPES:
                return child
        return None

    @property
    def effect_type(self) -> PP_TRANSITION_TYPE:
        """The effect named by the effect child, `PP_TRANSITION_TYPE.NONE` when there is none."""
        effect = self.effect
        return PP_TRANSITION_TYPE.NONE if effect is None else _EFFECT_TYPES[effect.tag]


def effective_transition(owner: BaseOxmlElement) -> CT_SlideTransition | None:
    """The `p:transition` that applies for `owner`, a `p:sld`, `p:sldLayout` or `p:sldMaster`.

    PowerPoint 2010 and later write a transition that uses newer markup (a `p14:dur`
    duration, a PowerPoint 2010 effect, Morph) inside `mc:AlternateContent`, with an ECMA-376
    equivalent as its `mc:Fallback`. The richer `mc:Choice` is used when pypptx reads every
    namespace it requires; otherwise the fallback is. None when there is no transition.
    """
    for child in owner.iterchildren():
        if child.tag == qn("p:transition"):
            return child  # pyright: ignore[reportReturnType]
        if child.tag == qn("mc:AlternateContent"):
            transition = _transition_in_alternate_content(child)
            if transition is not None:
                return transition
    return None


def _transition_in_alternate_content(alternateContent: _Element) -> CT_SlideTransition | None:
    """The `p:transition` in the readable `mc:Choice` of `alternateContent`, else its fallback."""
    for choice in alternateContent.iterchildren(qn("mc:Choice")):
        prefixes = (choice.get("Requires") or "").split()
        if prefixes and all(choice.nsmap.get(p) in _READABLE_NAMESPACES for p in prefixes):
            transition = next(choice.iterchildren(qn("p:transition")), None)
            if transition is not None:
                return transition  # pyright: ignore[reportReturnType]
    for fallback in alternateContent.iterchildren(qn("mc:Fallback")):
        transition = next(fallback.iterchildren(qn("p:transition")), None)
        if transition is not None:
            return transition  # pyright: ignore[reportReturnType]
    return None
