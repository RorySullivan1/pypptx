"""Read-only access to the transition into a slide."""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING

from pptx.enum.animation import PP_TRANSITION_SPEED, PP_TRANSITION_TYPE
from pptx.oxml.ns import qn
from pptx.oxml.transition import effective_transition
from pptx.shared import ElementProxy

if TYPE_CHECKING:
    from pptx.oxml.transition import CT_SlideTransition


class SlideTransition(ElementProxy):
    """The transition PowerPoint plays when it moves to a slide, read only.

    Obtained from `Slide.transition`, and also from `SlideLayout.transition` and
    `SlideMaster.transition`, which report the transition stored on the layout or master
    itself. pypptx does not work out whether PowerPoint applies a layout's or master's
    transition to a slide that has none of its own.

    A transition PowerPoint 2010 or later wrote in `mc:AlternateContent` is read from its richer
    `mc:Choice` (which carries the duration and the newer effects) rather than its fallback.
    With no transition, `type` is `PP_TRANSITION_TYPE.NONE`, `advance_on_click` is |True| and
    the other properties are |None|. Reading never changes the XML.
    """

    @property
    def advance_after(self) -> dt.timedelta | None:
        """Time after which the slide advances on its own, or |None| when it doesn't."""
        transition = self._transition
        if transition is None or transition.advTm is None:
            return None
        return dt.timedelta(milliseconds=transition.advTm)

    @property
    def advance_on_click(self) -> bool:
        """|True| when a mouse click advances to the next slide, the default."""
        transition = self._transition
        return True if transition is None else transition.advClick

    @property
    def direction(self) -> str | None:
        """The direction or orientation of the effect as it is stored, e.g. `"l"` or `"vert"`.

        This is the effect element's `dir` attribute, or its `orient` attribute for effects such
        as blinds that run horizontally or vertically. |None| when the effect has neither, which
        means the effect's default direction.
        """
        effect = self._effect
        if effect is None:
            return None
        return effect.get("dir", effect.get("orient"))

    @property
    def duration(self) -> dt.timedelta | None:
        """How long the effect takes, or |None| when the transition does not say.

        PowerPoint 2010 and later store an exact duration (`p14:dur`); older files store only
        `speed`.
        """
        transition = self._transition
        if transition is None or transition.dur is None:
            return None
        return dt.timedelta(milliseconds=transition.dur)

    @property
    def has_sound(self) -> bool:
        """|True| when the transition starts a sound."""
        transition = self._transition
        if transition is None or transition.sndAc is None:
            return False
        return transition.sndAc.find(qn("p:stSnd")) is not None

    @property
    def preset(self) -> str | None:
        """The name of a PowerPoint 2013 preset transition, e.g. `"curtains"`.

        Set only when `type` is `PP_TRANSITION_TYPE.PRESET`; |None| otherwise.
        """
        effect = self._effect
        if effect is None or effect.tag != qn("p15:prstTrans"):
            return None
        return effect.get("prst")

    @property
    def speed(self) -> PP_TRANSITION_SPEED | None:
        """The transition's speed, `PP_TRANSITION_SPEED.FAST` when it states none.

        |None| when there is no transition.
        """
        transition = self._transition
        return None if transition is None else transition.spd

    @property
    def type(self) -> PP_TRANSITION_TYPE:
        """The visual effect, a member of `PP_TRANSITION_TYPE`.

        `PP_TRANSITION_TYPE.NONE` when there is no transition, or one with timing but no effect.
        """
        transition = self._transition
        return PP_TRANSITION_TYPE.NONE if transition is None else transition.effect_type

    @property
    def _effect(self):
        transition = self._transition
        return None if transition is None else transition.effect

    @property
    def _transition(self) -> CT_SlideTransition | None:
        return effective_transition(self._element)
