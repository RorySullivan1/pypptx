"""Read-only access to the animation effects on a slide."""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING, Callable, Iterator, TypeVar

from pptx.enum.animation import (
    MSO_ANIMATION_EFFECT,
    MSO_ANIMATION_TRIGGER,
    PP_ANIMATION_CLASS,
)
from pptx.exc import InvalidXmlError
from pptx.oxml.simpletypes import ST_TLTime
from pptx.shared import ParentedElementProxy

if TYPE_CHECKING:
    from pptx.oxml.slide import CT_Slide
    from pptx.oxml.timing import (
        CT_TLCommonTimeNodeData,
        CT_TLTimeNodeContainer,
        CT_TLTimeTargetElement,
    )
    from pptx.shapes.base import BaseShape
    from pptx.shapes.shapetree import SlideShapes
    from pptx.slide import Slide

_T = TypeVar("_T")

_TRIGGER_BY_NODE_TYPE = {
    "clickEffect": MSO_ANIMATION_TRIGGER.ON_PAGE_CLICK,
    "withEffect": MSO_ANIMATION_TRIGGER.WITH_PREVIOUS,
    "afterEffect": MSO_ANIMATION_TRIGGER.AFTER_PREVIOUS,
}


class SlideAnimations(ParentedElementProxy):
    """The animation effects on a slide, in the order PowerPoint plays them. Read only.

    Obtained from `Slide.animations`. The effects of the main sequence come first, in play order,
    then those of each triggered sequence (started by clicking a particular shape). Supports
    `len()`, iteration and indexed access. Each effect is an |Animation|.

    The media nodes that keep a movie or sound playable are not effects and are not listed;
    "play", "pause" and "stop" effects are, with class `PP_ANIMATION_CLASS.MEDIA`.
    """

    _element: CT_Slide  # pyright: ignore[reportIncompatibleVariableOverride]

    def __init__(self, sld: CT_Slide, parent: Slide):
        super(SlideAnimations, self).__init__(sld, parent)
        self._slide = parent

    def __getitem__(self, idx: int) -> Animation:
        """The animation at `idx`, e.g. `slide.animations[0]`."""
        return list(self)[idx]

    def __iter__(self) -> Iterator[Animation]:
        """Generate each |Animation|, in play order."""
        for cTn in self._iter_effect_cTns():
            yield Animation(cTn, self._slide)

    def __len__(self) -> int:
        return sum(1 for _ in self._iter_effect_cTns())

    def _iter_effect_cTns(self) -> Iterator[CT_TLCommonTimeNodeData]:
        """Each effect's `p:cTn`: those of the main sequence, then of each triggered sequence."""
        timing = self._element.timing
        if timing is None:
            return
        effect_cTns = timing.xpath("./p:tnLst//p:cTn[@presetClass]")

        def sort_key(cTn: CT_TLCommonTimeNodeData) -> int:
            seq = _owning_sequence(cTn)
            if seq is None:
                return 2
            return 0 if _node_type(seq.cTn) == "mainSeq" else 1

        # -- `sorted()` is stable, so document order holds within the main sequence and across
        # -- the triggered sequences --
        yield from sorted(effect_cTns, key=sort_key)


class Animation(ParentedElementProxy):
    """One animation effect on a slide, read only.

    An effect is identified by its class (`preset_class`, e.g. entrance) together with its
    `preset_id` and `preset_subtype`, which PowerPoint numbers within each class. `effect_type`
    names the common entrance and exit effects.

    Times are `datetime.timedelta` objects; one that PowerPoint stores as "indefinite" (for
    example an effect that runs until the next click) is |None|.
    """

    _element: CT_TLCommonTimeNodeData  # pyright: ignore[reportIncompatibleVariableOverride]

    def __init__(self, cTn: CT_TLCommonTimeNodeData, slide: Slide):
        super(Animation, self).__init__(cTn, slide)
        self._slide = slide

    @property
    def delay(self) -> dt.timedelta | None:
        """How long after its trigger the effect starts; |None| when "indefinite"."""
        stCondLst = self._element.stCondLst
        conds = [] if stCondLst is None else stCondLst.cond_lst
        if not conds:
            return dt.timedelta(0)
        delay = _safe(lambda: conds[0].delay, None)
        if delay is None:
            return dt.timedelta(0)
        return _timedelta(delay)

    @property
    def duration(self) -> dt.timedelta | None:
        """How long one run of the effect takes, from its start to the end of its last behavior.

        |None| when a behavior runs "indefinitely" or no behavior states a duration. Repeats are
        not counted.
        """
        ends: list[int] = []
        for cTn in self._element.xpath(
            "./p:childTnLst//p:cBhvr/p:cTn | ./p:childTnLst//p:cMediaNode/p:cTn"
        ):
            dur = _safe(lambda: cTn.dur, None)
            if dur is None:
                continue
            if dur == ST_TLTime.INDEFINITE:
                return None
            stCondLst = cTn.stCondLst
            conds = [] if stCondLst is None else stCondLst.cond_lst
            delay = _safe(lambda: conds[0].delay, 0) if conds else 0
            start = delay if isinstance(delay, int) else 0
            ends.append(start + dur)
        if not ends:
            return None
        return dt.timedelta(milliseconds=max(ends))

    @property
    def effect_type(self) -> MSO_ANIMATION_EFFECT | None:
        """The named entrance or exit effect, e.g. `MSO_ANIMATION_EFFECT.FADE`.

        Set only for entrance and exit effects with a `preset_id` from 1 to 31, where
        PowerPoint's file format and its `MsoAnimEffect` enumeration agree; |None| otherwise.
        """
        if self.preset_class not in (PP_ANIMATION_CLASS.ENTRANCE, PP_ANIMATION_CLASS.EXIT):
            return None
        preset_id = self.preset_id
        if preset_id is None or not 1 <= preset_id <= 31:
            return None
        return MSO_ANIMATION_EFFECT(preset_id)

    @property
    def paragraphs(self) -> range | None:
        """The paragraphs the effect animates, when it animates part of a shape's text.

        A `range` of zero-based paragraph indexes, e.g. `range(0, 1)` for the first paragraph;
        |None| when the effect animates the whole shape.
        """
        tgtEl = self._target
        spTgt = None if tgtEl is None else tgtEl.spTgt
        txEl = None if spTgt is None else spTgt.txEl
        pRg = None if txEl is None else txEl.pRg
        if pRg is None:
            return None
        return _safe(lambda: range(pRg.st, pRg.end + 1), None)

    @property
    def preset_class(self) -> PP_ANIMATION_CLASS | None:
        """The kind of effect, e.g. `PP_ANIMATION_CLASS.ENTRANCE`."""
        presetClass = _safe(lambda: self._element.presetClass, None)
        if presetClass is None:
            return None
        return _safe(lambda: PP_ANIMATION_CLASS.from_xml(presetClass), None)

    @property
    def preset_id(self) -> int | None:
        """PowerPoint's number for the effect within its class, e.g. 10 for an entrance Fade."""
        return _safe(lambda: self._element.presetID, None)

    @property
    def preset_subtype(self) -> int | None:
        """PowerPoint's number for the effect's variant, such as its direction, or |None|."""
        return _safe(lambda: self._element.presetSubtype, None)

    @property
    def shape(self) -> BaseShape | None:
        """The shape the effect animates, or |None| when its target is not a shape on the slide."""
        shape_id = self.shape_id
        return None if shape_id is None else _find_shape(self._slide.shapes, shape_id)

    @property
    def shape_id(self) -> int | None:
        """The id of the shape the effect animates, or |None| when its target is not a shape."""
        tgtEl = self._target
        spTgt = None if tgtEl is None else tgtEl.spTgt
        return None if spTgt is None else _safe(lambda: spTgt.spid, None)

    @property
    def trigger(self) -> MSO_ANIMATION_TRIGGER:
        """What starts the effect, e.g. `MSO_ANIMATION_TRIGGER.ON_PAGE_CLICK`.

        An effect in a triggered sequence is `ON_SHAPE_CLICK` (`trigger_shape` is the shape), or
        `ON_MEDIA_BOOKMARK` when a media bookmark starts it. An effect that doesn't say how it
        starts is taken as `ON_PAGE_CLICK`.
        """
        seq = _owning_sequence(self._element)
        if seq is not None and _node_type(seq.cTn) == "interactiveSeq":
            evts = [cond.evt for cond in _start_conditions(seq.cTn)]
            if "onMediaBookmark" in evts:
                return MSO_ANIMATION_TRIGGER.ON_MEDIA_BOOKMARK
            return MSO_ANIMATION_TRIGGER.ON_SHAPE_CLICK
        return _TRIGGER_BY_NODE_TYPE.get(
            _node_type(self._element) or "", MSO_ANIMATION_TRIGGER.ON_PAGE_CLICK
        )

    @property
    def trigger_shape(self) -> BaseShape | None:
        """The shape whose click starts the effect, for an effect in a triggered sequence."""
        shape_id = self.trigger_shape_id
        return None if shape_id is None else _find_shape(self._slide.shapes, shape_id)

    @property
    def trigger_shape_id(self) -> int | None:
        """The id of the shape whose click starts the effect, or |None|."""
        seq = _owning_sequence(self._element)
        if seq is None or _node_type(seq.cTn) != "interactiveSeq":
            return None
        for cond in _start_conditions(seq.cTn):
            tgtEl = cond.tgtEl
            spTgt = None if tgtEl is None else tgtEl.spTgt
            if spTgt is not None:
                return _safe(lambda: spTgt.spid, None)
        return None

    @property
    def _target(self) -> CT_TLTimeTargetElement | None:
        """The target of the effect's first behavior or media node."""
        tgtEls = self._element.xpath(
            "./p:childTnLst//p:cBhvr/p:tgtEl | ./p:childTnLst//p:cMediaNode/p:tgtEl"
        )
        return tgtEls[0] if tgtEls else None


def _find_shape(shapes: SlideShapes, shape_id: int) -> BaseShape | None:
    """The shape with `shape_id` in `shapes`, looking inside group shapes too."""
    shape = shapes.get_by_id(shape_id)
    if shape is not None:
        return shape
    for candidate in shapes:
        group_shapes = getattr(candidate, "shapes", None)
        if group_shapes is not None:
            shape = _find_shape(group_shapes, shape_id)
            if shape is not None:
                return shape
    return None


def _node_type(cTn: CT_TLCommonTimeNodeData) -> str | None:
    return _safe(lambda: cTn.nodeType, None)


def _owning_sequence(cTn: CT_TLCommonTimeNodeData) -> CT_TLTimeNodeContainer | None:
    """The nearest `p:seq` that contains `cTn`, or None."""
    seqs = cTn.xpath("ancestor::p:seq[1]")
    return seqs[0] if seqs else None


def _safe(read: Callable[[], _T], default: _T) -> _T:
    """`read()`, or `default` when the XML holds a value the schema doesn't allow.

    Animation markup that a tool other than PowerPoint wrote is read as far as it makes sense,
    rather than raising.
    """
    try:
        return read()
    except (ValueError, InvalidXmlError):
        return default


def _start_conditions(cTn: CT_TLCommonTimeNodeData):
    stCondLst = cTn.stCondLst
    return [] if stCondLst is None else stCondLst.cond_lst


def _timedelta(milliseconds: int | str) -> dt.timedelta | None:
    if milliseconds == ST_TLTime.INDEFINITE or not isinstance(milliseconds, int):
        return None
    return dt.timedelta(milliseconds=milliseconds)
