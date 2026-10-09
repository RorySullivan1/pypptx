"""Enumerations used by slide transitions and animations."""

from __future__ import annotations

from pptx.enum.base import BaseEnum, BaseXmlEnum


class PP_TRANSITION_SPEED(BaseXmlEnum):
    """Specifies the speed of a slide transition.

    Example::

        from pptx.enum.animation import PP_TRANSITION_SPEED

        assert slide.transition.speed == PP_TRANSITION_SPEED.FAST

    MS API name: `PpTransitionSpeed`

    https://learn.microsoft.com/office/vba/api/powerpoint.pptransitionspeed
    """

    SLOW = (1, "slow", "Slow.")
    """Slow."""

    MEDIUM = (2, "med", "Medium.")
    """Medium."""

    FAST = (3, "fast", "Fast. The default when a transition specifies no speed.")
    """Fast. The default when a transition specifies no speed."""


class PP_TRANSITION_TYPE(BaseXmlEnum):
    """Specifies the visual effect of a slide transition.

    There is one member for each transition element PowerPoint writes: the ECMA-376 effects,
    the PowerPoint 2010 effects (`p14`), the PowerPoint 2013 preset transitions (`p15`) and
    Morph (`p159`). The XML value is the element's namespace-prefixed tag.

    PowerPoint's own `PpEntryEffect` combines an effect with its direction, so it has no
    one-to-one counterpart here; the values of this enumeration are pypptx's own. A
    transition's direction is available separately, from ``SlideTransition.direction``.

    Example::

        from pptx.enum.animation import PP_TRANSITION_TYPE

        if slide.transition.type == PP_TRANSITION_TYPE.MORPH:
            print("this slide morphs in")
    """

    NONE = (0, "", "No visual effect; the slide replaces the previous one at once.")
    """No visual effect; the slide replaces the previous one at once."""

    BLINDS = (1, "p:blinds", "Blinds.")
    """Blinds."""

    CHECKER = (2, "p:checker", "Checkerboard.")
    """Checkerboard."""

    CIRCLE = (3, "p:circle", "Circle.")
    """Circle."""

    COMB = (4, "p:comb", "Comb.")
    """Comb."""

    COVER = (5, "p:cover", "Cover.")
    """Cover."""

    CUT = (6, "p:cut", "Cut.")
    """Cut."""

    DIAMOND = (7, "p:diamond", "Diamond.")
    """Diamond."""

    DISSOLVE = (8, "p:dissolve", "Dissolve.")
    """Dissolve."""

    FADE = (9, "p:fade", "Fade.")
    """Fade."""

    NEWSFLASH = (10, "p:newsflash", "Newsflash.")
    """Newsflash."""

    PLUS = (11, "p:plus", "Plus.")
    """Plus."""

    PULL = (12, "p:pull", "Pull (uncover).")
    """Pull (uncover)."""

    PUSH = (13, "p:push", "Push.")
    """Push."""

    RANDOM = (14, "p:random", "Random.")
    """Random."""

    RANDOM_BAR = (15, "p:randomBar", "Random bars.")
    """Random bars."""

    SPLIT = (16, "p:split", "Split.")
    """Split."""

    STRIPS = (17, "p:strips", "Strips.")
    """Strips."""

    WEDGE = (18, "p:wedge", "Wedge.")
    """Wedge."""

    WHEEL = (19, "p:wheel", "Wheel.")
    """Wheel."""

    WIPE = (20, "p:wipe", "Wipe.")
    """Wipe."""

    ZOOM = (21, "p:zoom", "Zoom.")
    """Zoom."""

    CONVEYOR = (22, "p14:conveyor", "Conveyor (PowerPoint 2010).")
    """Conveyor (PowerPoint 2010)."""

    DOORS = (23, "p14:doors", "Doors (PowerPoint 2010).")
    """Doors (PowerPoint 2010)."""

    FERRIS = (24, "p14:ferris", "Ferris wheel (PowerPoint 2010).")
    """Ferris wheel (PowerPoint 2010)."""

    FLASH = (25, "p14:flash", "Flash (PowerPoint 2010).")
    """Flash (PowerPoint 2010)."""

    FLIP = (26, "p14:flip", "Flip (PowerPoint 2010).")
    """Flip (PowerPoint 2010)."""

    FLYTHROUGH = (27, "p14:flythrough", "Fly through (PowerPoint 2010).")
    """Fly through (PowerPoint 2010)."""

    GALLERY = (28, "p14:gallery", "Gallery (PowerPoint 2010).")
    """Gallery (PowerPoint 2010)."""

    GLITTER = (29, "p14:glitter", "Glitter (PowerPoint 2010).")
    """Glitter (PowerPoint 2010)."""

    HONEYCOMB = (30, "p14:honeycomb", "Honeycomb (PowerPoint 2010).")
    """Honeycomb (PowerPoint 2010)."""

    PAN = (31, "p14:pan", "Pan (PowerPoint 2010).")
    """Pan (PowerPoint 2010)."""

    PRISM = (32, "p14:prism", "Prism; Cube and Box are variants (PowerPoint 2010).")
    """Prism; Cube and Box are variants (PowerPoint 2010)."""

    REVEAL = (33, "p14:reveal", "Reveal (PowerPoint 2010).")
    """Reveal (PowerPoint 2010)."""

    RIPPLE = (34, "p14:ripple", "Ripple (PowerPoint 2010).")
    """Ripple (PowerPoint 2010)."""

    SHRED = (35, "p14:shred", "Shred (PowerPoint 2010).")
    """Shred (PowerPoint 2010)."""

    SWITCH = (36, "p14:switch", "Switch (PowerPoint 2010).")
    """Switch (PowerPoint 2010)."""

    VORTEX = (37, "p14:vortex", "Vortex (PowerPoint 2010).")
    """Vortex (PowerPoint 2010)."""

    WARP = (38, "p14:warp", "Warp (PowerPoint 2010).")
    """Warp (PowerPoint 2010)."""

    WHEEL_REVERSE = (39, "p14:wheelReverse", "Reverse wheel (PowerPoint 2010).")
    """Reverse wheel (PowerPoint 2010)."""

    WINDOW = (40, "p14:window", "Window (PowerPoint 2010).")
    """Window (PowerPoint 2010)."""

    PRESET = (
        41,
        "p15:prstTrans",
        "A PowerPoint 2013 preset transition, named by ``SlideTransition.preset``.",
    )
    """A PowerPoint 2013 preset transition, named by ``SlideTransition.preset``."""

    MORPH = (42, "p159:morph", "Morph (PowerPoint 2016 and later).")
    """Morph (PowerPoint 2016 and later)."""


class PP_ANIMATION_CLASS(BaseXmlEnum):
    """Specifies the kind of an animation effect: entrance, exit, emphasis and so on.

    Read from `p:cTn@presetClass`. PowerPoint numbers its effects within each class, so an
    effect is identified by its class together with ``Animation.preset_id``.

    Example::

        from pptx.enum.animation import PP_ANIMATION_CLASS

        entrances = [a for a in slide.animations if a.preset_class == PP_ANIMATION_CLASS.ENTRANCE]
    """

    ENTRANCE = (1, "entr", "An entrance effect.")
    """An entrance effect."""

    EXIT = (2, "exit", "An exit effect.")
    """An exit effect."""

    EMPHASIS = (3, "emph", "An emphasis effect.")
    """An emphasis effect."""

    MOTION_PATH = (4, "path", "A motion path.")
    """A motion path."""

    OLE_ACTION = (5, "verb", "An OLE action verb, e.g. to open an embedded object.")
    """An OLE action verb, e.g. to open an embedded object."""

    MEDIA = (6, "mediacall", "A media action: play, pause or stop.")
    """A media action: play, pause or stop."""


class MSO_ANIMATION_TRIGGER(BaseEnum):
    """Specifies what starts an animation effect.

    Example::

        from pptx.enum.animation import MSO_ANIMATION_TRIGGER

        clicks = [
            a for a in slide.animations if a.trigger == MSO_ANIMATION_TRIGGER.ON_PAGE_CLICK
        ]

    MS API name: `MsoAnimTriggerType`

    https://learn.microsoft.com/office/vba/api/powerpoint.msoanimtriggertype
    """

    NONE = (0, "No trigger.")
    """No trigger."""

    ON_PAGE_CLICK = (1, "A click on the slide, or the next-slide key.")
    """A click on the slide, or the next-slide key."""

    WITH_PREVIOUS = (2, "Starts together with the previous effect.")
    """Starts together with the previous effect."""

    AFTER_PREVIOUS = (3, "Starts when the previous effect ends.")
    """Starts when the previous effect ends."""

    ON_SHAPE_CLICK = (4, "A click on a particular shape, the effect's trigger shape.")
    """A click on a particular shape, the effect's trigger shape."""

    ON_MEDIA_BOOKMARK = (5, "Reaching a bookmark while media plays.")
    """Reaching a bookmark while media plays."""


class MSO_ANIMATION_EFFECT(BaseEnum):
    """Names a PowerPoint entrance or exit effect.

    ``Animation.effect_type`` uses this only for entrance and exit effects with a
    ``Animation.preset_id`` from 1 to 31. For those IDs PowerPoint's file format and its
    `MsoAnimEffect` enumeration use the same numbers (cross-checked against LibreOffice's
    importer); an exit effect is the reverse of the entrance effect of the same name. Other
    IDs, and emphasis and motion-path effects, are numbered differently in the file format,
    so they have no member here.

    Example::

        from pptx.enum.animation import MSO_ANIMATION_EFFECT

        fades = [a for a in slide.animations if a.effect_type == MSO_ANIMATION_EFFECT.FADE]

    MS API name: `MsoAnimEffect` (values 1 to 31)

    https://learn.microsoft.com/office/vba/api/powerpoint.msoanimeffect
    """

    APPEAR = (1, "Appear.")
    """Appear."""

    FLY = (2, "Fly in or out.")
    """Fly in or out."""

    BLINDS = (3, "Blinds.")
    """Blinds."""

    BOX = (4, "Box.")
    """Box."""

    CHECKERBOARD = (5, "Checkerboard.")
    """Checkerboard."""

    CIRCLE = (6, "Circle.")
    """Circle."""

    CRAWL = (7, "Crawl in or out.")
    """Crawl in or out."""

    DIAMOND = (8, "Diamond.")
    """Diamond."""

    DISSOLVE = (9, "Dissolve.")
    """Dissolve."""

    FADE = (10, "Fade.")
    """Fade."""

    FLASH_ONCE = (11, "Flash once.")
    """Flash once."""

    PEEK = (12, "Peek in or out.")
    """Peek in or out."""

    PLUS = (13, "Plus.")
    """Plus."""

    RANDOM_BARS = (14, "Random bars.")
    """Random bars."""

    SPIRAL = (15, "Spiral in or out.")
    """Spiral in or out."""

    SPLIT = (16, "Split.")
    """Split."""

    STRETCH = (17, "Stretch.")
    """Stretch."""

    STRIPS = (18, "Strips.")
    """Strips."""

    SWIVEL = (19, "Swivel.")
    """Swivel."""

    WEDGE = (20, "Wedge.")
    """Wedge."""

    WHEEL = (21, "Wheel.")
    """Wheel."""

    WIPE = (22, "Wipe.")
    """Wipe."""

    ZOOM = (23, "Zoom.")
    """Zoom."""

    RANDOM_EFFECTS = (24, "Random effects.")
    """Random effects."""

    BOOMERANG = (25, "Boomerang.")
    """Boomerang."""

    BOUNCE = (26, "Bounce.")
    """Bounce."""

    COLOR_REVEAL = (27, "Color reveal.")
    """Color reveal."""

    CREDITS = (28, "Credits.")
    """Credits."""

    EASE_IN = (29, "Ease in or out.")
    """Ease in or out."""

    FLOAT = (30, "Float.")
    """Float."""

    GROW_AND_TURN = (31, "Grow and turn.")
    """Grow and turn."""
