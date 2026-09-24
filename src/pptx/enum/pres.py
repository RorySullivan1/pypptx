"""Enumerations used by presentation-level objects."""

from __future__ import annotations

from pptx.enum.base import BaseEnum


class PP_SLIDE_SHOW_TYPE(BaseEnum):
    """Specifies how a presentation is displayed when it is shown as a slide show.

    Corresponds to the `p:showPr` choice of `p:present`, `p:browse`, or `p:kiosk` in
    `ppt/presProps.xml`. Alias: `PP_SLIDE_SHOW_TYPE`.

    Example::

        from pptx.enum.pres import PP_SLIDE_SHOW_TYPE

        settings = prs.slide_show_settings
        settings.show_type = PP_SLIDE_SHOW_TYPE.KIOSK
    """

    SPEAKER = (1, "The presentation is shown full screen, presented by a speaker (`p:present`).")
    """The presentation is shown full screen, presented by a speaker (`p:present`)."""

    BROWSE = (2, "The presentation is shown in a window, browsed by an individual (`p:browse`).")
    """The presentation is shown in a window, browsed by an individual (`p:browse`)."""

    KIOSK = (3, "The presentation is shown full screen, browsed at a kiosk (`p:kiosk`).")
    """The presentation is shown full screen, browsed at a kiosk (`p:kiosk`)."""
