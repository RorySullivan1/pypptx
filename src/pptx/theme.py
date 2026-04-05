"""Theme API objects for accessing presentation theme properties."""

from __future__ import annotations


class Theme:
    """Provides access to theme properties of a presentation.

    Accessed via the ``theme`` property on a slide master or presentation.
    """

    def __init__(self, theme_elm):
        self._theme = theme_elm

    @property
    def color_scheme(self):
        """A |ColorScheme| object providing access to theme colors.

        Returns None if no color scheme is defined.
        """
        themeElements = self._theme.themeElements
        if themeElements is None:
            return None
        clrScheme = themeElements.clrScheme
        if clrScheme is None:
            return None
        return ColorScheme(clrScheme)

    @property
    def font_scheme(self):
        """A |FontScheme| object providing access to theme fonts.

        Returns None if no font scheme is defined.
        """
        themeElements = self._theme.themeElements
        if themeElements is None:
            return None
        fontScheme = themeElements.fontScheme
        if fontScheme is None:
            return None
        return FontScheme(fontScheme)

    @property
    def name(self):
        """Read-only string name of this theme, from the ``name`` attribute."""
        return self._theme.get("name")


#: Maps color scheme slot names to their XML element tag suffixes.
_COLOR_SLOT_NAMES = (
    "dk1",
    "lt1",
    "dk2",
    "lt2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
)


class ColorScheme:
    """Provides access to the 12 theme color slots.

    Each color slot can be read as an RGB hex string (e.g. ``"4472C4"``).
    """

    def __init__(self, clrScheme):
        self._clrScheme = clrScheme

    @property
    def name(self):
        """Read/write string name of this color scheme."""
        return self._clrScheme.name

    @name.setter
    def name(self, value):
        self._clrScheme.name = value

    def __getitem__(self, slot_name):
        """Return the RGB hex string for the given color slot name.

        *slot_name* is one of ``"dk1"``, ``"lt1"``, ``"dk2"``, ``"lt2"``,
        ``"accent1"`` through ``"accent6"``, ``"hlink"``, ``"folHlink"``.

        Returns None if the slot is empty or has no resolvable color.
        """
        if slot_name not in _COLOR_SLOT_NAMES:
            raise KeyError(f"'{slot_name}' is not a valid color slot name")
        color_elm = getattr(self._clrScheme, slot_name)
        if color_elm is None:
            return None
        return _extract_rgb(color_elm)

    def __iter__(self):
        """Yield ``(slot_name, rgb_hex)`` tuples for all 12 color slots."""
        for name in _COLOR_SLOT_NAMES:
            yield name, self[name]

    def __len__(self):
        return len(_COLOR_SLOT_NAMES)

    @property
    def dark_1(self):
        """RGB hex string for the dark 1 color, or None."""
        return self["dk1"]

    @property
    def light_1(self):
        """RGB hex string for the light 1 color, or None."""
        return self["lt1"]

    @property
    def dark_2(self):
        """RGB hex string for the dark 2 color, or None."""
        return self["dk2"]

    @property
    def light_2(self):
        """RGB hex string for the light 2 color, or None."""
        return self["lt2"]

    @property
    def accent_1(self):
        """RGB hex string for accent 1, or None."""
        return self["accent1"]

    @property
    def accent_2(self):
        """RGB hex string for accent 2, or None."""
        return self["accent2"]

    @property
    def accent_3(self):
        """RGB hex string for accent 3, or None."""
        return self["accent3"]

    @property
    def accent_4(self):
        """RGB hex string for accent 4, or None."""
        return self["accent4"]

    @property
    def accent_5(self):
        """RGB hex string for accent 5, or None."""
        return self["accent5"]

    @property
    def accent_6(self):
        """RGB hex string for accent 6, or None."""
        return self["accent6"]

    @property
    def hyperlink(self):
        """RGB hex string for hyperlink color, or None."""
        return self["hlink"]

    @property
    def followed_hyperlink(self):
        """RGB hex string for followed hyperlink color, or None."""
        return self["folHlink"]


class FontScheme:
    """Provides access to theme font definitions (major and minor)."""

    def __init__(self, fontScheme):
        self._fontScheme = fontScheme

    @property
    def name(self):
        """Read/write string name of this font scheme."""
        return self._fontScheme.name

    @name.setter
    def name(self, value):
        self._fontScheme.name = value

    @property
    def major_font(self):
        """A |FontCollection| for headings/titles, or None."""
        majorFont = self._fontScheme.majorFont
        if majorFont is None:
            return None
        return FontCollection(majorFont)

    @property
    def minor_font(self):
        """A |FontCollection| for body text, or None."""
        minorFont = self._fontScheme.minorFont
        if minorFont is None:
            return None
        return FontCollection(minorFont)


class FontCollection:
    """Provides access to Latin, East Asian, and Complex Script typefaces."""

    def __init__(self, fontCollection):
        self._fc = fontCollection

    @property
    def latin(self):
        """Read/write string typeface name for Latin text, or None."""
        elm = self._fc.latin
        if elm is None:
            return None
        return elm.typeface

    @latin.setter
    def latin(self, value):
        elm = self._fc.get_or_add_latin()
        elm.typeface = value

    @property
    def east_asian(self):
        """Read/write string typeface name for East Asian text, or None."""
        elm = self._fc.ea
        if elm is None:
            return None
        return elm.typeface

    @east_asian.setter
    def east_asian(self, value):
        elm = self._fc.get_or_add_ea()
        elm.typeface = value

    @property
    def complex_script(self):
        """Read/write string typeface name for Complex Script text, or None."""
        elm = self._fc.cs
        if elm is None:
            return None
        return elm.typeface

    @complex_script.setter
    def complex_script(self, value):
        elm = self._fc.get_or_add_cs()
        elm.typeface = value


def _extract_rgb(color_elm):
    """Return RGB hex string from a theme color element.

    Theme color elements contain either ``a:srgbClr`` (with ``val`` attribute)
    or ``a:sysClr`` (with ``lastClr`` attribute). Returns None if no
    recognizable color child is found.
    """
    srgb = color_elm.find("{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr")
    if srgb is not None:
        return srgb.get("val")
    sys_clr = color_elm.find("{http://schemas.openxmlformats.org/drawingml/2006/main}sysClr")
    if sys_clr is not None:
        return sys_clr.get("lastClr")
    return None
