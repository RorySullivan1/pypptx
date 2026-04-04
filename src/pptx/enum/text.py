"""Enumerations used by text and related objects."""

from __future__ import annotations

from pptx.enum.base import BaseEnum, BaseXmlEnum


class MSO_AUTO_SIZE(BaseEnum):
    """Determines the type of automatic sizing allowed.

    The following names can be used to specify the automatic sizing behavior used to fit a shape's
    text within the shape bounding box, for example::

        from pptx.enum.text import MSO_AUTO_SIZE

        shape.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

    The word-wrap setting of the text frame interacts with the auto-size setting to determine the
    specific auto-sizing behavior.

    Note that `TextFrame.auto_size` can also be set to |None|, which removes the auto size setting
    altogether. This causes the setting to be inherited, either from the layout placeholder, in the
    case of a placeholder shape, or from the theme.

    MS API Name: `MsoAutoSize`

    http://msdn.microsoft.com/en-us/library/office/ff865367(v=office.15).aspx
    """

    NONE = (
        0,
        "No automatic sizing of the shape or text will be done.\n\nText can freely extend beyond"
        " the horizontal and vertical edges of the shape bounding box.",
    )
    """No automatic sizing of the shape or text will be done.

    Text can freely extend beyond the horizontal and vertical edges of the shape bounding box.
    """

    SHAPE_TO_FIT_TEXT = (
        1,
        "The shape height and possibly width are adjusted to fit the text.\n\nNote this setting"
        " interacts with the TextFrame.word_wrap property setting. If word wrap is turned on,"
        " only the height of the shape will be adjusted; soft line breaks will be used to fit the"
        " text horizontally.",
    )
    """The shape height and possibly width are adjusted to fit the text.

    Note this setting interacts with the TextFrame.word_wrap property setting. If word wrap is
    turned on, only the height of the shape will be adjusted; soft line breaks will be used to fit
    the text horizontally.
    """

    TEXT_TO_FIT_SHAPE = (
        2,
        "The font size is reduced as necessary to fit the text within the shape.",
    )
    """The font size is reduced as necessary to fit the text within the shape."""

    MIXED = (-2, "Return value only; indicates a combination of automatic sizing schemes are used.")
    """Return value only; indicates a combination of automatic sizing schemes are used."""


class MSO_TEXT_UNDERLINE_TYPE(BaseXmlEnum):
    """
    Indicates the type of underline for text. Used with
    :attr:`.Font.underline` to specify the style of text underlining.

    Alias: ``MSO_UNDERLINE``

    Example::

        from pptx.enum.text import MSO_UNDERLINE

        run.font.underline = MSO_UNDERLINE.DOUBLE_LINE

    MS API Name: `MsoTextUnderlineType`

    http://msdn.microsoft.com/en-us/library/aa432699.aspx
    """

    NONE = (0, "none", "Specifies no underline.")
    """Specifies no underline."""

    DASH_HEAVY_LINE = (8, "dashHeavy", "Specifies a dash underline.")
    """Specifies a dash underline."""

    DASH_LINE = (7, "dash", "Specifies a dash line underline.")
    """Specifies a dash line underline."""

    DASH_LONG_HEAVY_LINE = (10, "dashLongHeavy", "Specifies a long heavy line underline.")
    """Specifies a long heavy line underline."""

    DASH_LONG_LINE = (9, "dashLong", "Specifies a dashed long line underline.")
    """Specifies a dashed long line underline."""

    DOT_DASH_HEAVY_LINE = (12, "dotDashHeavy", "Specifies a dot dash heavy line underline.")
    """Specifies a dot dash heavy line underline."""

    DOT_DASH_LINE = (11, "dotDash", "Specifies a dot dash line underline.")
    """Specifies a dot dash line underline."""

    DOT_DOT_DASH_HEAVY_LINE = (
        14,
        "dotDotDashHeavy",
        "Specifies a dot dot dash heavy line underline.",
    )
    """Specifies a dot dot dash heavy line underline."""

    DOT_DOT_DASH_LINE = (13, "dotDotDash", "Specifies a dot dot dash line underline.")
    """Specifies a dot dot dash line underline."""

    DOTTED_HEAVY_LINE = (6, "dottedHeavy", "Specifies a dotted heavy line underline.")
    """Specifies a dotted heavy line underline."""

    DOTTED_LINE = (5, "dotted", "Specifies a dotted line underline.")
    """Specifies a dotted line underline."""

    DOUBLE_LINE = (3, "dbl", "Specifies a double line underline.")
    """Specifies a double line underline."""

    HEAVY_LINE = (4, "heavy", "Specifies a heavy line underline.")
    """Specifies a heavy line underline."""

    SINGLE_LINE = (2, "sng", "Specifies a single line underline.")
    """Specifies a single line underline."""

    WAVY_DOUBLE_LINE = (17, "wavyDbl", "Specifies a wavy double line underline.")
    """Specifies a wavy double line underline."""

    WAVY_HEAVY_LINE = (16, "wavyHeavy", "Specifies a wavy heavy line underline.")
    """Specifies a wavy heavy line underline."""

    WAVY_LINE = (15, "wavy", "Specifies a wavy line underline.")
    """Specifies a wavy line underline."""

    WORDS = (1, "words", "Specifies underlining words.")
    """Specifies underlining words."""

    MIXED = (-2, "", "Specifies a mix of underline types (read-only).")
    """Specifies a mix of underline types (read-only)."""


MSO_UNDERLINE = MSO_TEXT_UNDERLINE_TYPE


class MSO_VERTICAL_ANCHOR(BaseXmlEnum):
    """Specifies the vertical alignment of text in a text frame.

    Used with the `.vertical_anchor` property of the |TextFrame| object. Note that the
    `vertical_anchor` property can also have the value None, indicating there is no directly
    specified vertical anchor setting and its effective value is inherited from its placeholder if
    it has one or from the theme. |None| may also be assigned to remove an explicitly specified
    vertical anchor setting.

    MS API Name: `MsoVerticalAnchor`

    http://msdn.microsoft.com/en-us/library/office/ff865255.aspx
    """

    TOP = (1, "t", "Aligns text to top of text frame")
    """Aligns text to top of text frame"""

    MIDDLE = (3, "ctr", "Centers text vertically")
    """Centers text vertically"""

    BOTTOM = (4, "b", "Aligns text to bottom of text frame")
    """Aligns text to bottom of text frame"""

    MIXED = (-2, "", "Return value only; indicates a combination of the other states.")
    """Return value only; indicates a combination of the other states."""


MSO_ANCHOR = MSO_VERTICAL_ANCHOR


class PP_PARAGRAPH_ALIGNMENT(BaseXmlEnum):
    """Specifies the horizontal alignment for one or more paragraphs.

    Alias: `PP_ALIGN`

    Example::

        from pptx.enum.text import PP_ALIGN

        shape.paragraphs[0].alignment = PP_ALIGN.CENTER

    MS API Name: `PpParagraphAlignment`

    http://msdn.microsoft.com/en-us/library/office/ff745375(v=office.15).aspx
    """

    CENTER = (2, "ctr", "Center align")
    """Center align"""

    DISTRIBUTE = (
        5,
        "dist",
        "Evenly distributes e.g. Japanese characters from left to right within a line",
    )
    """Evenly distributes e.g. Japanese characters from left to right within a line"""

    JUSTIFY = (
        4,
        "just",
        "Justified, i.e. each line both begins and ends at the margin.\n\nSpacing between words"
        " is adjusted such that the line exactly fills the width of the paragraph.",
    )
    """Justified, i.e. each line both begins and ends at the margin.

    Spacing between words is adjusted such that the line exactly fills the width of the paragraph.
    """

    JUSTIFY_LOW = (7, "justLow", "Justify using a small amount of space between words.")
    """Justify using a small amount of space between words."""

    LEFT = (1, "l", "Left aligned")
    """Left aligned"""

    RIGHT = (3, "r", "Right aligned")
    """Right aligned"""

    THAI_DISTRIBUTE = (6, "thaiDist", "Thai distributed")
    """Thai distributed"""

    MIXED = (-2, "", "Multiple alignments are present in a set of paragraphs (read-only).")
    """Multiple alignments are present in a set of paragraphs (read-only)."""


PP_ALIGN = PP_PARAGRAPH_ALIGNMENT


class MSO_TEXT_STRIKE_TYPE(BaseXmlEnum):
    """Specifies the type of strikethrough for text.

    Used with :attr:`.Font.strikethrough` to specify the style of text strikethrough.

    Example::

        from pptx.enum.text import MSO_TEXT_STRIKE_TYPE

        run.font.strikethrough = MSO_TEXT_STRIKE_TYPE.SINGLE_STRIKE

    MS API Name: `MsoTextStrikeType`
    """

    NO_STRIKE = (0, "noStrike", "Specifies no strike.")
    """Specifies no strike."""

    SINGLE_STRIKE = (1, "sngStrike", "Specifies a single strike.")
    """Specifies a single strike."""

    DOUBLE_STRIKE = (2, "dblStrike", "Specifies a double strike.")
    """Specifies a double strike."""


class MSO_TEXT_CAPS(BaseXmlEnum):
    """Specifies the capitalization for text.

    Used with :attr:`.Font.caps` to specify the capitalization style.

    Example::

        from pptx.enum.text import MSO_TEXT_CAPS

        run.font.caps = MSO_TEXT_CAPS.ALL

    MS API Name: `MsoTextCaps`
    """

    NONE = (0, "none", "Specifies no capitalization.")
    """Specifies no capitalization."""

    ALL = (1, "all", "Specifies all capitals.")
    """Specifies all capitals."""

    SMALL = (2, "small", "Specifies small capitals.")
    """Specifies small capitals."""


class MSO_TEXT_FONT_ALIGN(BaseXmlEnum):
    """Specifies the vertical alignment of text relative to the text body baseline.

    Used with the `fontAlgn` attribute on paragraph properties.

    MS API Name: `MsoTextFontAlign`
    """

    AUTO = (0, "auto", "Automatic alignment.")
    """Automatic alignment."""

    TOP = (1, "t", "Text is aligned to the top of the text body.")
    """Text is aligned to the top of the text body."""

    CENTER = (2, "ctr", "Text is vertically centered.")
    """Text is vertically centered."""

    BASELINE = (3, "base", "Text is aligned to the baseline.")
    """Text is aligned to the baseline."""

    BOTTOM = (4, "b", "Text is aligned to the bottom of the text body.")
    """Text is aligned to the bottom of the text body."""


class MSO_TEXT_VERTICAL_TYPE(BaseXmlEnum):
    """Specifies the text direction within a text body.

    Used with the `vert` attribute on `a:bodyPr` to control text orientation.

    MS API Name: `MsoTextOrientation`
    """

    HORIZONTAL = (1, "horz", "Text flows horizontally (default).")
    """Text flows horizontally (default)."""

    VERTICAL = (2, "vert", "Text flows top to bottom; each line is rotated 90 degrees.")
    """Text flows top to bottom; each line is rotated 90 degrees."""

    VERTICAL_270 = (3, "vert270", "Text flows bottom to top; each line is rotated 270 degrees.")
    """Text flows bottom to top; each line is rotated 270 degrees."""

    WORD_ART_VERTICAL = (
        4,
        "wordArtVert",
        "Text flows top to bottom; characters are not rotated, stacked vertically.",
    )
    """Text flows top to bottom; characters are not rotated, stacked vertically."""

    EAST_ASIAN_VERTICAL = (
        5,
        "eaVert",
        "East Asian vertical text flow (top to bottom, right to left).",
    )
    """East Asian vertical text flow (top to bottom, right to left)."""

    MONGOLIAN_VERTICAL = (6, "mongolianVert", "Mongolian vertical text flow.")
    """Mongolian vertical text flow."""

    WORD_ART_VERTICAL_RTL = (
        7,
        "wordArtVertRtl",
        "Right-to-left WordArt vertical text.",
    )
    """Right-to-left WordArt vertical text."""


class MSO_PRESET_TEXT_SHAPE(BaseXmlEnum):
    """Specifies the preset text warp shape.

    Used with the `prst` attribute on `a:prstTxWarp` to control text warp effects.
    """

    TEXT_NO_SHAPE = (0, "textNoShape", "No text shape (default).")
    """No text shape (default)."""

    TEXT_PLAIN = (1, "textPlain", "Plain text.")
    """Plain text."""

    TEXT_STOP = (2, "textStop", "Octagon (stop sign) text.")
    """Octagon (stop sign) text."""

    TEXT_TRIANGLE = (3, "textTriangle", "Triangle text.")
    """Triangle text."""

    TEXT_TRIANGLE_INVERTED = (4, "textTriangleInverted", "Inverted triangle text.")
    """Inverted triangle text."""

    TEXT_CHEVRON = (5, "textChevron", "Chevron text.")
    """Chevron text."""

    TEXT_CHEVRON_INVERTED = (6, "textChevronInverted", "Inverted chevron text.")
    """Inverted chevron text."""

    TEXT_RING_INSIDE = (7, "textRingInside", "Ring inside text.")
    """Ring inside text."""

    TEXT_RING_OUTSIDE = (8, "textRingOutside", "Ring outside text.")
    """Ring outside text."""

    TEXT_ARCH_UP = (9, "textArchUp", "Arch up text.")
    """Arch up text."""

    TEXT_ARCH_DOWN = (10, "textArchDown", "Arch down text.")
    """Arch down text."""

    TEXT_CIRCLE = (11, "textCircle", "Circle text.")
    """Circle text."""

    TEXT_BUTTON = (12, "textButton", "Button text.")
    """Button text."""

    TEXT_ARCH_UP_POUR = (13, "textArchUpPour", "Arch up pour text.")
    """Arch up pour text."""

    TEXT_ARCH_DOWN_POUR = (14, "textArchDownPour", "Arch down pour text.")
    """Arch down pour text."""

    TEXT_CIRCLE_POUR = (15, "textCirclePour", "Circle pour text.")
    """Circle pour text."""

    TEXT_BUTTON_POUR = (16, "textButtonPour", "Button pour text.")
    """Button pour text."""

    TEXT_CURVE_UP = (17, "textCurveUp", "Curve up text.")
    """Curve up text."""

    TEXT_CURVE_DOWN = (18, "textCurveDown", "Curve down text.")
    """Curve down text."""

    TEXT_CAN_UP = (19, "textCanUp", "Can up text.")
    """Can up text."""

    TEXT_CAN_DOWN = (20, "textCanDown", "Can down text.")
    """Can down text."""

    TEXT_WAVE1 = (21, "textWave1", "Wave 1 text.")
    """Wave 1 text."""

    TEXT_WAVE2 = (22, "textWave2", "Wave 2 text.")
    """Wave 2 text."""

    TEXT_DOUBLE_WAVE1 = (23, "textDoubleWave1", "Double wave 1 text.")
    """Double wave 1 text."""

    TEXT_WAVE4 = (24, "textWave4", "Wave 4 text.")
    """Wave 4 text."""

    TEXT_INFLATE = (25, "textInflate", "Inflate text.")
    """Inflate text."""

    TEXT_DEFLATE = (26, "textDeflate", "Deflate text.")
    """Deflate text."""

    TEXT_INFLATE_BOTTOM = (27, "textInflateBottom", "Inflate bottom text.")
    """Inflate bottom text."""

    TEXT_DEFLATE_BOTTOM = (28, "textDeflateBottom", "Deflate bottom text.")
    """Deflate bottom text."""

    TEXT_INFLATE_TOP = (29, "textInflateTop", "Inflate top text.")
    """Inflate top text."""

    TEXT_DEFLATE_TOP = (30, "textDeflateTop", "Deflate top text.")
    """Deflate top text."""

    TEXT_DEFLATE_INFLATE = (31, "textDeflateInflate", "Deflate-inflate text.")
    """Deflate-inflate text."""

    TEXT_DEFLATE_INFLATE_DEFLATE = (
        32,
        "textDeflateInflateDeflate",
        "Deflate-inflate-deflate text.",
    )
    """Deflate-inflate-deflate text."""

    TEXT_FADE_RIGHT = (33, "textFadeRight", "Fade right text.")
    """Fade right text."""

    TEXT_FADE_LEFT = (34, "textFadeLeft", "Fade left text.")
    """Fade left text."""

    TEXT_FADE_UP = (35, "textFadeUp", "Fade up text.")
    """Fade up text."""

    TEXT_FADE_DOWN = (36, "textFadeDown", "Fade down text.")
    """Fade down text."""

    TEXT_SLANT_UP = (37, "textSlantUp", "Slant up text.")
    """Slant up text."""

    TEXT_SLANT_DOWN = (38, "textSlantDown", "Slant down text.")
    """Slant down text."""

    TEXT_CASCADE_UP = (39, "textCascadeUp", "Cascade up text.")
    """Cascade up text."""

    TEXT_CASCADE_DOWN = (40, "textCascadeDown", "Cascade down text.")
    """Cascade down text."""
