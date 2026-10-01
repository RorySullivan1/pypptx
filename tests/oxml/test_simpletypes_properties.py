"""Property-based tests (Hypothesis) for the converters in `pptx.oxml.simpletypes`.

Every concrete simple type is checked against three properties:

- `to_xml()` on any value either returns a string or raises |InvalidValueError| /
  |InvalidTypeError|, never anything else;
- whatever `to_xml()` writes, `from_xml()` reads back, and writing that again gives the same
  text (the XML form is stable);
- `from_xml()` on arbitrary text either returns a value or raises `ValueError` /
  |InvalidXmlError|.

Runs are derandomized so a failure reproduces, and use no example database.
"""

from __future__ import annotations

import inspect
import math

import pytest
from hypothesis import HealthCheck, given, settings
from hypothesis import strategies as st

import pptx.oxml.simpletypes as simpletypes
from pptx.exc import InvalidTypeError, InvalidValueError, InvalidXmlError

# -- abstract bases, used only through their subclasses --
_ABSTRACT = {"XsdStringEnumeration", "XsdTokenEnumeration"}

SIMPLE_TYPES = sorted(
    (
        cls
        for name, cls in vars(simpletypes).items()
        if inspect.isclass(cls)
        and issubclass(cls, simpletypes.BaseSimpleType)
        and cls.__module__ == simpletypes.__name__
        and not name.startswith("Base")
        and name not in _ABSTRACT
    ),
    key=lambda cls: cls.__name__,
)
# -- types that are only ever read (no `validate()`), e.g. ST_UniversalMeasure --
WRITABLE_TYPES = [cls for cls in SIMPLE_TYPES if hasattr(cls, "validate")]

ANY_VALUE = st.one_of(
    st.none(),
    st.booleans(),
    st.integers(min_value=-(10**12), max_value=10**12),
    st.floats(allow_nan=True, allow_infinity=True),
    st.text(max_size=12),
    st.sampled_from(["0", "1", "true", "false", "FF0000", "50%", "12.5%", "-1", "3.2", "auto"]),
)


# -- values most range-limited types accept, which the wide draws above rarely land on --
IN_RANGE = st.one_of(
    st.integers(min_value=0, max_value=100),
    st.integers(min_value=-600, max_value=600),
    st.integers(min_value=0, max_value=2_000_000),
    st.floats(min_value=0.0, max_value=1.0),
    st.integers(min_value=914400, max_value=51206400),  # -- slide sizes: 1 to 56 inches --
    st.text(alphabet="0123456789ABCDEFabcdef", min_size=6, max_size=6),
    st.sampled_from(["Internal", "External"]),
)


def values_for(cls) -> st.SearchStrategy:
    """Values to try on `cls`: the wide draws, in-range ones, and any enumeration members."""
    strategies = [ANY_VALUE, IN_RANGE]
    members = getattr(cls, "_members", None)
    if members:
        strategies.insert(0, st.sampled_from(sorted(members)))
    return st.one_of(*strategies)


PROPERTY_SETTINGS = settings(
    max_examples=200,
    derandomize=True,
    database=None,
    suppress_health_check=[HealthCheck.function_scoped_fixture],
)


class DescribeSimpleTypeProperties:
    """Properties that hold for every simple type."""

    @pytest.mark.parametrize("cls", WRITABLE_TYPES, ids=lambda cls: cls.__name__)
    def it_writes_a_valid_value_or_rejects_it_with_a_validation_error(self, cls):
        @PROPERTY_SETTINGS
        @given(values_for(cls))
        def check(value):
            try:
                xml_value = cls.to_xml(value)
            except (InvalidValueError, InvalidTypeError):
                return
            assert isinstance(xml_value, str)

        check()

    @pytest.mark.parametrize("cls", WRITABLE_TYPES, ids=lambda cls: cls.__name__)
    def it_reads_back_what_it_writes_and_writes_it_the_same_again(self, cls):
        @PROPERTY_SETTINGS
        @given(values_for(cls))
        def check(value):
            try:
                xml_value = cls.to_xml(value)
            except (InvalidValueError, InvalidTypeError):
                return

            assert cls.to_xml(cls.from_xml(xml_value)) == xml_value

        check()

    @pytest.mark.parametrize("cls", SIMPLE_TYPES, ids=lambda cls: cls.__name__)
    def it_reads_any_text_or_rejects_it_as_a_ValueError_or_InvalidXmlError(self, cls):
        @PROPERTY_SETTINGS
        @given(st.text(max_size=16))
        def check(text):
            try:
                cls.from_xml(text)
            except (ValueError, InvalidXmlError):
                pass

        check()


class DescribeSimpleTypeRegressions:
    """Defects the properties above found, pinned as plain examples."""

    @pytest.mark.parametrize(
        "cls", [simpletypes.XsdInt, simpletypes.ST_Coordinate, simpletypes.ST_TextIndentLevelType]
    )
    @pytest.mark.parametrize("value", [True, False])
    def it_rejects_a_bool_where_an_int_is_expected(self, cls, value):
        # -- bool is an Integral, and was written into the XML as "True"/"False" --
        with pytest.raises(InvalidTypeError):
            cls.to_xml(value)

    @pytest.mark.parametrize(
        "cls",
        [
            simpletypes.ST_Angle,
            simpletypes.ST_Percentage,
            simpletypes.ST_PositiveFixedAngle,
            simpletypes.ST_PositiveFixedPercentage,
        ],
    )
    @pytest.mark.parametrize("value", [math.nan, math.inf, -math.inf])
    def it_rejects_a_value_that_is_not_finite(self, cls, value):
        with pytest.raises(InvalidValueError, match="finite"):
            cls.to_xml(value)

    @pytest.mark.parametrize("value", [1.7976931348623157e308, -1.004e305])
    def it_writes_a_huge_angle_without_overflowing(self, value):
        xml_value = simpletypes.ST_Angle.to_xml(value)

        assert 0 <= int(xml_value) < simpletypes.ST_Angle.THREE_SIXTY

    @pytest.mark.parametrize("degrees", [-1.192092896e-07, 359.999999999, -360.0, 720.0])
    def it_keeps_a_positive_fixed_angle_below_360_degrees(self, degrees):
        # -- `a:lin@ang` must be less than 21600000; -1.19e-07 used to be written as 21600000 --
        xml_value = simpletypes.ST_PositiveFixedAngle.to_xml(degrees)

        assert 0 <= int(xml_value) < 21600000

    def it_writes_a_huge_int_angle_without_overflowing(self):
        # -- `math.isfinite()` overflows on an int this large; only floats are checked --
        assert simpletypes.ST_Angle.to_xml(10**400) == simpletypes.ST_Angle.to_xml(
            10**400 % 360
        )

    def it_rejects_a_universal_measure_with_an_unknown_unit(self):
        with pytest.raises(InvalidXmlError, match="universal measure must end in one of"):
            simpletypes.ST_UniversalMeasure.from_xml("5qq")
