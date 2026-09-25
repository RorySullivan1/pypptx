"""Unit-test suite for `pptx.chart.datatable` module."""

from __future__ import annotations

import pytest

from pptx.chart.datatable import DataTable

from ..unitutil.cxml import element, xml


class DescribeDataTable:
    def it_knows_whether_it_has_a_horizontal_border(self, horizontal_border_get_fixture):
        data_table, expected_value = horizontal_border_get_fixture
        assert data_table.horizontal_border == expected_value

    def it_can_change_whether_it_has_a_horizontal_border(self, horizontal_border_set_fixture):
        data_table, new_value, expected_xml = horizontal_border_set_fixture
        data_table.horizontal_border = new_value
        assert data_table._element.xml == expected_xml

    def it_knows_whether_it_has_a_vertical_border(self, vertical_border_get_fixture):
        data_table, expected_value = vertical_border_get_fixture
        assert data_table.vertical_border == expected_value

    def it_can_change_whether_it_has_a_vertical_border(self, vertical_border_set_fixture):
        data_table, new_value, expected_xml = vertical_border_set_fixture
        data_table.vertical_border = new_value
        assert data_table._element.xml == expected_xml

    def it_knows_whether_it_has_an_outline(self, outline_get_fixture):
        data_table, expected_value = outline_get_fixture
        assert data_table.outline == expected_value

    def it_can_change_whether_it_has_an_outline(self, outline_set_fixture):
        data_table, new_value, expected_xml = outline_set_fixture
        data_table.outline = new_value
        assert data_table._element.xml == expected_xml

    def it_knows_whether_it_shows_legend_keys(self, show_keys_get_fixture):
        data_table, expected_value = show_keys_get_fixture
        assert data_table.show_keys == expected_value

    def it_can_change_whether_it_shows_legend_keys(self, show_keys_set_fixture):
        data_table, new_value, expected_xml = show_keys_set_fixture
        data_table.show_keys = new_value
        assert data_table._element.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("c:dTable", True),
            ("c:dTable/c:showHorzBorder{val=1}", True),
            ("c:dTable/c:showHorzBorder{val=0}", False),
        ]
    )
    def horizontal_border_get_fixture(self, request):
        dTable_cxml, expected_value = request.param
        data_table = DataTable(element(dTable_cxml))
        return data_table, expected_value

    @pytest.fixture(
        params=[
            # -- True is the schema default for c:showHorzBorder, so assigning it removes
            # -- any existing val attribute (or writes no val attribute) rather than writing
            # -- val="1" explicitly.
            ("c:dTable", True, "c:dTable/c:showHorzBorder"),
            ("c:dTable", False, "c:dTable/c:showHorzBorder{val=0}"),
            (
                "c:dTable/c:showHorzBorder{val=0}",
                True,
                "c:dTable/c:showHorzBorder",
            ),
        ]
    )
    def horizontal_border_set_fixture(self, request):
        dTable_cxml, new_value, expected_cxml = request.param
        data_table = DataTable(element(dTable_cxml))
        expected_xml = xml(expected_cxml)
        return data_table, new_value, expected_xml

    @pytest.fixture(
        params=[
            ("c:dTable", True),
            ("c:dTable/c:showVertBorder{val=1}", True),
            ("c:dTable/c:showVertBorder{val=0}", False),
        ]
    )
    def vertical_border_get_fixture(self, request):
        dTable_cxml, expected_value = request.param
        data_table = DataTable(element(dTable_cxml))
        return data_table, expected_value

    @pytest.fixture(
        params=[
            ("c:dTable", True, "c:dTable/c:showVertBorder"),
            ("c:dTable", False, "c:dTable/c:showVertBorder{val=0}"),
        ]
    )
    def vertical_border_set_fixture(self, request):
        dTable_cxml, new_value, expected_cxml = request.param
        data_table = DataTable(element(dTable_cxml))
        expected_xml = xml(expected_cxml)
        return data_table, new_value, expected_xml

    @pytest.fixture(
        params=[
            ("c:dTable", True),
            ("c:dTable/c:showOutline{val=1}", True),
            ("c:dTable/c:showOutline{val=0}", False),
        ]
    )
    def outline_get_fixture(self, request):
        dTable_cxml, expected_value = request.param
        data_table = DataTable(element(dTable_cxml))
        return data_table, expected_value

    @pytest.fixture(
        params=[
            ("c:dTable", True, "c:dTable/c:showOutline"),
            ("c:dTable", False, "c:dTable/c:showOutline{val=0}"),
        ]
    )
    def outline_set_fixture(self, request):
        dTable_cxml, new_value, expected_cxml = request.param
        data_table = DataTable(element(dTable_cxml))
        expected_xml = xml(expected_cxml)
        return data_table, new_value, expected_xml

    @pytest.fixture(
        params=[
            ("c:dTable", True),
            ("c:dTable/c:showKeys{val=1}", True),
            ("c:dTable/c:showKeys{val=0}", False),
        ]
    )
    def show_keys_get_fixture(self, request):
        dTable_cxml, expected_value = request.param
        data_table = DataTable(element(dTable_cxml))
        return data_table, expected_value

    @pytest.fixture(
        params=[
            ("c:dTable", True, "c:dTable/c:showKeys"),
            ("c:dTable", False, "c:dTable/c:showKeys{val=0}"),
            (
                "c:dTable/c:showKeys{val=0}",
                True,
                "c:dTable/c:showKeys",
            ),
        ]
    )
    def show_keys_set_fixture(self, request):
        dTable_cxml, new_value, expected_cxml = request.param
        data_table = DataTable(element(dTable_cxml))
        expected_xml = xml(expected_cxml)
        return data_table, new_value, expected_xml
