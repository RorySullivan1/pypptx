"""Unit-test suite for pptx.oxml.table module"""

from __future__ import annotations

import pytest

from pptx.oxml.ns import nsdecls
from pptx.oxml.table import CT_Table, TcRange

from ..unitutil.cxml import element


class DescribeCT_Table:
    def it_can_create_a_new_tbl_element_tree(self):
        """
        Indirectly tests that column widths are a proportional split of total
        width and that row heights a proportional split of total height.
        """
        expected_xml = (
            '<a:tbl %s>\n  <a:tblPr firstRow="1" bandRow="1">\n    <a:tableSt'
            "yleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>\n "
            ' </a:tblPr>\n  <a:tblGrid>\n    <a:gridCol w="111"/>\n    <a:gri'
            'dCol w="111"/>\n    <a:gridCol w="112"/>\n  </a:tblGrid>\n  <a:t'
            'r h="222">\n    <a:tc>\n      <a:txBody>\n        <a:bodyPr/>\n '
            "       <a:lstStyle/>\n        <a:p/>\n      </a:txBody>\n      <"
            "a:tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n        <a:"
            "bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      </a:txBod"
            "y>\n      <a:tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n"
            "        <a:bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n    "
            "  </a:txBody>\n      <a:tcPr/>\n    </a:tc>\n  </a:tr>\n  <a:tr "
            'h="223">\n    <a:tc>\n      <a:txBody>\n        <a:bodyPr/>\n   '
            "     <a:lstStyle/>\n        <a:p/>\n      </a:txBody>\n      <a:"
            "tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n        <a:bo"
            "dyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      </a:txBody>"
            "\n      <a:tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n  "
            "      <a:bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      "
            "</a:txBody>\n      <a:tcPr/>\n    </a:tc>\n  </a:tr>\n</a:tbl>\n" % nsdecls("a")
        )
        tbl = CT_Table.new_tbl(2, 3, 334, 445)
        assert tbl.xml == expected_xml

    def it_provides_access_to_its_tc_elements(self):
        tbl_cxml = "a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))"
        tbl = element(tbl_cxml)
        tcs = tbl.xpath(".//a:tc")

        assert tbl.tc(0, 0) is tcs[0]
        assert tbl.tc(0, 1) is tcs[1]
        assert tbl.tc(1, 0) is tcs[2]
        assert tbl.tc(1, 1) is tcs[3]

    def it_can_insert_a_row(self):
        tbl = CT_Table.new_tbl(2, 3, 900, 900)

        new_tr = tbl.insert_tr(1)

        assert len(tbl.tr_lst) == 3
        assert tbl.tr_lst[1] is new_tr
        # -- every row has one tc per grid column --
        assert all(len(tr.tc_lst) == 3 for tr in tbl.tr_lst)

    def it_can_append_a_row_when_idx_equals_the_row_count(self):
        tbl = CT_Table.new_tbl(2, 2, 900, 900)

        new_tr = tbl.insert_tr(2)

        assert len(tbl.tr_lst) == 3
        assert tbl.tr_lst[2] is new_tr

    def it_extends_a_vertical_merge_when_a_row_is_inserted_inside_it(self):
        tbl_cxml = (
            "a:tbl/(a:tblGrid/(a:gridCol{w=100},a:gridCol{w=100}),"
            'a:tr{h=10}/(a:tc{rowSpan=2}/a:txBody/a:p/a:r/a:t"origin",a:tc),'
            "a:tr{h=10}/(a:tc{vMerge=1},a:tc),"
            "a:tr{h=10}/(a:tc,a:tc))"
        )
        tbl = element(tbl_cxml)

        new_tr = tbl.insert_tr(1)

        # -- origin's span grows to include the new row --
        assert tbl.tc(0, 0).rowSpan == 3
        assert tbl.tc(0, 0).vMerge is False
        # -- the new row's cell becomes a merge continuation --
        assert new_tr.tc_lst[0].vMerge is True
        assert new_tr.tc_lst[0].hMerge is False
        # -- unrelated column is unaffected --
        assert new_tr.tc_lst[1].vMerge is False
        assert new_tr.tc_lst[1].rowSpan == 1
        # -- row/cell counts stay consistent --
        assert len(tbl.tr_lst) == 4
        assert all(len(tr.tc_lst) == 2 for tr in tbl.tr_lst)

    def it_does_not_extend_a_merge_when_inserting_outside_its_range(self):
        tbl_cxml = (
            "a:tbl/(a:tblGrid/(a:gridCol{w=100},a:gridCol{w=100}),"
            "a:tr{h=10}/(a:tc{rowSpan=2},a:tc),"
            "a:tr{h=10}/(a:tc{vMerge=1},a:tc),"
            "a:tr{h=10}/(a:tc,a:tc))"
        )
        tbl = element(tbl_cxml)

        new_tr = tbl.insert_tr(0)

        assert tbl.tc(1, 0).rowSpan == 2  # -- unchanged --
        assert new_tr.tc_lst[0].vMerge is False
        assert new_tr.tc_lst[0].rowSpan == 1

    def it_can_remove_a_row(self):
        tbl = CT_Table.new_tbl(3, 2, 900, 900)

        tbl.remove_tr(1)

        assert len(tbl.tr_lst) == 2
        assert all(len(tr.tc_lst) == 2 for tr in tbl.tr_lst)

    def it_moves_the_merge_origin_down_when_its_row_is_removed(self):
        tbl_cxml = (
            "a:tbl/(a:tblGrid/a:gridCol{w=100},"
            'a:tr{h=10}/a:tc{rowSpan=3}/a:txBody/a:p/a:r/a:t"origin text",'
            "a:tr{h=10}/a:tc{vMerge=1},"
            "a:tr{h=10}/a:tc{vMerge=1})"
        )
        tbl = element(tbl_cxml)

        tbl.remove_tr(0)

        assert len(tbl.tr_lst) == 2
        new_origin = tbl.tc(0, 0)
        assert new_origin.rowSpan == 2
        assert new_origin.vMerge is False
        assert new_origin.text == "origin text"

    def it_unmerges_when_removing_a_row_leaves_a_span_of_one(self):
        tbl_cxml = (
            "a:tbl/(a:tblGrid/a:gridCol{w=100},"
            "a:tr{h=10}/a:tc{rowSpan=2},"
            "a:tr{h=10}/a:tc{vMerge=1})"
        )
        tbl = element(tbl_cxml)

        tbl.remove_tr(0)

        assert len(tbl.tr_lst) == 1
        cell = tbl.tc(0, 0)
        assert cell.rowSpan == 1
        assert cell.vMerge is False

    def it_shrinks_the_span_when_a_mid_row_of_a_merge_is_removed(self):
        tbl_cxml = (
            "a:tbl/(a:tblGrid/a:gridCol{w=100},"
            "a:tr{h=10}/a:tc{rowSpan=3},"
            "a:tr{h=10}/a:tc{vMerge=1},"
            "a:tr{h=10}/a:tc{vMerge=1})"
        )
        tbl = element(tbl_cxml)

        tbl.remove_tr(1)

        assert len(tbl.tr_lst) == 2
        origin = tbl.tc(0, 0)
        assert origin.rowSpan == 2
        assert origin.vMerge is False
        assert tbl.tc(1, 0).vMerge is True

    def it_can_insert_a_column(self):
        tbl = CT_Table.new_tbl(2, 2, 900, 900)

        new_gridCol = tbl.insert_gridCol(1)

        assert len(tbl.tblGrid.gridCol_lst) == 3
        assert tbl.tblGrid.gridCol_lst[1] is new_gridCol
        assert all(len(tr.tc_lst) == 3 for tr in tbl.tr_lst)

    def it_extends_a_horizontal_merge_when_a_column_is_inserted_inside_it(self):
        tbl_cxml = (
            "a:tbl/(a:tblGrid/(a:gridCol{w=100},a:gridCol{w=100}),"
            'a:tr{h=10}/(a:tc{gridSpan=2}/a:txBody/a:p/a:r/a:t"origin",a:tc{hMerge=1}),'
            "a:tr{h=10}/(a:tc,a:tc))"
        )
        tbl = element(tbl_cxml)

        tbl.insert_gridCol(1)

        assert tbl.tc(0, 0).gridSpan == 3
        assert tbl.tc(0, 1).hMerge is True
        assert tbl.tc(0, 1).vMerge is False
        # -- the second (unmerged) row is unaffected by the merge extension --
        assert tbl.tc(1, 1).hMerge is False
        assert len(tbl.tblGrid.gridCol_lst) == 3
        assert all(len(tr.tc_lst) == 3 for tr in tbl.tr_lst)

    def it_can_remove_a_column(self):
        tbl = CT_Table.new_tbl(2, 3, 900, 900)

        tbl.remove_gridCol(1)

        assert len(tbl.tblGrid.gridCol_lst) == 2
        assert all(len(tr.tc_lst) == 2 for tr in tbl.tr_lst)

    def it_moves_the_merge_origin_right_when_its_column_is_removed(self):
        tbl_cxml = (
            "a:tbl/(a:tblGrid/(a:gridCol{w=100},a:gridCol{w=100},a:gridCol{w=100}),"
            'a:tr{h=10}/(a:tc{gridSpan=3}/a:txBody/a:p/a:r/a:t"origin text",'
            "a:tc{hMerge=1},a:tc{hMerge=1}))"
        )
        tbl = element(tbl_cxml)

        tbl.remove_gridCol(0)

        assert len(tbl.tblGrid.gridCol_lst) == 2
        new_origin = tbl.tc(0, 0)
        assert new_origin.gridSpan == 2
        assert new_origin.hMerge is False
        assert new_origin.text == "origin text"


class DescribeTcRange:
    def it_knows_when_the_range_contains_a_merged_cell(self, contains_merge_fixture):
        tc, other_tc, expected_value = contains_merge_fixture
        tc_range = TcRange(tc, other_tc)

        contains_merged_cell = tc_range.contains_merged_cell

        assert contains_merged_cell is expected_value

    def it_knows_how_big_the_merge_range_is(self, dimensions_fixture):
        tc, other_tc, expected_value = dimensions_fixture
        tc_range = TcRange(tc, other_tc)

        dimensions = tc_range.dimensions

        assert dimensions == expected_value

    def it_knows_when_tcs_are_in_the_same_tbl(self, in_same_table_fixture):
        tc, other_tc, expected_value = in_same_table_fixture
        tc_range = TcRange(tc, other_tc)

        in_same_table = tc_range.in_same_table

        assert in_same_table is expected_value

    def it_can_iterate_tcs_not_in_left_col_of_range(self, except_left_fixture):
        tc, other_tc, expected_value = except_left_fixture
        tc_range = TcRange(tc, other_tc)

        tcs = list(tc_range.iter_except_left_col_tcs())

        assert tcs == expected_value

    def it_can_iterate_tcs_not_in_top_row_of_range(self, except_top_fixture):
        tc, other_tc, expected_value = except_top_fixture
        tc_range = TcRange(tc, other_tc)

        tcs = list(tc_range.iter_except_top_row_tcs())

        assert tcs == expected_value

    def it_can_iterate_left_col_of_range_tcs(self, left_col_fixture):
        tc, other_tc, expected_value = left_col_fixture
        tc_range = TcRange(tc, other_tc)

        tcs = list(tc_range.iter_left_col_tcs())

        assert tcs == expected_value

    def it_can_iterate_top_row_of_range_tcs(self, top_row_fixture):
        tc, other_tc, expected_value = top_row_fixture
        tc_range = TcRange(tc, other_tc)

        tcs = list(tc_range.iter_top_row_tcs())

        assert tcs == expected_value

    def it_can_migrate_range_content_to_origin_cell(self, move_fixture):
        tc, other_tc, expected_text = move_fixture
        tc_range = TcRange(tc, other_tc)

        tc_range.move_content_to_origin()

        assert tc.text == expected_text
        assert other_tc.text == ""

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            ("a:tbl/a:tr/(a:tc,a:tc)", False),
            ("a:tbl/a:tr/(a:tc{gridSpan=1},a:tc{hMerge=false})", False),
            ("a:tbl/a:tr/(a:tc{gridSpan=2},a:tc{hMerge=1})", True),
            ("a:tbl/(a:tr/a:tc,a:tr/a:tc)", False),
            ("a:tbl/(a:tr/a:tc{rowSpan=1},a:tr/a:tc{vMerge=false})", False),
            ("a:tbl/(a:tr/a:tc{rowSpan=2},a:tr/a:tc{vMerge=true})", True),
        ]
    )
    def contains_merge_fixture(self, request):
        tbl_cxml, expected_value = request.param
        tcs = element(tbl_cxml).xpath("//a:tc")
        return tcs[0], tcs[1], expected_value

    @pytest.fixture(
        params=[
            ("a:tbl/a:tr/(a:tc,a:tc)", (1, 2)),
            ("a:tbl/(a:tr/a:tc,a:tr/a:tc)", (2, 1)),
            ("a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))", (2, 2)),
        ]
    )
    def dimensions_fixture(self, request):
        tbl_cxml, expected_value = request.param
        tcs = element(tbl_cxml).xpath("//a:tc")
        return tcs[0], tcs[-1], expected_value

    @pytest.fixture(
        params=[
            ("a:tbl/a:tr/(a:tc,a:tc)", [0, 1], [1]),
            ("a:tbl/(a:tr/a:tc,a:tr/a:tc)", [0, 1], []),
            ("a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))", [2, 1], [1, 3]),
            (
                "a:tbl/(a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc" ",a:tc))",
                [0, 8],
                [1, 2, 4, 5, 7, 8],
            ),
        ]
    )
    def except_left_fixture(self, request):
        tbl_cxml, tc_idxs, expected_tc_idxs = request.param
        tcs = element(tbl_cxml).xpath("//a:tc")
        tc, other_tc = tcs[tc_idxs[0]], tcs[tc_idxs[1]]
        expected_value = [tcs[idx] for idx in expected_tc_idxs]
        return tc, other_tc, expected_value

    @pytest.fixture(
        params=[
            ("a:tbl/a:tr/(a:tc,a:tc)", [0, 1], []),
            ("a:tbl/(a:tr/a:tc,a:tr/a:tc)", [0, 1], [1]),
            ("a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))", [2, 1], [2, 3]),
            (
                "a:tbl/(a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc" ",a:tc))",
                [0, 8],
                [3, 4, 5, 6, 7, 8],
            ),
        ]
    )
    def except_top_fixture(self, request):
        tbl_cxml, tc_idxs, expected_tc_idxs = request.param
        tcs = element(tbl_cxml).xpath("//a:tc")
        tc, other_tc = tcs[tc_idxs[0]], tcs[tc_idxs[1]]
        expected_value = [tcs[idx] for idx in expected_tc_idxs]
        return tc, other_tc, expected_value

    @pytest.fixture(params=[True, False])
    def in_same_table_fixture(self, request):
        expected_value = request.param
        tbl = element("a:tbl/a:tr/(a:tc,a:tc)")
        other_tbl = element("a:tbl/a:tr/(a:tc,a:tc)")
        tc = tbl.xpath("//a:tc")[0]
        other_tc = tbl.xpath("//a:tc")[1] if expected_value else other_tbl.xpath("//a:tc")[1]
        return tc, other_tc, expected_value

    @pytest.fixture(
        params=[
            ("a:tbl/a:tr/(a:tc,a:tc)", (0, 1), (0,)),
            ("a:tbl/(a:tr/a:tc,a:tr/a:tc)", (0, 1), (0, 1)),
            ("a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))", (2, 1), (0, 2)),
            (
                "a:tbl/(a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc" ",a:tc))",
                (4, 8),
                (4, 7),
            ),
        ]
    )
    def left_col_fixture(self, request):
        tbl_cxml, tc_idxs, expected_tc_idxs = request.param
        tcs = element(tbl_cxml).xpath("//a:tc")
        tc, other_tc = tcs[tc_idxs[0]], tcs[tc_idxs[1]]
        expected_value = [tcs[idx] for idx in expected_tc_idxs]
        return tc, other_tc, expected_value

    @pytest.fixture(
        params=[
            ("a:tbl/a:tr/(a:tc/a:txBody/a:p,a:tc/a:txBody/a:p)", ""),
            ('a:tbl/a:tr/(a:tc/a:txBody/a:p,a:tc/a:txBody/a:p/a:r/a:t"b")', "b"),
            ('a:tbl/a:tr/(a:tc/a:txBody/a:p/a:r/a:t"a",a:tc/a:txBody/a:p)', "a"),
            (
                'a:tbl/a:tr/(a:tc/a:txBody/a:p/a:r/a:t"a",a:tc/a:txBody/a:p/a:r/a:t' '"b")',
                "a\nb",
            ),
            (
                'a:tbl/a:tr/(a:tc/a:txBody/a:p/a:r/a:t"a",a:tc/a:txBody/(a:p,a:p))',
                "a\n\n",
            ),
            (
                'a:tbl/a:tr/(a:tc/a:txBody/(a:p,a:p),a:tc/a:txBody/a:p/a:r/a:t"b")',
                "\n\nb",
            ),
        ]
    )
    def move_fixture(self, request):
        tbl_cxml, expected_text = request.param
        tcs = element(tbl_cxml).xpath("//a:tc")
        return tcs[0], tcs[1], expected_text

    @pytest.fixture(
        params=[
            ("a:tbl/a:tr/(a:tc,a:tc)", (0, 1), (0, 1)),
            ("a:tbl/(a:tr/a:tc,a:tr/a:tc)", (0, 1), (0,)),
            ("a:tbl/(a:tr/(a:tc,a:tc),a:tr/(a:tc,a:tc))", (2, 1), (0, 1)),
            (
                "a:tbl/(a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc,a:tc),a:tr/(a:tc,a:tc" ",a:tc))",
                (4, 8),
                (4, 5),
            ),
        ]
    )
    def top_row_fixture(self, request):
        tbl_cxml, tc_idxs, expected_tc_idxs = request.param
        tcs = element(tbl_cxml).xpath("//a:tc")
        tc, other_tc = tcs[tc_idxs[0]], tcs[tc_idxs[1]]
        expected_value = [tcs[idx] for idx in expected_tc_idxs]
        return tc, other_tc, expected_value
