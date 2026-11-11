"""Tests for StyledTable.autosize_rows() including the cell text margins."""

from __future__ import annotations

from pptx import Presentation
from pptx.tablekit import CellStyle, StyledTable, min_row_height_for_font
from pptx.util import Cm, Inches, Pt

DEFAULT_MARGINS = Inches(0.1)  # 0.05" top + 0.05" bottom


def _styled(nrows=3, ncols=2):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    st = StyledTable.add(slide.shapes, nrows, ncols, 0, 0, Cm(10), Cm(4))
    return st


def _heights(st):
    return [row.height for row in st.table.rows]


class DescribeStyledTable_autosize_rows:
    def it_includes_the_default_cell_margins(self):
        st = _styled()
        st.autosize_rows(font_pt=10)
        assert _heights(st) == [Pt(12) + DEFAULT_MARGINS] * 3

    def it_uses_header_pt_for_the_first_row_only(self):
        st = _styled()
        st.autosize_rows(font_pt=10, header_pt=20)
        assert _heights(st) == [
            Pt(24) + DEFAULT_MARGINS,
            Pt(12) + DEFAULT_MARGINS,
            Pt(12) + DEFAULT_MARGINS,
        ]

    def it_multiplies_only_the_text_part_by_lines(self):
        st = _styled()
        st.autosize_rows(font_pt=10, lines=2)
        assert _heights(st) == [Pt(24) + DEFAULT_MARGINS] * 3

    def it_adds_pad_cm_on_top_of_text_and_margins(self):
        st = _styled()
        st.autosize_rows(font_pt=10, pad_cm=0.5)
        assert _heights(st) == [Pt(12) + DEFAULT_MARGINS + Cm(0.5)] * 3

    def it_reproduces_the_old_numbers_when_include_margins_is_False(self):
        st = _styled()
        st.autosize_rows(font_pt=10, lines=2, pad_cm=0.3, header_pt=14, include_margins=False)
        expected_header = min_row_height_for_font(14, lines=2, pad_cm=0.3)
        expected_body = min_row_height_for_font(10, lines=2, pad_cm=0.3)
        assert _heights(st) == [expected_header, expected_body, expected_body]

    def it_honours_custom_margins_set_on_a_cell(self):
        st = _styled()
        st.cell(1, 0).margin_top = Inches(0.2)
        st.cell(1, 0).margin_bottom = Inches(0.3)
        st.cell(1, 1).margin_top = Inches(0.2)
        st.cell(1, 1).margin_bottom = Inches(0.3)
        st.autosize_rows(font_pt=10)
        assert _heights(st) == [
            Pt(12) + DEFAULT_MARGINS,
            Pt(12) + Inches(0.5),
            Pt(12) + DEFAULT_MARGINS,
        ]

    def it_uses_the_largest_margin_sum_in_a_row(self):
        st = _styled()
        st.cell(0, 0).margin_top = Inches(0.01)
        st.cell(0, 0).margin_bottom = Inches(0.01)
        st.cell(0, 1).margin_top = Inches(0.4)
        st.cell(0, 1).margin_bottom = Inches(0.1)
        st.autosize_rows(font_pt=10)
        assert st.table.rows[0].height == Pt(12) + Inches(0.5)
        assert st.table.rows[1].height == Pt(12) + DEFAULT_MARGINS

    def it_adds_nothing_for_zero_margins_applied_via_CellStyle(self):
        st = _styled()
        zero = CellStyle(margins_cm=(None, None, 0.0, 0.0))
        for c in range(2):
            st.style_cell(1, c, zero)
        st.autosize_rows(font_pt=10)
        assert st.table.rows[1].height == Pt(12)
        assert st.table.rows[0].height == Pt(12) + DEFAULT_MARGINS

    def it_honours_non_zero_margins_from_CellStyle(self):
        st = _styled()
        style = CellStyle(margins_cm=(None, None, 0.1, 0.2))
        for c in range(2):
            st.style_cell(0, c, style)
        st.autosize_rows(font_pt=10)
        assert st.table.rows[0].height == Pt(12) + Cm(0.1) + Cm(0.2)

    def it_does_not_fail_on_a_vertical_merge(self):
        st = _styled(nrows=3, ncols=2)
        st.cell(0, 0).merge(st.cell(1, 0))
        st.autosize_rows(font_pt=10)
        assert _heights(st) == [Pt(12) + DEFAULT_MARGINS] * 3

    def it_does_not_fail_on_a_horizontal_merge(self):
        st = _styled(nrows=3, ncols=3)
        st.cell(1, 0).merge(st.cell(1, 2))
        st.autosize_rows(font_pt=10)
        assert _heights(st) == [Pt(12) + DEFAULT_MARGINS] * 3

    def it_returns_the_StyledTable_for_chaining(self):
        st = _styled()
        assert st.autosize_rows(font_pt=10) is st
