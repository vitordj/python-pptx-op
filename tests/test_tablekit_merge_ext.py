"""Tests for StyledTable.merge_equal_runs() (no duplicated text in merge origin)."""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.tablekit import StyledTable
from pptx.util import Cm, Pt

SKIP = ("", "X")


def _styled(rows, cols=2):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    st = StyledTable.add(slide.shapes, len(rows), cols, 0, 0, Cm(10), Cm(4))
    st.fill_rows(rows)
    return prs, st


def _col_texts(st, col=0):
    return [st.cell(r, col).text for r in range(len(st.table.rows))]


class DescribeStyledTable_merge_equal_runs:
    def it_does_not_duplicate_text_in_the_merge_origin(self):
        _, st = _styled([["h1", "h2"], ["A", "1"], ["A", "2"], ["B", "3"]])
        st.merge_equal_runs(0, skip=SKIP)
        origin = st.cell(1, 0)
        assert origin.text == "A"
        assert len(origin.text_frame.paragraphs) == 1
        assert origin.is_merge_origin
        assert origin.span_height == 2
        assert st.cell(2, 0).is_spanned
        assert st.cell(3, 0).text == "B"

    def it_merges_a_run_of_three_equal_cells(self):
        _, st = _styled([["h", "x"], ["A", "1"], ["A", "2"], ["A", "3"], ["B", "4"]])
        st.merge_equal_runs(0, skip=SKIP)
        origin = st.cell(1, 0)
        assert origin.text == "A"
        assert len(origin.text_frame.paragraphs) == 1
        assert origin.span_height == 3
        assert st.cell(2, 0).is_spanned
        assert st.cell(3, 0).is_spanned
        assert not st.cell(4, 0).is_spanned

    def it_merges_two_separate_runs_in_one_column(self):
        _, st = _styled([["A", "1"], ["A", "2"], ["B", "3"], ["B", "4"]])
        st.merge_equal_runs(0, skip=SKIP)
        for r, text in ((0, "A"), (2, "B")):
            cell = st.cell(r, 0)
            assert cell.is_merge_origin
            assert cell.span_height == 2
            assert cell.text == text
            assert len(cell.text_frame.paragraphs) == 1

    def it_preserves_the_origin_cells_run_formatting(self):
        _, st = _styled([["h", "x"], ["A", "1"], ["A", "2"]])
        run = st.cell(1, 0).text_frame.paragraphs[0].runs[0]
        run.font.bold = True
        run.font.size = Pt(17)
        st.merge_equal_runs(0, skip=SKIP)
        origin = st.cell(1, 0)
        runs = origin.text_frame.paragraphs[0].runs
        assert origin.text == "A"
        assert len(runs) == 1
        assert runs[0].font.bold is True
        assert runs[0].font.size == Pt(17)

    def it_does_not_merge_values_in_skip(self):
        _, st = _styled([["X", "1"], ["X", "2"], ["X", "3"]])
        st.merge_equal_runs(0, skip=SKIP)
        assert _col_texts(st) == ["X", "X", "X"]
        assert not any(st.cell(r, 0).is_merge_origin for r in range(3))

    def it_does_not_merge_blank_cells(self):
        _, st = _styled([["", "1"], ["", "2"], ["A", "3"]])
        st.merge_equal_runs(0, skip=SKIP)
        assert not st.cell(0, 0).is_merge_origin
        assert not st.cell(1, 0).is_spanned

    def it_does_not_treat_non_contiguous_rows_as_adjacent(self):
        _, st = _styled([["h", "x"], ["A", "1"], ["B", "2"], ["A", "3"]])
        st.merge_equal_runs(0, rows=[1, 3], skip=SKIP)
        assert _col_texts(st) == ["h", "A", "B", "A"]
        assert not any(st.cell(r, 0).is_merge_origin for r in range(4))
        assert not any(st.cell(r, 0).is_spanned for r in range(4))

    def it_restricts_merging_to_the_given_rows(self):
        _, st = _styled([["h", "x"], ["A", "1"], ["A", "2"], ["A", "3"]])
        st.merge_equal_runs(0, rows=[1, 2], skip=SKIP)
        assert st.cell(1, 0).span_height == 2
        assert st.cell(1, 0).text == "A"
        assert not st.cell(3, 0).is_spanned
        assert st.cell(3, 0).text == "A"

    def it_keeps_merged_text_single_across_a_save_and_reload(self):
        prs, st = _styled([["h", "x"], ["A", "1"], ["A", "2"], ["B", "3"]])
        st.merge_equal_runs(0, skip=SKIP)
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        shape = next(iter(Presentation(buf).slides[0].shapes))
        cell = shape.table.cell(1, 0)
        assert cell.text == "A"
        assert cell.is_merge_origin
        assert cell.span_height == 2


    def it_changes_nothing_when_the_run_holds_an_already_merged_cell(self):
        _, st = _styled([["h", "x"], ["A", "A"], ["A", "2"]])
        st.cell(1, 0).merge(st.cell(1, 1))  # -- horizontal merge; origin text is "A\nA" --
        st.cell(1, 0).text = "A"

        with pytest.raises(ValueError, match="merged cells"):
            st.merge_equal_runs(0, skip=SKIP)

        assert _col_texts(st) == ["h", "A", "A"]
        assert st.cell(1, 0).span_height == 1


class DescribeCell_merge_generic:
    def it_still_concatenates_text_of_different_cells(self):
        _, st = _styled([["A", "1"], ["B", "2"]])
        st.cell(0, 0).merge(st.cell(1, 0))
        assert st.cell(0, 0).text == "A\nB"
