"""Unit-test suite for the `pptx.tablekit` module (fork extension).

Drives the public API on a real table and asserts against the produced XML, so the tests
also document the OOXML that each helper emits.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.text import MSO_VERTICAL_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.tablekit import (
    NO_LINE,
    BorderSpec,
    CellStyle,
    Edge,
    StyledTable,
    _distribute,
    _rgb,
    apply_borders,
    apply_cell_style,
    dataframe_to_table,
    min_row_height_for_font,
)
from pptx.util import Cm, Emu, Pt


def _slide():
    prs = Presentation()
    return prs, prs.slides.add_slide(prs.slide_layouts[6])


def _styled_table(rows=3, cols=3, strip_style=True):
    _, slide = _slide()
    return StyledTable.add(
        slide.shapes, rows, cols, Cm(1), Cm(1), Cm(12), Cm(3), strip_style=strip_style
    )


def _ln(cell, edge):
    """The `a:lnX` element for `edge` ("L", "R", "T", "B") of `cell`, or None."""
    return cell._tc.tcPr.find(qn("a:ln%s" % edge))


class Describe_rgb:
    @pytest.mark.parametrize("value", ["62993D", "#62993D", RGBColor(0x62, 0x99, 0x3D)])
    def it_coerces_hex_strings_and_RGBColor(self, value):
        assert _rgb(value) == RGBColor(0x62, 0x99, 0x3D)


class DescribeBorderSpec:
    def it_builds_a_box_with_four_equal_edges(self):
        edge = Edge("FF0000", 1.0)
        spec = BorderSpec.box(edge)
        assert (spec.left, spec.right, spec.top, spec.bottom) == (edge,) * 4

    @pytest.mark.parametrize(
        ("factory", "expected_top", "expected_bottom"),
        [
            ("underline", None, "edge"),
            ("hlines", "edge", "edge"),
            ("topline", "edge", NO_LINE),
        ],
    )
    def it_suppresses_side_edges_by_default(self, factory, expected_top, expected_bottom):
        edge = Edge()
        spec = getattr(BorderSpec, factory)(edge)

        assert spec.left is NO_LINE
        assert spec.right is NO_LINE
        assert spec.top == (edge if expected_top == "edge" else expected_top)
        assert spec.bottom == (edge if expected_bottom == "edge" else expected_bottom)

    @pytest.mark.parametrize("factory", ["underline", "hlines", "topline"])
    def but_it_leaves_side_edges_alone_when_asked(self, factory):
        spec = getattr(BorderSpec, factory)(Edge(), suppress_sides=False)
        assert spec.left is None
        assert spec.right is None


class DescribeCellStyle:
    def it_derives_a_copy_with_overrides(self):
        base = CellStyle(bold=True, fill="FFFFFF", size=9)
        derived = base.derive(fill="000000")

        assert derived.fill == "000000"
        assert derived.bold is True
        assert derived.size == 9
        assert base.fill == "FFFFFF"


class Describe_apply_borders:
    def it_draws_an_edge_as_a_solid_line(self):
        cell = _styled_table().cell(0, 0)

        apply_borders(cell, BorderSpec(bottom=Edge("D3D3D3", 1.5, MSO_LINE_DASH_STYLE.DASH)))

        lnB = _ln(cell, "B")
        assert lnB.get("w") == str(Pt(1.5))
        assert lnB.find(qn("a:solidFill")).find(qn("a:srgbClr")).get("val") == "D3D3D3"
        assert lnB.find(qn("a:prstDash")).get("val") == "dash"

    def it_writes_noFill_for_NO_LINE(self):
        cell = _styled_table().cell(0, 0)

        apply_borders(cell, BorderSpec(left=NO_LINE))

        assert _ln(cell, "L").find(qn("a:noFill")) is not None

    def it_leaves_None_edges_untouched(self):
        cell = _styled_table().cell(0, 0)

        apply_borders(cell, BorderSpec(top=Edge()))

        assert _ln(cell, "T") is not None
        tcPr = cell._tc.tcPr
        assert all(tcPr.find(qn("a:ln%s" % e)) is None for e in "LRB")


class Describe_apply_cell_style:
    def it_applies_every_set_aspect(self):
        cell = _styled_table().cell(0, 0)
        cell.text = "Header"
        style = CellStyle(
            font_name="Arial",
            size=9,
            bold=True,
            italic=False,
            color="FFFFFF",
            fill="62993D",
            align="center",
            anchor="middle",
            margins_cm=(0.1, 0.2, None, 0.05),
        )

        apply_cell_style(cell, style)

        assert cell.fill.fore_color.rgb == RGBColor(0x62, 0x99, 0x3D)
        assert cell.vertical_anchor == MSO_VERTICAL_ANCHOR.MIDDLE
        assert cell.margin_left == Cm(0.1)
        assert cell.margin_right == Cm(0.2)
        assert cell.margin_bottom == Cm(0.05)
        paragraph = cell.text_frame.paragraphs[0]
        assert paragraph.alignment == PP_ALIGN.CENTER
        for font in (paragraph.font, paragraph.runs[0].font):
            assert font.name == "Arial"
            assert font.size == Pt(9)
            assert font.bold is True
            assert font.italic is False
            assert font.color.rgb == RGBColor(0xFF, 0xFF, 0xFF)

    def it_removes_the_fill_for_the_literal_none(self):
        cell = _styled_table().cell(0, 0)
        cell.fill.solid()

        apply_cell_style(cell, CellStyle(fill="none"))

        assert cell._tc.tcPr.find(qn("a:noFill")) is not None

    def it_does_nothing_for_an_empty_style(self):
        cell = _styled_table().cell(0, 0)
        cell.text = "x"

        apply_cell_style(cell, CellStyle())

        assert cell._tc.tcPr is None or len(cell._tc.tcPr) == 0
        rPr = cell.text_frame.paragraphs[0].runs[0]._r.rPr
        assert rPr is None or (len(rPr) == 0 and dict(rPr.attrib) in ({}, {"lang": "en-US"}))


class Describe_distribute:
    @pytest.mark.parametrize(
        ("total", "weights", "expected"),
        [
            (100, [1, 1, 1], [33, 33, 34]),
            (100, [1, 3], [25, 75]),
            (10, [0, 0, 0], [3, 3, 4]),
        ],
    )
    def it_splits_a_total_proportionally_and_exactly(self, total, weights, expected):
        result = _distribute(total, weights)
        assert result == expected
        assert sum(result) == total


class Describe_min_row_height_for_font:
    def it_is_1_point_2_times_the_font_size_per_line(self):
        assert min_row_height_for_font(10) == Pt(12)
        assert min_row_height_for_font(10, lines=2) == Pt(24)

    def it_adds_padding(self):
        assert min_row_height_for_font(10, pad_cm=1) == Pt(12) + Cm(1)


class DescribeStyledTable:
    def it_strips_the_default_table_style_on_add(self):
        st = _styled_table()

        tblPr = st.table._tbl.tblPr
        assert len(tblPr) == 0
        assert "firstRow" not in tblPr.attrib
        assert "bandRow" not in tblPr.attrib

    def but_it_keeps_the_table_style_when_asked(self):
        st = _styled_table(strip_style=False)

        tblPr = st.table._tbl.tblPr
        assert tblPr.find(qn("a:tableStyleId")) is not None

    def it_exposes_the_wrapped_table_and_frame(self):
        st = _styled_table()
        assert st.graphic_frame.table._tbl is st.table._tbl
        assert st.cell(1, 2) == st.table.cell(1, 2)

    def it_sets_cell_text_coercing_to_str(self):
        st = _styled_table()

        st.set(0, 0, 42)
        st.set(0, 1, None)

        assert st.cell(0, 0).text == "42"
        assert st.cell(0, 1).text == ""

    def it_fills_rows_with_header_and_body_styles(self):
        st = _styled_table(rows=2, cols=2)

        st.fill_rows(
            [["A", "B"], ["1", "2"]],
            header_style=CellStyle(bold=True),
            body_style=CellStyle(bold=False),
        )

        assert st.cell(0, 1).text == "B"
        assert st.cell(0, 0).text_frame.paragraphs[0].runs[0].font.bold is True
        assert st.cell(1, 0).text_frame.paragraphs[0].runs[0].font.bold is False

    def it_can_treat_every_row_as_body(self):
        st = _styled_table(rows=2, cols=1)

        st.fill_rows(
            [["A"], ["B"]],
            header_style=CellStyle(bold=True),
            body_style=CellStyle(italic=True),
            first_row_is_header=False,
        )

        assert st.cell(0, 0).text_frame.paragraphs[0].runs[0].font.italic is True

    def it_styles_a_range_row_and_column(self):
        st = _styled_table(rows=3, cols=3)
        red = CellStyle(fill="FF0000")

        st.style_range(2, 2, 1, 1, red)  # -- corners in any order --
        st.style_row(0, CellStyle(fill="00FF00"), cols=[0])
        st.style_col(0, CellStyle(fill="0000FF"), rows=[2])

        assert st.cell(1, 1).fill.fore_color.rgb == RGBColor(0xFF, 0, 0)
        assert st.cell(2, 2).fill.fore_color.rgb == RGBColor(0xFF, 0, 0)
        assert st.cell(0, 0).fill.fore_color.rgb == RGBColor(0, 0xFF, 0)
        assert st.cell(2, 0).fill.fore_color.rgb == RGBColor(0, 0, 0xFF)
        assert (
            st.cell(0, 1)._tc.tcPr is None or st.cell(0, 1)._tc.tcPr.find(qn("a:solidFill")) is None
        )

    def it_merges_equal_runs_in_a_column(self):
        st = _styled_table(rows=5, cols=1)
        st.fill_rows([["A"], ["A"], [""], [""], ["B"]], first_row_is_header=False)

        st.merge_equal_runs(0)

        assert st.cell(0, 0).is_merge_origin
        assert st.cell(0, 0).span_height == 2
        # -- blanks are skipped by default --
        assert not st.cell(2, 0).is_merge_origin
        assert not st.cell(3, 0).is_spanned

    def it_respects_a_custom_skip_list(self):
        st = _styled_table(rows=3, cols=1)
        st.fill_rows([["Sub"], ["Sub"], ["X"]], first_row_is_header=False)

        st.merge_equal_runs(0, skip=("", "Sub"))

        assert not st.cell(0, 0).is_merge_origin

    def it_sets_column_widths(self):
        st = _styled_table(rows=1, cols=2)

        st.col_widths([3, 5])

        assert [c.width for c in st.table.columns] == [Cm(3), Cm(5)]
        assert st.graphic_frame.width == Cm(8)

    def it_fits_to_an_exact_total_width_and_height(self):
        st = _styled_table(rows=2, cols=3)
        st.col_widths([1, 2, 3])

        st.fit(width_cm=12, height_cm=4)

        widths = [c.width for c in st.table.columns]
        assert sum(widths) == Cm(12)
        assert widths[0] < widths[1] < widths[2]
        assert sum(r.height for r in st.table.rows) == Cm(4)

    def it_autosizes_rows_with_a_header_size(self):
        st = _styled_table(rows=3, cols=1)

        st.autosize_rows(font_pt=10, header_pt=20)

        heights = [r.height for r in st.table.rows]
        assert heights == [Pt(24), Pt(12), Pt(12)]

    def it_sets_a_uniform_row_height(self):
        st = _styled_table(rows=2, cols=1)

        st.set_uniform_row_height(1.5)

        assert all(r.height == Cm(1.5) for r in st.table.rows)

    def it_survives_a_round_trip(self):
        prs, slide = _slide()
        st = StyledTable.add(slide.shapes, 2, 2, Cm(1), Cm(1), Cm(10), Cm(2))
        st.fill_rows(
            [["H1", "H2"], ["a", "b"]],
            header_style=CellStyle(fill="62993D", borders=BorderSpec.box(Edge("FFFFFF", 1))),
        )

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        cell = Presentation(buf).slides[0].shapes[0].table.cell(0, 0)

        assert cell.text == "H1"
        assert cell.fill.fore_color.rgb == RGBColor(0x62, 0x99, 0x3D)
        assert cell.borders.top.width == Pt(1)


class Describe_dataframe_to_table:
    """`dataframe_to_table` only needs `.columns`, `.index.name` and `.itertuples()`."""

    class _FakeIndex:
        name = "idx"

    class _FakeFrame:
        def __init__(self, columns, rows):
            self.columns = columns
            self._rows = rows
            self.index = Describe_dataframe_to_table._FakeIndex()

        def itertuples(self, index=True):
            for i, row in enumerate(self._rows):
                yield ((i,) + tuple(row)) if index else tuple(row)

    def it_builds_a_table_from_a_dataframe(self):
        _, slide = _slide()
        df = self._FakeFrame(["a", "b"], [(1, None), (3, 4.5)])

        st = dataframe_to_table(
            slide.shapes, df, Cm(1), Cm(1), Cm(8), Cm(3), header_style=CellStyle(bold=True)
        )

        texts = [[c.text for c in row.cells] for row in st.table.rows]
        assert texts == [["a", "b"], ["1", ""], ["3", "4.5"]]
        assert st.cell(0, 0).text_frame.paragraphs[0].runs[0].font.bold is True

    def it_can_include_the_index_and_omit_the_header(self):
        _, slide = _slide()
        df = self._FakeFrame(["a"], [(1,), (2,)])

        st = dataframe_to_table(
            slide.shapes, df, Cm(1), Cm(1), Cm(8), Cm(3), include_header=False, index=True
        )

        texts = [[c.text for c in row.cells] for row in st.table.rows]
        assert texts == [["0", "1"], ["1", "2"]]

    def it_returns_a_table_whose_height_is_emu(self):
        _, slide = _slide()
        st = dataframe_to_table(slide.shapes, self._FakeFrame(["a"], [(1,)]), 0, 0, Cm(2), Cm(2))
        assert isinstance(st.table.rows[0].height, Emu)
