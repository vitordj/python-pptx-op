"""Unit-test suite for column formats and conditional styles in `pptx.tablekit`."""

from __future__ import annotations

import math
from decimal import Decimal

import pytest

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.tablekit import (
    BorderSpec,
    CellStyle,
    ColumnFormat,
    Edge,
    NumberFormat,
    StyledTable,
    StyleRule,
    dataframe_to_table,
)
from pptx.util import Cm, Pt

RED = CellStyle(color="C00000")
GREEN = CellStyle(color="00B050")
BODY = CellStyle(font_name="Arial", size=9, color="000000", fill="FFFFFF", align="center")
HEADER = CellStyle(bold=True, fill="1F4E79", color="FFFFFF")


def _styled_table(rows=3, cols=3):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    return StyledTable.add(slide.shapes, rows, cols, Cm(1), Cm(1), Cm(12), Cm(3))


def _font(cell):
    return cell.text_frame.paragraphs[0].runs[0].font


def _texts(st):
    return [[c.text for c in row.cells] for row in st.table.rows]


class DescribeCellStyle_overlay:
    def it_puts_the_set_fields_of_the_other_style_on_top(self):
        result = BODY.overlay(CellStyle(color="C00000", bold=True))
        assert result == BODY.derive(color="C00000", bold=True)

    def it_returns_itself_for_None_or_an_empty_style(self):
        assert BODY.overlay(None) is BODY
        assert BODY.overlay(CellStyle()) is BODY

    def it_replaces_borders_whole(self):
        base = CellStyle(borders=BorderSpec.box(Edge("000000", 1)))
        over = CellStyle(borders=BorderSpec(bottom=Edge("FF0000", 2)))
        assert base.overlay(over).borders == over.borders


class DescribeNumberFormat:
    @pytest.mark.parametrize(
        ("fmt", "value", "expected"),
        [
            (NumberFormat(), 1234.5, "1,234.50"),
            (NumberFormat(0), 1234.5, "1,234"),
            (NumberFormat(2, decimal_sep=",", thousands_sep="."), -1234567.891, "-1.234.567,89"),
            (NumberFormat(2, thousands_sep=""), 1234.5, "1234.50"),
            (NumberFormat(1, percent=True), 0.12345, "12.3%"),
            (NumberFormat(2, percent=True, decimal_sep=","), -0.0525, "-5,25%"),
            (NumberFormat(2, prefix="R$ "), -10, "-R$ 10.00"),
            (NumberFormat(1, suffix=" bp"), 12, "12.0 bp"),
            (NumberFormat(2), -0.001, "0.00"),
            (NumberFormat(2), Decimal("3.14159"), "3.14"),
            (NumberFormat(2), 7, "7.00"),
        ],
    )
    def it_formats_numbers(self, fmt, value, expected):
        assert fmt(value) == expected

    @pytest.mark.parametrize(("value", "expected"), [("n/a", "n/a"), (None, ""), (True, "True")])
    def it_passes_non_numbers_through(self, value, expected):
        assert NumberFormat()(value) == expected

    def it_rejects_negative_decimals(self):
        with pytest.raises(ValueError, match="decimals"):
            NumberFormat(-1)


class DescribeStyleRule:
    @pytest.mark.parametrize(
        ("rule", "value", "expected"),
        [
            (StyleRule.negative(RED), -0.01, True),
            (StyleRule.negative(RED), 0, False),
            (StyleRule.negative(RED), Decimal("-1"), True),
            (StyleRule.positive(GREEN), 3, True),
            (StyleRule.positive(GREEN), -3, False),
            (StyleRule.zero(RED), 0.0, True),
            (StyleRule.zero(RED), 1, False),
            (StyleRule.negative(RED), "-5", False),
            (StyleRule.negative(RED), True, False),
            (StyleRule.negative(RED), None, False),
            (StyleRule.negative(RED), float("nan"), False),
        ],
    )
    def it_matches_by_sign_and_skips_non_numbers(self, rule, value, expected):
        assert rule.matches(value) is expected

    def it_matches_a_custom_condition_on_the_raw_value(self):
        rule = StyleRule(lambda v: v == "Total", CellStyle(bold=True))
        assert rule.matches("Total") is True
        assert rule.matches("Sub") is False

    def it_never_matches_a_missing_value(self):
        rule = StyleRule(lambda v: True, RED)
        assert rule.matches(None) is False
        assert rule.matches(math.nan) is False


class DescribeColumnFormat:
    @pytest.mark.parametrize(
        ("formatter", "value", "expected"),
        [
            (None, 1.5, "1.5"),
            (",.2f", 1234.5, "1,234.50"),
            (".1%", 0.256, "25.6%"),
            ("R$ {:,.0f}", 2500, "R$ 2,500"),
            (lambda v: f"<{v}>", 3, "<3>"),
            (NumberFormat(1, decimal_sep=","), 2.25, "2,2"),
        ],
    )
    def it_formats_the_text(self, formatter, value, expected):
        assert ColumnFormat(formatter).text(value) == expected

    @pytest.mark.parametrize("value", [None, float("nan")])
    def it_writes_the_na_text_for_a_missing_value(self, value):
        calls = []
        fmt = ColumnFormat(lambda v: calls.append(v) or "x", na="–")
        assert fmt.text(value) == "–"
        assert calls == []

    def it_accepts_a_single_rule_and_stores_rules_as_a_tuple(self):
        rule = StyleRule.negative(RED)
        assert ColumnFormat(rules=rule).rules == (rule,)
        assert ColumnFormat(rules=[rule]).rules == (rule,)

    def it_layers_base_column_style_and_matching_rules_in_order(self):
        fmt = ColumnFormat(
            style=CellStyle(align="right"),
            rules=[
                StyleRule.negative(CellStyle(color="C00000", bold=True)),
                StyleRule(lambda v: v < -10, CellStyle(color="800000")),
            ],
        )
        assert fmt.style_for(5, BODY) == BODY.derive(align="right")
        assert fmt.style_for(-1, BODY) == BODY.derive(align="right", color="C00000", bold=True)
        assert fmt.style_for(-50, BODY) == BODY.derive(align="right", color="800000", bold=True)

    def it_returns_None_when_there_is_nothing_to_apply(self):
        assert ColumnFormat().style_for(1) is None
        assert ColumnFormat(rules=[StyleRule.negative(RED)]).style_for(1) is None
        assert ColumnFormat(rules=[StyleRule.negative(RED)]).style_for(-1) == RED


class DescribeStyledTable_fill_rows_with_column_formats:
    def it_formats_and_styles_body_cells_from_raw_values(self):
        st = _styled_table(rows=3, cols=3)
        data = [
            ["Ativo", "Peso", "Retorno"],
            ["A", 0.6, 0.0125],
            ["B", 0.4, -0.031],
        ]
        pct = NumberFormat(2, percent=True, decimal_sep=",")

        st.fill_rows(
            data,
            header_style=HEADER,
            body_style=BODY,
            column_formats={
                1: ColumnFormat(pct),
                2: ColumnFormat(pct, rules=[StyleRule.negative(RED), StyleRule.positive(GREEN)]),
            },
        )

        assert _texts(st) == [
            ["Ativo", "Peso", "Retorno"],
            ["A", "60,00%", "1,25%"],
            ["B", "40,00%", "-3,10%"],
        ]
        assert _font(st.cell(1, 2)).color.rgb == RGBColor.from_string("00B050")
        assert _font(st.cell(2, 2)).color.rgb == RGBColor.from_string("C00000")
        # -- the rest of the body style is still applied to formatted cells --
        assert _font(st.cell(2, 2)).name == "Arial"
        assert _font(st.cell(2, 2)).size == Pt(9)
        assert st.cell(2, 2).fill.fore_color.rgb == RGBColor.from_string("FFFFFF")
        # -- unformatted columns and the header are as before --
        assert _font(st.cell(2, 0)).color.rgb == RGBColor.from_string("000000")
        assert _font(st.cell(0, 2)).color.rgb == RGBColor.from_string("FFFFFF")
        assert _font(st.cell(0, 2)).bold is True

    def it_formats_every_row_when_there_is_no_header(self):
        st = _styled_table(rows=2, cols=1)
        st.fill_rows(
            [[-1], [2]],
            body_style=BODY,
            first_row_is_header=False,
            column_formats={0: ColumnFormat(",.1f", rules=[StyleRule.negative(RED)])},
        )
        assert _texts(st) == [["-1.0"], ["2.0"]]
        assert _font(st.cell(0, 0)).color.rgb == RGBColor.from_string("C00000")

    def it_writes_missing_values_with_the_body_style(self):
        st = _styled_table(rows=2, cols=1)
        st.fill_rows(
            [["x"], [None]],
            body_style=BODY,
            column_formats={0: ColumnFormat(",.2f", rules=[StyleRule.negative(RED)], na="–")},
        )
        assert st.cell(1, 0).text == "–"
        assert _font(st.cell(1, 0)).color.rgb == RGBColor.from_string("000000")

    @pytest.mark.parametrize("key", [3, -1, "Retorno"])
    def it_rejects_a_key_that_is_not_a_column_index(self, key):
        st = _styled_table(rows=2, cols=3)
        with pytest.raises(KeyError, match="column_formats key"):
            st.fill_rows([["a", "b", "c"]], column_formats={key: ColumnFormat()})
        assert st.cell(0, 0).text == ""


class Describe_dataframe_to_table_with_column_formats:
    class _FakeIndex:
        name = "Data"

    class _FakeFrame:
        def __init__(self, columns, rows):
            self.columns = columns
            self._rows = rows
            self.index = Describe_dataframe_to_table_with_column_formats._FakeIndex()

        def itertuples(self, index=True):
            for i, row in enumerate(self._rows):
                yield ((f"d{i}",) + tuple(row)) if index else tuple(row)

    def _slide(self):
        prs = Presentation()
        return prs.slides.add_slide(prs.slide_layouts[6])

    def it_formats_columns_by_label_from_raw_values(self):
        df = self._FakeFrame(["Fundo", "Retorno"], [("X", 0.021), ("Y", -0.004), ("Z", None)])

        st = dataframe_to_table(
            self._slide().shapes,
            df,
            Cm(1),
            Cm(1),
            Cm(8),
            Cm(3),
            body_style=BODY,
            column_formats={
                "Retorno": ColumnFormat(".1%", rules=[StyleRule.negative(RED)], na="-")
            },
        )

        assert _texts(st) == [["Fundo", "Retorno"], ["X", "2.1%"], ["Y", "-0.4%"], ["Z", "-"]]
        assert _font(st.cell(2, 1)).color.rgb == RGBColor.from_string("C00000")
        assert _font(st.cell(1, 1)).color.rgb == RGBColor.from_string("000000")

    def it_can_format_the_index_by_its_name(self):
        df = self._FakeFrame(["v"], [(1,), (2,)])
        st = dataframe_to_table(
            self._slide().shapes,
            df,
            Cm(1),
            Cm(1),
            Cm(8),
            Cm(3),
            index=True,
            column_formats={"Data": ColumnFormat(str.upper), "v": ColumnFormat(".1f")},
        )
        assert _texts(st) == [["Data", "v"], ["D0", "1.0"], ["D1", "2.0"]]

    def it_rejects_an_unknown_label(self):
        df = self._FakeFrame(["Retorno"], [(1,)])
        with pytest.raises(KeyError, match="'Retorn'"):
            dataframe_to_table(
                self._slide().shapes,
                df,
                Cm(1),
                Cm(1),
                Cm(8),
                Cm(3),
                column_formats={"Retorn": ColumnFormat()},
            )

    def it_works_with_a_real_DataFrame(self):
        pd = pytest.importorskip("pandas")
        df = pd.DataFrame({"Fundo": ["X", "Y"], "Retorno": [0.02, float("nan")]})
        st = dataframe_to_table(
            self._slide().shapes,
            df,
            Cm(1),
            Cm(1),
            Cm(8),
            Cm(3),
            column_formats={"Retorno": ColumnFormat(".0%", rules=[StyleRule.negative(RED)])},
        )
        assert _texts(st) == [["Fundo", "Retorno"], ["X", "2%"], ["Y", ""]]
