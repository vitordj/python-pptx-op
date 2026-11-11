"""Unit tests for `_Cell.borders` (per-cell border lines)."""

import io

import pytest
from lxml import etree

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_FILL, MSO_LINE_DASH_STYLE
from pptx.oxml.ns import qn
from pptx.table import _Cell
from pptx.util import Emu, Inches, Pt

from .unitutil.cxml import element

EDGES = [
    ("left", "lnL"),
    ("right", "lnR"),
    ("top", "lnT"),
    ("bottom", "lnB"),
    ("tl_to_br", "lnTlToBr"),
    ("bl_to_tr", "lnBlToTr"),
]
BORDER_ORDER = ["lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr"]


def _local_names(parent):
    return [child.tag.split("}")[1] for child in parent]


def _bare_cell():
    """A cell whose `a:tc` has no `a:tcPr` child."""
    return _Cell(element("a:tc/a:txBody/(a:bodyPr,a:lstStyle,a:p)"), None)


def _table(rows=3, cols=3):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    shape = slide.shapes.add_table(rows, cols, Inches(1), Inches(1), Inches(6), Inches(3))
    return prs, shape.table


def _reload(prs):
    stream = io.BytesIO()
    prs.save(stream)
    stream.seek(0)
    return Presentation(stream)


def _set_all(cell, names):
    for name in names:
        line = getattr(cell.borders, name)
        line.width = Pt(1.5)
        line.color.rgb = RGBColor(0x12, 0x34, 0x56)


class DescribeCellBorders:
    @pytest.mark.parametrize(("name", "tag"), EDGES)
    def it_writes_each_edge_to_its_own_element(self, name, tag):
        cell = _bare_cell()
        line = getattr(cell.borders, name)
        line.width = Pt(2)
        line.color.rgb = RGBColor(0x8C, 0xC6, 0x40)

        tcPr = cell._tc.tcPr
        assert _local_names(tcPr) == [tag]
        edge = tcPr.find(qn("a:" + tag))
        assert edge.get("w") == str(Pt(2)) == "25400"
        srgb = edge.find(qn("a:solidFill") + "/" + qn("a:srgbClr"))
        assert srgb.get("val") == "8CC640"

    @pytest.mark.parametrize(("name", "tag"), EDGES)
    def it_reads_width_and_color_back(self, name, tag):
        cell = _bare_cell()
        line = getattr(cell.borders, name)
        line.width = Pt(3)
        line.color.rgb = RGBColor(1, 2, 3)
        fresh = _Cell(cell._tc, None)  # -- new proxies, no cached state --
        assert getattr(fresh.borders, name).width == Pt(3)
        assert getattr(fresh.borders, name).color.rgb == RGBColor(1, 2, 3)
        assert getattr(fresh.borders, name).width == Emu(38100)

    def it_does_not_mutate_the_XML_when_reading_width_or_dash_style(self):
        cell = _bare_cell()
        before = etree.tostring(cell._tc)
        for name, _ in EDGES:
            line = getattr(cell.borders, name)
            assert line.width == 0
            assert line.dash_style is None
        assert cell._tc.tcPr is None
        assert etree.tostring(cell._tc) == before

    def it_does_not_create_tcPr_when_reading_from_a_cell_without_one(self):
        cell = _bare_cell()
        assert cell.borders.left.width == 0
        assert cell._tc.tcPr is None

    def it_applies_the_dash_style(self):
        cell = _bare_cell()
        cell.borders.bottom.dash_style = MSO_LINE_DASH_STYLE.DASH
        lnB = cell._tc.tcPr.find(qn("a:lnB"))
        assert lnB.find(qn("a:prstDash")).get("val") == "dash"
        assert _Cell(cell._tc, None).borders.bottom.dash_style == MSO_LINE_DASH_STYLE.DASH

    @pytest.mark.parametrize(("name", "tag"), EDGES)
    def it_writes_noFill_for_fill_background(self, name, tag):
        cell = _bare_cell()
        getattr(cell.borders, name).fill.background()
        edge = cell._tc.tcPr.find(qn("a:" + tag))
        assert _local_names(edge) == ["noFill"]
        assert getattr(cell.borders, name).fill.type == MSO_FILL.BACKGROUND

    def it_does_not_duplicate_an_edge_set_twice(self):
        cell = _bare_cell()
        cell.borders.top.width = Pt(1)
        cell.borders.top.width = Pt(4)
        cell.borders.top.color.rgb = RGBColor(1, 1, 1)
        cell.borders.top.color.rgb = RGBColor(2, 2, 2)
        cell.borders.top.fill.background()
        cell.borders.top.fill.solid()
        tcPr = cell._tc.tcPr
        assert _local_names(tcPr) == ["lnT"]
        assert tcPr.find(qn("a:lnT")).get("w") == str(Pt(4))
        assert len(tcPr.find(qn("a:lnT"))) == 1

    @pytest.mark.parametrize(
        "order",
        [
            BORDER_ORDER,
            list(reversed(BORDER_ORDER)),
            ["lnTlToBr", "lnL", "lnB", "lnBlToTr", "lnR", "lnT"],
            ["lnB", "lnT", "lnBlToTr", "lnR", "lnTlToBr", "lnL"],
        ],
    )
    @pytest.mark.parametrize("fill_first", [True, False])
    def it_keeps_borders_in_schema_order_before_the_fill(self, order, fill_first):
        by_tag = {tag: name for name, tag in EDGES}
        cell = _bare_cell()
        if fill_first:
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0xFF, 0, 0)
        _set_all(cell, [by_tag[tag] for tag in order])
        if not fill_first:
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0xFF, 0, 0)
        assert _local_names(cell._tc.tcPr) == BORDER_ORDER + ["solidFill"]

    def it_keeps_borders_before_a_no_fill_cell_fill(self):
        cell = _bare_cell()
        cell.fill.background()
        cell.borders.bottom.width = Pt(1)
        cell.borders.left.width = Pt(1)
        assert _local_names(cell._tc.tcPr) == ["lnL", "lnB", "noFill"]

    def it_preserves_borders_across_save_and_reload(self):
        prs, table = _table()
        cell = table.cell(1, 1)
        for i, (name, _) in enumerate(EDGES, start=1):
            line = getattr(cell.borders, name)
            line.width = Pt(i)
            line.color.rgb = RGBColor(i, i, i)
        cell.borders.top.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT
        cell.borders.right.fill.background()

        reloaded = _reload(prs)
        rcell = [s for s in reloaded.slides[0].shapes if s.has_table][0].table.cell(1, 1)
        for i, (name, _) in enumerate(EDGES, start=1):
            line = getattr(rcell.borders, name)
            assert line.width == Pt(i)
            if name != "right":
                assert line.color.rgb == RGBColor(i, i, i)
        assert rcell.borders.top.dash_style == MSO_LINE_DASH_STYLE.ROUND_DOT
        assert rcell.borders.right.fill.type == MSO_FILL.BACKGROUND
        assert _local_names(rcell._tc.tcPr) == BORDER_ORDER

    def it_works_on_a_merge_origin_cell(self):
        prs, table = _table()
        origin = table.cell(0, 0)
        origin.merge(table.cell(1, 1))
        assert origin.is_merge_origin
        origin.borders.left.width = Pt(2)
        origin.borders.bottom.color.rgb = RGBColor(0xAA, 0xBB, 0xCC)
        origin.fill.solid()
        origin.fill.fore_color.rgb = RGBColor(0, 0, 0)
        origin.borders.right.width = Pt(1)

        reloaded = _reload(prs)
        rtable = [s for s in reloaded.slides[0].shapes if s.has_table][0].table
        rorigin = rtable.cell(0, 0)
        assert rorigin.is_merge_origin
        assert rorigin.borders.left.width == Pt(2)
        assert rorigin.borders.right.width == Pt(1)
        assert rorigin.borders.bottom.color.rgb == RGBColor(0xAA, 0xBB, 0xCC)
        assert _local_names(rorigin._tc.tcPr) == ["lnL", "lnR", "lnB", "solidFill"]
        # -- spanned cells are untouched --
        assert rtable.cell(0, 1).borders.left.width == 0
