"""Unit-test suite for the optional chartEx formatting on `ChartExData`.

Covers per-point colours, waterfall totals, the number format and the label font: the
values accepted by the setters, the XML the writer emits for them (including child order,
which PowerPoint enforces), and a save/reload through `shapes.add_chartex()`.
"""

from __future__ import annotations

import io

import pytest
from lxml import etree

from pptx import Presentation
from pptx.chartex import ChartExData
from pptx.chartex.xmlwriter import new_chartex_xml
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_EX_TYPE
from pptx.util import Cm

NS = {
    "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def _waterfall_data() -> ChartExData:
    data = ChartExData()
    data.categories = ["Anterior", "Compra A", "Venda B", "Total"]
    data.add_series("Atribuição", (0.10, 0.03, -0.02, 0.11))
    return data


def _root(data: ChartExData, chart_type=XL_CHART_EX_TYPE.WATERFALL):
    return etree.fromstring(new_chartex_xml(chart_type, data))


def _children(element) -> list[str]:
    return [etree.QName(child).localname for child in element]


class DescribeChartExData_formatting:
    """Setters and defaults of the optional formatting properties."""

    def it_has_no_formatting_by_default(self):
        data = ChartExData()
        assert data.point_colors == []
        assert data.subtotal_idxs is None
        assert data.number_format is None
        assert data.font_name is None
        assert data.font_size is None

    @pytest.mark.parametrize(
        ("given", "expected"),
        [
            (["92d050", "#FF0000", None], ["92D050", "FF0000", None]),
            ([RGBColor(0x12, 0xAB, 0xEF)], ["12ABEF"]),
            ((c for c in ["000000"]), ["000000"]),
        ],
    )
    def it_normalises_point_colors(self, given, expected):
        data = ChartExData()
        data.point_colors = given
        assert data.point_colors == expected

    def it_clears_point_colors_with_None(self):
        data = ChartExData()
        data.point_colors = ["FF0000"]
        data.point_colors = None
        assert data.point_colors == []

    @pytest.mark.parametrize("bad", ["red", "FF00", "#GG0000", "FF00000"])
    def it_rejects_a_point_color_that_is_not_six_hex_digits(self, bad):
        data = ChartExData()
        data.point_colors = ["00FF00"]
        with pytest.raises(ValueError, match="index 1"):
            data.point_colors = ["FF0000", bad]
        assert data.point_colors == ["00FF00"]

    def it_sorts_and_dedupes_subtotal_idxs(self):
        data = ChartExData()
        data.subtotal_idxs = [3, 0, 3]
        assert data.subtotal_idxs == [0, 3]
        data.subtotal_idxs = []
        assert data.subtotal_idxs == []
        data.subtotal_idxs = None
        assert data.subtotal_idxs is None

    @pytest.mark.parametrize("size", [0, 0.5, -9, 4001])
    def it_rejects_an_out_of_range_font_size(self, size):
        data = ChartExData()
        with pytest.raises(ValueError, match="font_size"):
            data.font_size = size

    def it_rejects_a_misspelled_attribute(self):
        data = ChartExData()
        with pytest.raises(AttributeError):
            data.number_fmt = "0%"  # pyright: ignore[reportAttributeAccessIssue]

    def it_rejects_subtotal_idxs_out_of_range_when_validated(self):
        data = _waterfall_data()
        data.subtotal_idxs = [3, 4]
        with pytest.raises(ValueError, match="subtotal index 4"):
            data.validate()

    def it_rejects_more_point_colors_than_values_when_validated(self):
        data = _waterfall_data()
        data.point_colors = ["00FF00"] * 5
        with pytest.raises(ValueError, match="5 point colors given for 4 values"):
            data.validate()


class DescribeChartExXmlWriter_formatting:
    """The XML emitted for each formatting option."""

    def it_writes_no_formatting_elements_by_default(self):
        root = _root(_waterfall_data())
        assert root.xpath("//cx:dataPt | //cx:numFmt | //cx:txPr", namespaces=NS) == []
        assert root.xpath("//cx:numDim/cx:lvl/@formatCode", namespaces=NS) == ["General"]

    def it_marks_the_first_and_last_points_as_totals_by_default(self):
        root = _root(_waterfall_data())
        assert root.xpath("//cx:subtotals/cx:idx/@val", namespaces=NS) == ["0", "3"]

    @pytest.mark.parametrize(
        ("idxs", "expected"),
        [([3], ["3"]), ([0, 2, 3], ["0", "2", "3"]), ([], [])],
    )
    def it_marks_only_the_given_totals(self, idxs, expected):
        data = _waterfall_data()
        data.subtotal_idxs = idxs
        root = _root(data)
        assert root.xpath("//cx:subtotals/cx:idx/@val", namespaces=NS) == expected

    def it_writes_one_dataPt_per_coloured_point(self):
        data = _waterfall_data()
        data.point_colors = ["8CC640", None, "FF0000"]
        root = _root(data)
        points = root.xpath("//cx:series/cx:dataPt", namespaces=NS)
        assert [p.get("idx") for p in points] == ["0", "2"]
        assert root.xpath("//cx:dataPt//a:srgbClr/@val", namespaces=NS) == ["8CC640", "FF0000"]

    def it_applies_the_number_format_to_values_axis_and_labels(self):
        data = _waterfall_data()
        data.number_format = '0.00%;[Red]-0.00%;"–"'
        root = _root(data)
        fmt = '0.00%;[Red]-0.00%;"–"'
        assert root.xpath("//cx:numDim/cx:lvl/@formatCode", namespaces=NS) == [fmt]
        assert root.xpath("//cx:dataLabels/cx:numFmt/@formatCode", namespaces=NS) == [fmt]
        val_axis = root.xpath("//cx:axis[cx:valScaling]", namespaces=NS)[0]
        cat_axis = root.xpath("//cx:axis[cx:catScaling]", namespaces=NS)[0]
        assert val_axis.xpath("cx:numFmt/@formatCode", namespaces=NS) == [fmt]
        assert val_axis.xpath("cx:numFmt/@sourceLinked", namespaces=NS) == ["0"]
        assert cat_axis.xpath("cx:numFmt", namespaces=NS) == []

    def it_applies_the_font_to_tick_labels_and_data_labels(self):
        data = _waterfall_data()
        data.font_name = "Montserrat"
        data.font_size = 9
        root = _root(data)
        tx_prs = root.xpath("//cx:txPr", namespaces=NS)
        # -- one on the data labels and one on each of the two waterfall axes --
        assert len(tx_prs) == 3
        assert set(root.xpath("//cx:txPr//a:latin/@typeface", namespaces=NS)) == {"Montserrat"}
        assert set(root.xpath("//cx:txPr//a:defRPr/@sz", namespaces=NS)) == {"900"}

    def it_can_set_the_font_size_alone(self):
        data = _waterfall_data()
        data.font_size = 10.5
        root = _root(data)
        assert root.xpath("//cx:txPr//a:latin", namespaces=NS) == []
        assert set(root.xpath("//cx:txPr//a:defRPr/@sz", namespaces=NS)) == {"1050"}

    def it_keeps_schema_child_order_with_every_option_set(self):
        data = _waterfall_data()
        data.point_colors = ["8CC640", "FF0000"]
        data.number_format = "0.0%"
        data.font_name = "Arial"
        data.font_size = 8
        root = _root(data)

        series = root.xpath("//cx:series", namespaces=NS)[0]
        assert _children(series) == ["tx", "dataPt", "dataPt", "dataLabels", "dataId", "layoutPr"]
        labels = root.xpath("//cx:dataLabels", namespaces=NS)[0]
        assert _children(labels) == ["numFmt", "txPr", "visibility"]
        val_axis = root.xpath("//cx:axis[cx:valScaling]", namespaces=NS)[0]
        assert _children(val_axis) == [
            "valScaling",
            "majorGridlines",
            "tickLabels",
            "numFmt",
            "txPr",
        ]
        cat_axis = root.xpath("//cx:axis[cx:catScaling]", namespaces=NS)[0]
        assert _children(cat_axis) == ["catScaling", "majorGridlines", "tickLabels", "txPr"]

    def it_escapes_quotes_and_markup_in_the_format_and_font(self):
        data = _waterfall_data()
        data.number_format = '"R$ "#,##0'
        data.font_name = 'A&B "Sans"'
        root = _root(data)
        assert root.xpath("//cx:dataLabels/cx:numFmt/@formatCode", namespaces=NS) == ['"R$ "#,##0']
        assert set(root.xpath("//a:latin/@typeface", namespaces=NS)) == {'A&B "Sans"'}

    @pytest.mark.parametrize(
        "chart_type",
        [XL_CHART_EX_TYPE.FUNNEL, XL_CHART_EX_TYPE.BOX_WHISKER, XL_CHART_EX_TYPE.CLUSTERED_COLUMN],
    )
    def it_formats_other_flat_chart_types_too(self, chart_type):
        data = _waterfall_data()
        data.point_colors = ["8CC640"]
        data.number_format = "0%"
        root = _root(data, chart_type)
        assert root.xpath("//cx:dataPt/@idx", namespaces=NS) == ["0"]
        assert root.xpath("//cx:dataLabels/cx:numFmt/@formatCode", namespaces=NS) == ["0%"]
        assert root.xpath("//cx:subtotals", namespaces=NS) == []


class DescribeShapes_add_chartex_formatting:
    def it_survives_a_save_and_reload(self):
        data = _waterfall_data()
        data.point_colors = ["8CC640", "8CC640", "FF0000", "8CC640"]
        data.subtotal_idxs = [3]
        data.number_format = "0.00%"
        data.font_name = "Montserrat"
        data.font_size = 9
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_chartex(XL_CHART_EX_TYPE.WATERFALL, Cm(1), Cm(1), Cm(16), Cm(9), data)

        stream = io.BytesIO()
        prs.save(stream)
        stream.seek(0)
        reopened = Presentation(stream)

        chartex_parts = [
            p
            for p in reopened.part.package.iter_parts()
            if p.partname.endswith(".xml")
            and "chartEx" in str(p.partname)
            and "style" not in str(p.partname)
            and "colors" not in str(p.partname)
        ]
        assert len(chartex_parts) == 1
        root = etree.fromstring(chartex_parts[0].blob)
        assert root.xpath("//cx:subtotals/cx:idx/@val", namespaces=NS) == ["3"]
        assert len(root.xpath("//cx:dataPt", namespaces=NS)) == 4
        assert root.xpath("//cx:dataLabels/cx:numFmt/@formatCode", namespaces=NS) == ["0.00%"]

    def it_adds_nothing_when_the_formatting_is_invalid(self):
        data = _waterfall_data()
        data.subtotal_idxs = [9]
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        n_parts = len(list(prs.part.package.iter_parts()))

        with pytest.raises(ValueError, match="subtotal index 9"):
            slide.shapes.add_chartex(XL_CHART_EX_TYPE.WATERFALL, Cm(1), Cm(1), Cm(16), Cm(9), data)

        assert len(slide.shapes) == 0
        assert len(list(prs.part.package.iter_parts())) == n_parts
