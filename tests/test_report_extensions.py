"""Regression coverage for report tables and waterfall formatting."""

from io import BytesIO
from zipfile import ZipFile

from lxml import etree

from pptx import Presentation
from pptx.chartex.data import ChartExData
from pptx.chartex.xmlwriter import ChartExXmlWriter
from pptx.enum.chart import XL_CHART_EX_TYPE
from pptx.tablekit import StyledTable
from pptx.util import Inches

NS = {
    "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def _waterfall_data():
    data = ChartExData()
    data.categories = ["Previous", "Contribution", "Total"]
    data.add_series("Return", [0.1, -0.02, 0.08])
    return data


class TestReportExtensions:
    def test_waterfall_keeps_default_totals(self):
        root = etree.fromstring(
            ChartExXmlWriter(XL_CHART_EX_TYPE.WATERFALL, _waterfall_data()).xml_bytes()
        )
        assert root.xpath("//cx:subtotals/cx:idx/@val", namespaces=NS) == ["0", "2"]

    def test_waterfall_serializes_format_colors_fonts_and_explicit_totals(self):
        data = _waterfall_data()
        data.subtotal_idxs = [2]
        data.point_colors = ["#8cc640", "FF0000", None]
        data.number_format = "0.00%"
        data.font_name = "Montserrat"
        data.font_size = 9
        root = etree.fromstring(ChartExXmlWriter(XL_CHART_EX_TYPE.WATERFALL, data).xml_bytes())
        assert root.xpath("//cx:subtotals/cx:idx/@val", namespaces=NS) == ["2"]
        assert root.xpath("//cx:dataPt//a:srgbClr/@val", namespaces=NS) == ["8CC640", "FF0000"]
        assert set(root.xpath("//cx:numFmt/@formatCode", namespaces=NS)) == {"0.00%"}
        assert root.xpath("//cx:numDim/cx:lvl/@formatCode", namespaces=NS) == ["0.00%"]
        assert set(root.xpath("//a:latin/@typeface", namespaces=NS)) == {"Montserrat"}
        assert set(root.xpath("//a:defRPr/@sz", namespaces=NS)) == {"900"}
        series = root.xpath("//cx:series", namespaces=NS)[0]
        assert [etree.QName(child).localname for child in series] == [
            "tx", "dataPt", "dataPt", "dataLabels", "dataId", "layoutPr"
        ]

    def test_merge_equal_runs_does_not_repeat_labels_after_roundtrip(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        table = slide.shapes.add_table(3, 1, 0, 0, Inches(4), Inches(2)).table
        for row in table.rows:
            row.cells[0].text = "Same issuer"
        StyledTable(table).merge_equal_runs(0)
        stream = BytesIO()
        prs.save(stream)
        stream.seek(0)
        reopened = Presentation(stream)
        assert reopened.slides[0].shapes[0].table.cell(0, 0).text == "Same issuer"
        stream.seek(0)
        with ZipFile(stream) as archive:
            assert archive.testzip() is None
