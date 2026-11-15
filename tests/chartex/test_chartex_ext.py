"""Unit-test suite for the `pptx.chartex` package (fork extension).

Covers the |ChartExData| model, the XML writer for every supported `layoutId`, the part
and its required style/colour sidecars, and `shapes.add_chartex()` end to end.
"""

from __future__ import annotations

import io
import zipfile

import pytest

from pptx import Presentation
from pptx.chartex import ChartExData
from pptx.chartex.part import ChartExColorsPart, ChartExPart, ChartExStylePart
from pptx.chartex.xmlwriter import _self_check_parses, new_chartex_xml
from pptx.enum.chart import XL_CHART_EX_TYPE
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Cm

SUPPORTED = [
    XL_CHART_EX_TYPE.WATERFALL,
    XL_CHART_EX_TYPE.FUNNEL,
    XL_CHART_EX_TYPE.TREEMAP,
    XL_CHART_EX_TYPE.SUNBURST,
    XL_CHART_EX_TYPE.BOX_WHISKER,
    XL_CHART_EX_TYPE.CLUSTERED_COLUMN,
]
HIERARCHICAL = (XL_CHART_EX_TYPE.TREEMAP, XL_CHART_EX_TYPE.SUNBURST)


def _flat_data():
    data = ChartExData()
    data.categories = ["Opening", "Q1", "Q2", "Closing"]
    data.add_series("Cash", (100, 25, -12, 113))
    return data


def _tree_data():
    data = ChartExData()
    data.categories = [("A", "a1"), ("A", "a2"), ("B", "b1")]
    data.add_series("Size", (3, 2, 5))
    return data


def _data_for(chart_type):
    return _tree_data() if chart_type in HIERARCHICAL else _flat_data()


class DescribeChartExData:
    def it_starts_empty(self):
        data = ChartExData()
        assert data.categories == []
        assert data.values == []
        assert data.series_name == "Series 1"
        assert data.depth == 1
        assert data.is_hierarchical is False
        assert data.levels() == [[]]

    def it_stores_a_single_series_as_floats(self):
        data = ChartExData()
        data.add_series("first", (1, 2))
        data.add_series("second", (3, 4, 5))

        assert data.series_name == "second"
        assert data.values == [3.0, 4.0, 5.0]
        assert all(isinstance(v, float) for v in data.values)

    def it_has_one_level_for_flat_categories(self):
        data = _flat_data()
        assert data.is_hierarchical is False
        assert data.depth == 1
        assert data.levels() == [["Opening", "Q1", "Q2", "Closing"]]

    def it_orders_hierarchical_levels_leaf_first(self):
        data = _tree_data()
        assert data.is_hierarchical is True
        assert data.depth == 2
        assert data.levels() == [["a1", "a2", "b1"], ["A", "A", "B"]]


class DescribeChartExXmlWriter:
    @pytest.mark.parametrize("chart_type", SUPPORTED)
    def it_produces_well_formed_xml_for_every_supported_type(self, chart_type):
        chartSpace = _self_check_parses(chart_type, _data_for(chart_type))

        assert chartSpace.tag == qn("cx:chartSpace")
        series = chartSpace.find(".//" + qn("cx:series"))
        assert series.get("layoutId") == chart_type.xml_value

    def it_caches_the_data_inline(self):
        chartSpace = _self_check_parses(XL_CHART_EX_TYPE.FUNNEL, _flat_data())

        numDim = chartSpace.find(".//" + qn("cx:numDim"))
        assert numDim.get("type") == "val"
        assert [pt.text for pt in numDim.iter(qn("cx:pt"))] == ["100.0", "25.0", "-12.0", "113.0"]
        strDim = chartSpace.find(".//" + qn("cx:strDim"))
        assert [pt.text for pt in strDim.iter(qn("cx:pt"))] == ["Opening", "Q1", "Q2", "Closing"]

    def it_uses_a_size_dimension_and_one_level_per_depth_for_hierarchies(self):
        chartSpace = _self_check_parses(XL_CHART_EX_TYPE.TREEMAP, _tree_data())

        assert chartSpace.find(".//" + qn("cx:numDim")).get("type") == "size"
        assert len(chartSpace.find(".//" + qn("cx:strDim")).findall(qn("cx:lvl"))) == 2
        assert chartSpace.find(".//" + qn("cx:axis")) is None
        assert chartSpace.find(".//" + qn("cx:parentLabelLayout")) is not None

    def it_marks_the_first_and_last_waterfall_points_as_subtotals(self):
        chartSpace = _self_check_parses(XL_CHART_EX_TYPE.WATERFALL, _flat_data())

        subtotals = chartSpace.find(".//" + qn("cx:subtotals"))
        assert [idx.get("val") for idx in subtotals] == ["0", "3"]

    @pytest.mark.parametrize(
        ("chart_type", "n_axes"),
        [
            (XL_CHART_EX_TYPE.WATERFALL, 2),
            (XL_CHART_EX_TYPE.FUNNEL, 1),
            (XL_CHART_EX_TYPE.BOX_WHISKER, 2),
            (XL_CHART_EX_TYPE.CLUSTERED_COLUMN, 2),
            (XL_CHART_EX_TYPE.SUNBURST, 0),
        ],
    )
    def it_emits_the_axes_each_type_needs(self, chart_type, n_axes):
        chartSpace = _self_check_parses(chart_type, _data_for(chart_type))
        assert len(chartSpace.findall(".//" + qn("cx:axis"))) == n_axes

    def it_escapes_markup_in_labels_and_series_name(self):
        data = ChartExData()
        data.categories = ["R&D", "<other>"]
        data.add_series("P&L", (1, 2))

        chartSpace = _self_check_parses(XL_CHART_EX_TYPE.FUNNEL, data)

        assert chartSpace.find(".//" + qn("cx:v")).text == "P&L"
        labels = [pt.text for pt in chartSpace.find(".//" + qn("cx:strDim")).iter(qn("cx:pt"))]
        assert labels == ["R&D", "<other>"]

    @pytest.mark.parametrize(
        "chart_type", [XL_CHART_EX_TYPE.PARETO_LINE, XL_CHART_EX_TYPE.REGION_MAP]
    )
    def it_raises_on_a_type_it_cannot_emit_yet(self, chart_type):
        with pytest.raises(NotImplementedError, match=chart_type.xml_value):
            new_chartex_xml(chart_type, _flat_data())


class DescribeChartExPart:
    def it_creates_the_part_with_its_style_and_colour_sidecars(self):
        prs = Presentation()
        package = prs.part.package

        part = ChartExPart.new(XL_CHART_EX_TYPE.WATERFALL, _flat_data(), package)

        assert part.content_type == CT.OFC_CHART_EX
        assert part.partname == "/ppt/charts/chartEx1.xml"
        style = part.part_related_by(RT.CHART_STYLE)
        colors = part.part_related_by(RT.CHART_COLOR_STYLE)
        assert isinstance(style, ChartExStylePart)
        assert style.content_type == CT.OFC_CHART_STYLE
        assert isinstance(colors, ChartExColorsPart)
        assert colors.content_type == CT.OFC_CHART_COLORS


class DescribeShapes_add_chartex:
    def it_adds_a_chartex_graphic_frame_to_the_slide(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        frame = slide.shapes.add_chartex(
            XL_CHART_EX_TYPE.WATERFALL, Cm(1), Cm(2), Cm(10), Cm(6), _flat_data()
        )

        assert (frame.left, frame.top, frame.width, frame.height) == (
            Cm(1),
            Cm(2),
            Cm(10),
            Cm(6),
        )
        cx_chart = frame._element.find(".//" + qn("cx:chart"))
        rId = cx_chart.get(qn("r:id"))
        assert slide.part.rels[rId].reltype == RT.CHART_EX

    @pytest.mark.parametrize("chart_type", SUPPORTED)
    def it_writes_a_package_that_reopens(self, chart_type):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_chartex(chart_type, Cm(1), Cm(1), Cm(10), Cm(6), _data_for(chart_type))

        buf = io.BytesIO()
        prs.save(buf)

        names = zipfile.ZipFile(buf).namelist()
        assert "ppt/charts/chartEx1.xml" in names
        assert "ppt/charts/style1.xml" in names
        assert "ppt/charts/colors1.xml" in names
        buf.seek(0)
        assert len(Presentation(buf).slides[0].shapes) == 1
