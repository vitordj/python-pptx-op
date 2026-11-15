"""Unit and acceptance tests for build-time validation of |ChartExData|."""

from __future__ import annotations

import io
import zipfile

import pytest

from pptx import Presentation
from pptx.chartex.data import ChartExData
from pptx.chartex.xmlwriter import new_chartex_xml
from pptx.enum.chart import XL_CHART_EX_TYPE
from pptx.oxml import parse_xml
from pptx.util import Inches

_CX = "http://schemas.microsoft.com/office/drawing/2014/chartex"


def _data(categories, values):
    data = ChartExData()
    data.categories = categories
    data.add_series("S", values)
    return data


def _add(slide, chart_type, data):
    return slide.shapes.add_chartex(chart_type, Inches(1), Inches(1), Inches(5), Inches(4), data)


def _saved_names(prs):
    stream = io.BytesIO()
    prs.save(stream)
    with zipfile.ZipFile(io.BytesIO(stream.getvalue())) as zf:
        return zf.namelist()


class DescribeChartExDataValidate:
    """Unit-test suite for `pptx.chartex.data.ChartExData.validate()`."""

    def it_reports_more_categories_than_values(self):
        data = _data(["a", "b", "c"], [1, 2])
        with pytest.raises(ValueError, match=r"3 categories.*2 values"):
            data.validate()

    def it_reports_more_values_than_categories(self):
        data = _data(["a", "b"], [1, 2, 3])
        with pytest.raises(ValueError, match=r"2 categories.*3 values"):
            data.validate()

    def it_reports_the_index_and_value_of_a_nan(self):
        data = _data(["a", "b", "c"], [1.0, float("nan"), 3.0])
        with pytest.raises(ValueError, match=r"index 1.*nan"):
            data.validate()

    def it_reports_the_index_and_value_of_a_positive_inf(self):
        data = _data(["a", "b", "c"], [1.0, 2.0, float("inf")])
        with pytest.raises(ValueError, match=r"index 2.*inf"):
            data.validate()

    def it_reports_the_index_and_value_of_a_negative_inf(self):
        data = _data(["a", "b", "c"], [float("-inf"), 2.0, 3.0])
        with pytest.raises(ValueError, match=r"index 0.*-inf"):
            data.validate()

    def it_reports_only_the_first_non_finite_value(self):
        data = _data(["a", "b", "c"], [1.0, float("inf"), float("nan")])
        with pytest.raises(ValueError, match=r"index 1.*inf"):
            data.validate()

    def it_reports_ragged_tuples_with_index_and_depths(self):
        data = _data([("a", "x"), ("a", "y"), ("b", "x", "z")], [1, 2, 3])
        with pytest.raises(ValueError, match=r"category 2.*depth 3.*depth 2"):
            data.validate()

    def it_reports_a_shallower_tuple_than_the_first(self):
        data = _data([("a", "x", "z"), ("b", "y")], [1, 2])
        with pytest.raises(ValueError, match=r"category 1.*depth 2.*depth 3"):
            data.validate()

    def it_reports_a_str_after_tuples(self):
        data = _data([("a", "x"), "b"], [1, 2])
        with pytest.raises(ValueError, match=r"mix"):
            data.validate()

    def it_reports_a_tuple_after_strs(self):
        data = _data(["a", ("b", "x")], [1, 2])
        with pytest.raises(ValueError, match=r"mix"):
            data.validate()

    @pytest.mark.parametrize(
        ("categories", "values"),
        [
            (["a", "b", "c"], [1, -2.5, 0]),
            ([("r1", "l1"), ("r1", "l2"), ("r2", "l3")], [1, 2, 3]),
            ([("r", "m", "l")], [5]),
        ],
    )
    def it_returns_None_for_valid_data(self, categories, values):
        assert _data(categories, values).validate() is None

    def it_does_not_mutate_the_data(self):
        data = _data([("a", "x"), ("b", "y")], [1, 2])
        data.validate()
        assert data.categories == [("a", "x"), ("b", "y")]
        assert data.values == [1.0, 2.0]
        assert data.series_name == "S"


class DescribeNewChartExXml:
    """Build-time validation through the public API."""

    @pytest.mark.parametrize(
        ("categories", "values"),
        [
            (["a", "b", "c"], [1, 2]),
            (["a", "b"], [1, float("nan")]),
            ([("a", "x"), ("b",)], [1, 2]),
            (["a", ("b", "x")], [1, 2]),
        ],
    )
    def it_raises_and_adds_nothing_when_the_data_is_bad(self, categories, values):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        data = _data(categories, values)

        with pytest.raises(ValueError):
            _add(slide, XL_CHART_EX_TYPE.WATERFALL, data)

        assert len(slide.shapes) == 0
        assert not [n for n in _saved_names(prs) if n.startswith("ppt/charts/")]

    def it_raises_from_new_chartex_xml_directly(self):
        with pytest.raises(ValueError, match="3 categories"):
            new_chartex_xml(XL_CHART_EX_TYPE.FUNNEL, _data(["a", "b", "c"], [1, 2]))

    @pytest.mark.parametrize(
        ("chart_type", "categories"),
        [
            (XL_CHART_EX_TYPE.WATERFALL, ["a", "b", "c"]),
            (XL_CHART_EX_TYPE.FUNNEL, ["a", "b", "c"]),
            (XL_CHART_EX_TYPE.BOX_WHISKER, ["a", "b", "c"]),
            (XL_CHART_EX_TYPE.CLUSTERED_COLUMN, ["a", "b", "c"]),
            (XL_CHART_EX_TYPE.TREEMAP, [("r1", "a"), ("r1", "b"), ("r2", "c")]),
            (XL_CHART_EX_TYPE.SUNBURST, [("r1", "a"), ("r1", "b"), ("r2", "c")]),
        ],
    )
    def it_still_builds_a_chart_for_good_data(self, chart_type, categories):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        _add(slide, chart_type, _data(categories, [1, 2, 3]))

        assert len(slide.shapes) == 1
        names = _saved_names(prs)
        assert [n for n in names if n.startswith("ppt/charts/chartEx")]

    def it_writes_exactly_the_given_values(self):
        values = [0.13, -0.02, 1e-7, 250000.5]
        xml = new_chartex_xml(XL_CHART_EX_TYPE.WATERFALL, _data(["a", "b", "c", "d"], values))
        root = parse_xml(xml)
        pts = root.xpath("//cx:numDim//cx:pt", namespaces={"cx": _CX})
        assert [float(pt.text) for pt in pts] == values
