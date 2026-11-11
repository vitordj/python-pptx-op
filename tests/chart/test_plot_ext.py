"""Unit tests for DoughnutPlot.hole_size / .first_slice_angle (issue #493)."""

import pytest

from pptx.chart.data import CategoryChartData
from pptx.chart.plot import DoughnutPlot
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn
from pptx.util import Inches

from ..unitutil.cxml import element


def _doughnut_plot():
    data = CategoryChartData()
    data.categories = ["A", "B", "C"]
    data.add_series("S", (1, 2, 3))
    from pptx import Presentation

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    gframe = slide.shapes.add_chart(
        XL_CHART_TYPE.DOUGHNUT, Inches(1), Inches(1), Inches(5), Inches(4), data
    )
    return gframe.chart.plots[0]


class DescribeDoughnutPlotHoleSize:
    def it_defaults_to_10_when_absent(self):
        plot = DoughnutPlot(element("c:doughnutChart"), None)
        assert plot.hole_size == 10

    def it_reads_and_writes_hole_size(self):
        plot = _doughnut_plot()
        plot.hole_size = 60
        assert plot.hole_size == 60
        assert plot._element.find(qn("c:holeSize")).get("val") == "60"

    def it_reads_and_writes_first_slice_angle(self):
        plot = _doughnut_plot()
        plot.first_slice_angle = 90
        assert plot.first_slice_angle == 90
        assert plot._element.find(qn("c:firstSliceAng")).get("val") == "90"

    @pytest.mark.parametrize("bad", [0, 91, 200])
    def it_rejects_out_of_range_hole_size(self, bad):
        plot = _doughnut_plot()
        with pytest.raises(ValueError):
            plot.hole_size = bad

    @pytest.mark.parametrize("bad", [-1, 361])
    def it_rejects_out_of_range_angle(self, bad):
        plot = _doughnut_plot()
        with pytest.raises(ValueError):
            plot.first_slice_angle = bad


def _child_order(plot):
    """Local tag names of the `c:doughnutChart` children, in document order."""
    return [child.tag.split("}")[1] for child in plot._element]


def _plot_from(cxml):
    return DoughnutPlot(element(cxml), None)


class DescribeDoughnutChartChildOrder:
    def it_inserts_firstSliceAng_before_an_existing_holeSize(self):
        plot = _plot_from("c:doughnutChart/c:holeSize{val=50}")
        plot.first_slice_angle = 90
        assert _child_order(plot) == ["firstSliceAng", "holeSize"]

    def it_inserts_holeSize_after_an_existing_firstSliceAng(self):
        plot = _plot_from("c:doughnutChart/c:firstSliceAng{val=90}")
        plot.hole_size = 50
        assert _child_order(plot) == ["firstSliceAng", "holeSize"]

    @pytest.mark.parametrize("hole_first", [True, False])
    def it_orders_both_children_regardless_of_assignment_order(self, hole_first):
        plot = _plot_from("c:doughnutChart")
        if hole_first:
            plot.hole_size = 50
            plot.first_slice_angle = 90
        else:
            plot.first_slice_angle = 90
            plot.hole_size = 50
        assert _child_order(plot) == ["firstSliceAng", "holeSize"]
        assert plot.hole_size == 50
        assert plot.first_slice_angle == 90

    @pytest.mark.parametrize("hole_first", [True, False])
    def it_inserts_both_children_before_extLst(self, hole_first):
        plot = _plot_from("c:doughnutChart/c:extLst")
        if hole_first:
            plot.hole_size = 50
            plot.first_slice_angle = 90
        else:
            plot.first_slice_angle = 90
            plot.hole_size = 50
        assert _child_order(plot) == ["firstSliceAng", "holeSize", "extLst"]

    @pytest.mark.parametrize("hole_first", [True, False])
    def it_inserts_both_children_after_varyColors_ser_and_dLbls(self, hole_first):
        plot = _plot_from("c:doughnutChart/(c:varyColors,c:ser,c:ser,c:dLbls)")
        if hole_first:
            plot.hole_size = 50
            plot.first_slice_angle = 90
        else:
            plot.first_slice_angle = 90
            plot.hole_size = 50
        assert _child_order(plot) == [
            "varyColors",
            "ser",
            "ser",
            "dLbls",
            "firstSliceAng",
            "holeSize",
        ]

    def it_keeps_schema_order_when_firstSliceAng_is_removed_from_a_real_chart(self):
        plot = _doughnut_plot()
        doughnut = plot._element
        doughnut.remove(doughnut.find(qn("c:firstSliceAng")))
        assert "firstSliceAng" not in _child_order(plot)

        plot.first_slice_angle = 90

        assert _child_order(plot) == [
            "varyColors",
            "ser",
            "dLbls",
            "firstSliceAng",
            "holeSize",
        ]
        assert plot.first_slice_angle == 90

    def it_updates_existing_children_in_place_without_duplicating_them(self):
        plot = _plot_from("c:doughnutChart/(c:firstSliceAng{val=10},c:holeSize{val=20})")
        plot.first_slice_angle = 90
        plot.hole_size = 50
        assert _child_order(plot) == ["firstSliceAng", "holeSize"]
        assert plot._element.find(qn("c:firstSliceAng")).get("val") == "90"
        assert plot._element.find(qn("c:holeSize")).get("val") == "50"
