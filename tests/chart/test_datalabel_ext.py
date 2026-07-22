"""Integration tests for XyDataLabels cell-linked scatter labels.

Exercises the OOXML writer end-to-end on a real scatter chart, asserting the
`c15:datalabelsRange`, native `c:dLbls` flags, per-point `c:dLbl` manual layout,
and round-trip preservation of unknown extensions.
"""

import pytest

from pptx import Presentation
from pptx.chart.data import BubbleChartData, XyChartData
from pptx.chart.datalabel import XyDataLabels
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Inches

LABELS = ["Alpha", "Beta", "Gamma"]
XS = [0.20, 0.201, 0.24]  # Alpha & Beta cluster; Gamma isolated
YS = [0.03, 0.031, -0.08]


def _scatter_series(chart_type=XL_CHART_TYPE.XY_SCATTER):
    data = XyChartData()
    series = data.add_series("S")
    for x, y in zip(XS, YS):
        series.add_data_point(x, y)
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    gframe = slide.shapes.add_chart(chart_type, Inches(1), Inches(1), Inches(5), Inches(4), data)
    return gframe.chart.series[0]


class DescribeXyDataLabelsSetValuesFromCells:
    def it_is_the_data_labels_type_for_a_scatter_series(self):
        assert isinstance(_scatter_series().data_labels, XyDataLabels)

    def it_writes_a_datalabels_range_with_a_value_cache(self):
        series = _scatter_series()
        series.data_labels.set_values_from_cells(LABELS)
        data_range = series._element.find(
            "%s/%s/%s" % (qn("c:extLst"), qn("c:ext"), qn("c15:datalabelsRange"))
        )
        assert data_range is not None
        assert data_range.find(qn("c15:f")).text == "Sheet1!$C$2:$C$4"
        cache = data_range.find(qn("c15:dlblRangeCache"))
        assert cache.find(qn("c:ptCount")).get("val") == "3"
        assert [pt.find(qn("c:v")).text for pt in cache.findall(qn("c:pt"))] == LABELS

    def it_configures_native_dLbls_flags_and_leader_lines(self):
        dLbls = _scatter_series().data_labels
        dLbls.set_values_from_cells(LABELS, show_leader_lines=True)
        el = dLbls._element
        assert el.find(qn("c:showVal")).get("val") == "0"
        assert el.find(qn("c:showLeaderLines")).get("val") == "1"
        ext = el.find("%s/%s" % (qn("c:extLst"), qn("c:ext")))
        assert ext.find(qn("c15:showDataLabelsRange")).get("val") == "1"
        assert ext.find(qn("c15:showLeaderLines")).get("val") == "1"

    def it_adds_a_dLbl_only_for_offset_points(self):
        series = _scatter_series()
        offsets = [(0.05, -0.06), None, (0.0, 0.0)]
        series.data_labels.set_values_from_cells(LABELS, layout_offsets=offsets)
        dLbls = series.data_labels._element
        dLbl_lst = dLbls.findall(qn("c:dLbl"))
        assert len(dLbl_lst) == 1
        dLbl = dLbl_lst[0]
        assert dLbl.find(qn("c:idx")).get("val") == "0"
        manual = dLbl.find("%s/%s" % (qn("c:layout"), qn("c:manualLayout")))
        assert float(manual.find(qn("c:x")).get("val")) == pytest.approx(0.05)
        assert float(manual.find(qn("c:y")).get("val")) == pytest.approx(-0.06)
        fld = dLbl.find("%s/%s/%s/%s" % (qn("c:tx"), qn("c:rich"), qn("a:p"), qn("a:fld")))
        assert fld.get("type") == "CELLRANGE"

    def it_can_switch_leader_lines_off(self):
        dLbls = _scatter_series().data_labels
        dLbls.set_values_from_cells(LABELS, show_leader_lines=False)
        assert dLbls._element.find(qn("c:showLeaderLines")).get("val") == "0"

    def it_preserves_unknown_extensions_on_round_trip(self):
        series = _scatter_series()
        extLst = series._element.get_or_add_extLst()
        extLst.append(parse_xml('<c:ext %s uri="{DEAD-BEEF}"><c:foo/></c:ext>' % nsdecls("c")))
        series.data_labels.set_values_from_cells(LABELS)
        uris = [e.get("uri") for e in series._element.find(qn("c:extLst"))]
        assert "{DEAD-BEEF}" in uris
        assert "{02D57815-91ED-43cb-92C2-25804820EDAC}" in uris

    def it_uses_the_requested_workbook_column(self):
        series = _scatter_series()
        series.data_labels.set_values_from_cells(LABELS, workbook_column=5)
        f = series._element.find(
            "%s/%s/%s/%s" % (qn("c:extLst"), qn("c:ext"), qn("c15:datalabelsRange"), qn("c15:f"))
        )
        assert f.text == "Sheet1!$E$2:$E$4"

    def it_rejects_a_label_count_mismatch(self):
        series = _scatter_series()
        with pytest.raises(ValueError):
            series.data_labels.set_values_from_cells(["only", "two"])

    def it_rejects_a_layout_offsets_length_mismatch(self):
        series = _scatter_series()
        with pytest.raises(ValueError):
            series.data_labels.set_values_from_cells(LABELS, layout_offsets=[None])

    def it_rejects_an_empty_label_list(self):
        series = _scatter_series()
        with pytest.raises(ValueError):
            series.data_labels.set_values_from_cells([])

    def it_is_not_supported_on_bubble_series(self):
        data = BubbleChartData()
        series = data.add_series("S")
        for x, y in zip(XS, YS):
            series.add_data_point(x, y, 1)
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        gframe = slide.shapes.add_chart(
            XL_CHART_TYPE.BUBBLE, Inches(1), Inches(1), Inches(5), Inches(4), data
        )
        with pytest.raises(NotImplementedError):
            gframe.chart.series[0].data_labels.set_values_from_cells(LABELS)


class DescribeXyDataLabelsApplyHybridScatterOffsets:
    def it_reads_series_values_and_applies_offsets(self):
        series = _scatter_series()
        series.data_labels.apply_hybrid_scatter_offsets(LABELS)
        dLbls = series.data_labels._element
        # Alpha & Beta cluster get offsets; Gamma is isolated (no c:dLbl).
        assert len(dLbls.findall(qn("c:dLbl"))) == 2
        assert (
            dLbls.find("%s/%s/%s" % (qn("c:extLst"), qn("c:ext"), qn("c15:showDataLabelsRange")))
            is not None
        )

    def it_forwards_layout_keywords_to_the_algorithm(self):
        series = _scatter_series()
        series.data_labels.apply_hybrid_scatter_offsets(LABELS, magnitude=0.065)
        idx0 = series.data_labels._element.find(qn("c:dLbl"))
        manual = idx0.find("%s/%s" % (qn("c:layout"), qn("c:manualLayout")))
        x = float(manual.find(qn("c:x")).get("val"))
        y = float(manual.find(qn("c:y")).get("val"))
        assert (x**2 + y**2) ** 0.5 == pytest.approx(0.065, abs=1e-6)
