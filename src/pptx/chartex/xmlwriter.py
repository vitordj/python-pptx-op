"""Builds the ``ppt/charts/chartEx1.xml`` payload for a chartEx chart.

All chartEx charts share one envelope (``cx:chartSpace > cx:chartData`` +
``cx:chart > cx:plotArea > cx:plotAreaRegion > cx:series``); a given chart type is
selected by the ``layoutId`` attribute of ``cx:series``. The per-type differences are
captured in :data:`_LAYOUT_SPEC` so a single writer covers the whole family:

================  =========  ==========  ============  ====================
layoutId          num-dim    axes        categories    series ``layoutPr``
================  =========  ==========  ============  ====================
waterfall         val        val + cat   flat          subtotals
funnel            val        cat         flat          (none)
treemap           size       (none)      hierarchical  parentLabelLayout
sunburst          size       (none)      hierarchical  (none)
boxWhisker        val        val + cat   flat          statistics
clusteredColumn   val        cat + val   flat          binning (histogram)
================  =========  ==========  ============  ====================

The writer only *produces* XML (we never round-trip chartEx back into the API), so the
data values are cached inline in ``cx:pt`` elements and no embedded Excel workbook is
required for the chart to render.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable
from xml.sax.saxutils import escape

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

if TYPE_CHECKING:
    from pptx.chartex.data import ChartExData
    from pptx.enum.chart import XL_CHART_EX_TYPE


class _LayoutSpec:
    """Per-``layoutId`` rules describing how that chart type differs from the envelope."""

    def __init__(
        self,
        num_dim_type: str,
        axes: tuple[str, ...],
        layout_pr: Callable[["ChartExData"], str] | None = None,
        hierarchical: bool = False,
    ) -> None:
        self.num_dim_type = num_dim_type  # "val" or "size"
        self.axes = axes  # subset/order of ("val", "cat")
        self.layout_pr = layout_pr  # builds inner XML of cx:layoutPr, or None
        self.hierarchical = hierarchical


def _waterfall_layout_pr(data: "ChartExData") -> str:
    """`cx:subtotals` marking which points are totals/subtotals.

    If ``data.subtotal_idxs`` is set, exactly those indices are marked (e.g. ``[n-1]``
    when only the closing bar is a total). Otherwise the first and last categories are
    treated as totals -- the common "Anterior ... Total" shape.
    """
    n = len(data.values)
    if data.subtotal_idxs is not None:
        idxs = sorted({i for i in data.subtotal_idxs if 0 <= i < n})
    else:
        idxs = sorted({0, n - 1}) if n else []
    inner = "".join(f'<cx:idx val="{i}"/>' for i in idxs)
    return f"<cx:subtotals>{inner}</cx:subtotals>"


def _treemap_layout_pr(_data: "ChartExData") -> str:
    return '<cx:parentLabelLayout val="overlapping"/>'


_LAYOUT_SPEC: dict[str, _LayoutSpec] = {
    "waterfall": _LayoutSpec("val", ("val", "cat"), _waterfall_layout_pr),
    "funnel": _LayoutSpec("val", ("cat",)),
    "treemap": _LayoutSpec("size", (), _treemap_layout_pr, hierarchical=True),
    "sunburst": _LayoutSpec("size", (), hierarchical=True),
    "boxWhisker": _LayoutSpec("val", ("val", "cat")),
    "clusteredColumn": _LayoutSpec("val", ("cat", "val")),
}


class ChartExXmlWriter:
    """Serializes a :class:`ChartExData` into a chartEx part for a given chart type."""

    def __init__(self, chart_type: "XL_CHART_EX_TYPE", chart_data: "ChartExData") -> None:
        self._chart_type = chart_type
        self._data = chart_data

    @property
    def _layout_id(self) -> str:
        return self._chart_type.xml_value

    @property
    def _spec(self) -> _LayoutSpec:
        try:
            return _LAYOUT_SPEC[self._layout_id]
        except KeyError:  # pragma: no cover - guarded at the public API too
            raise NotImplementedError(
                f"chartEx layoutId '{self._layout_id}' is not yet supported"
            )

    def xml_bytes(self) -> bytes:
        return self.xml.encode("utf-8")

    @property
    def xml(self) -> str:
        return (
            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
            f"<cx:chartSpace {nsdecls('cx', 'a', 'r')}>"
            f"{self._chart_data_xml}"
            f"{self._chart_xml}"
            "</cx:chartSpace>"
        )

    # -- cx:chartData ---------------------------------------------------------

    @property
    def _chart_data_xml(self) -> str:
        return (
            "<cx:chartData>"
            f'<cx:data id="0">{self._str_dim_xml}{self._num_dim_xml}</cx:data>'
            "</cx:chartData>"
        )

    @property
    def _str_dim_xml(self) -> str:
        levels = self._data.levels()
        lvl_xml = "".join(self._lvl_xml(level) for level in levels)
        return f'<cx:strDim type="cat">{lvl_xml}</cx:strDim>'

    @property
    def _num_dim_xml(self) -> str:
        pts = "".join(f'<cx:pt idx="{i}">{v}</cx:pt>' for i, v in enumerate(self._data.values))
        n = len(self._data.values)
        return (
            f'<cx:numDim type="{self._spec.num_dim_type}">'
            f'<cx:lvl ptCount="{n}" formatCode="General">{pts}</cx:lvl>'
            "</cx:numDim>"
        )

    @staticmethod
    def _lvl_xml(level: list[str]) -> str:
        pts = "".join(f'<cx:pt idx="{i}">{escape(str(v))}</cx:pt>' for i, v in enumerate(level))
        return f'<cx:lvl ptCount="{len(level)}">{pts}</cx:lvl>'

    # -- cx:chart -------------------------------------------------------------

    @property
    def _chart_xml(self) -> str:
        return (
            "<cx:chart><cx:plotArea><cx:plotAreaRegion><cx:plotSurface/>"
            f"{self._series_xml}"
            "</cx:plotAreaRegion>"
            f"{self._axes_xml}"
            "</cx:plotArea></cx:chart>"
        )

    @property
    def _series_xml(self) -> str:
        layout_pr = ""
        if self._spec.layout_pr is not None:
            layout_pr = f"<cx:layoutPr>{self._spec.layout_pr(self._data)}</cx:layoutPr>"
        return (
            f'<cx:series layoutId="{self._layout_id}" uniqueId="{{00000000-0000-0000-0000-000000000000}}" formatIdx="0">'
            f"<cx:tx><cx:txData><cx:v>{escape(self._data.series_name)}</cx:v></cx:txData></cx:tx>"
            f"{self._data_pts_xml}"
            f"<cx:dataLabels>{self._txpr_xml}<cx:visibility seriesName=\"0\" categoryName=\"0\" value=\"1\"/></cx:dataLabels>"
            '<cx:dataId val="0"/>'
            f"{layout_pr}"
            "</cx:series>"
        )

    @property
    def _data_pts_xml(self) -> str:
        """Per-point ``cx:dataPt`` fills (``data.point_colors``), e.g. green/red bars."""
        colors = self._data.point_colors
        if not colors:
            return ""
        return "".join(
            f'<cx:dataPt idx="{i}"><cx:spPr><a:solidFill>'
            f'<a:srgbClr val="{escape(str(c))}"/></a:solidFill></cx:spPr></cx:dataPt>'
            for i, c in enumerate(colors)
        )

    @property
    def _txpr_xml(self) -> str:
        """``cx:txPr`` carrying the configured font; empty when no font is set."""
        name, size = self._data.font_name, self._data.font_size
        if not name and not size:
            return ""
        sz = f' sz="{int(round((size or 9) * 100))}"'
        latin = f'<a:latin typeface="{escape(name)}"/>' if name else ""
        # smtId=4294967295 (0xFFFFFFFF) = "sem modificação de estilo", como o Aspose emite
        return (
            "<cx:txPr><a:bodyPr/><a:p>"
            f'<a:pPr><a:defRPr{sz} smtId="4294967295">{latin}</a:defRPr></a:pPr>'
            f'<a:endParaRPr{sz} smtId="4294967295">{latin}</a:endParaRPr>'
            "</a:p></cx:txPr>"
        )

    @property
    def _axes_xml(self) -> str:
        num_fmt = self._data.number_format
        parts: list[str] = []
        for axis_id, kind in enumerate(self._spec.axes):
            if kind == "val":
                scaling = "<cx:valScaling/>"
                # numFmt só faz sentido no eixo de valores
                numfmt_xml = (
                    f'<cx:numFmt formatCode="{escape(num_fmt)}" sourceLinked="0"/>'
                    if num_fmt else ""
                )
            else:
                scaling = '<cx:catScaling gapWidth="0.5"/>'
                numfmt_xml = ""
            parts.append(
                f'<cx:axis id="{axis_id}">{scaling}'
                "<cx:majorGridlines><cx:spPr><a:ln><a:noFill/></a:ln></cx:spPr></cx:majorGridlines>"
                "<cx:minorGridlines><cx:spPr><a:ln><a:noFill/></a:ln></cx:spPr></cx:minorGridlines>"
                f"<cx:tickLabels/>{numfmt_xml}{self._txpr_xml}</cx:axis>"
            )
        return "".join(parts)


def new_chartex_xml(chart_type: "XL_CHART_EX_TYPE", chart_data: "ChartExData") -> bytes:
    """Return the chartEx part XML bytes for `chart_type` depicting `chart_data`."""
    return ChartExXmlWriter(chart_type, chart_data).xml_bytes()


def _self_check_parses(chart_type: "XL_CHART_EX_TYPE", chart_data: "ChartExData"):
    """Parse the produced XML to catch malformed output early (used in tests)."""
    return parse_xml(new_chartex_xml(chart_type, chart_data))
