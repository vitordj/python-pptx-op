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

    Defaults to the first and last categories (the opening and closing balance), which
    matches the common "Anterior ... Total" shape. ``ChartExData.subtotal_idxs``
    overrides that when only some of them are totals.
    """
    n = len(data.values)
    if data.subtotal_idxs is not None:
        idxs = [i for i in data.subtotal_idxs if 0 <= i < n]
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
        fmt = escape(self._data.number_format or "General", {'"': "&quot;"})
        return (
            f'<cx:numDim type="{self._spec.num_dim_type}">'
            f'<cx:lvl ptCount="{n}" formatCode="{fmt}">{pts}</cx:lvl>'
            "</cx:numDim>"
        )

    # -- shared formatting fragments ------------------------------------------

    @property
    def _num_fmt_xml(self) -> str:
        """``cx:numFmt`` for the values, or "" when no number format was given."""
        if not self._data.number_format:
            return ""
        fmt = escape(self._data.number_format, {'"': "&quot;"})
        return f'<cx:numFmt formatCode="{fmt}" sourceLinked="0"/>'

    @property
    def _tx_pr_xml(self) -> str:
        """``cx:txPr`` carrying the tick-label / data-label font, or "" if unset."""
        data = self._data
        if not data.font_name and data.font_size is None:
            return ""
        sz = "" if data.font_size is None else f' sz="{int(round(data.font_size * 100))}"'
        latin = "" if not data.font_name else f'<a:latin typeface="{escape(data.font_name)}"/>'
        return (
            "<cx:txPr><a:bodyPr/><a:p>"
            f"<a:pPr><a:defRPr{sz}>{latin}</a:defRPr></a:pPr>"
            f"<a:endParaRPr{sz}>{latin}</a:endParaRPr>"
            "</a:p></cx:txPr>"
        )

    @staticmethod
    def _lvl_xml(level: list[str]) -> str:
        pts = "".join(f'<cx:pt idx="{i}">{escape(str(v))}</cx:pt>' for i, v in enumerate(level))
        return f'<cx:lvl ptCount="{len(level)}">{pts}</cx:lvl>'

    # -- cx:chart -------------------------------------------------------------

    @property
    def _chart_xml(self) -> str:
        return (
            "<cx:chart><cx:plotArea><cx:plotAreaRegion>"
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
        # -- cx:series child order per the chartex schema: tx, dataPt*, dataLabels,
        # -- dataId, layoutPr. PowerPoint rejects the part if these are out of order.
        return (
            f'<cx:series layoutId="{self._layout_id}" uniqueId="{{00000000-0000-0000-0000-000000000000}}">'
            f"<cx:tx><cx:txData><cx:v>{escape(self._data.series_name)}</cx:v></cx:txData></cx:tx>"
            f"{self._data_pts_xml}"
            f"<cx:dataLabels>{self._num_fmt_xml}{self._tx_pr_xml}"
            '<cx:visibility seriesName="0" categoryName="0" value="1"/></cx:dataLabels>'
            '<cx:dataId val="0"/>'
            f"{layout_pr}"
            "</cx:series>"
        )

    @property
    def _data_pts_xml(self) -> str:
        """One ``cx:dataPt`` per explicitly colored point (uncolored points are skipped)."""
        parts: list[str] = []
        for idx, color in enumerate(self._data.point_colors[: len(self._data.values)]):
            if not color:
                continue
            parts.append(
                f'<cx:dataPt idx="{idx}">'
                f'<cx:spPr><a:solidFill><a:srgbClr val="{escape(color)}"/></a:solidFill></cx:spPr>'
                "</cx:dataPt>"
            )
        return "".join(parts)

    @property
    def _axes_xml(self) -> str:
        parts: list[str] = []
        for axis_id, kind in enumerate(self._spec.axes):
            if kind == "val":
                scaling = "<cx:valScaling/>"
            else:
                scaling = '<cx:catScaling gapWidth="0.5"/>'
            # -- cx:axis child order: scaling, majorGridlines, tickLabels, numFmt, txPr.
            # -- Only the value axis shows numbers, so the format goes there alone.
            num_fmt = self._num_fmt_xml if kind == "val" else ""
            parts.append(
                f'<cx:axis id="{axis_id}">{scaling}'
                "<cx:majorGridlines><cx:spPr><a:ln><a:noFill/></a:ln></cx:spPr></cx:majorGridlines>"
                f"<cx:tickLabels/>{num_fmt}{self._tx_pr_xml}</cx:axis>"
            )
        return "".join(parts)


def new_chartex_xml(chart_type: "XL_CHART_EX_TYPE", chart_data: "ChartExData") -> bytes:
    """Return the chartEx part XML bytes for `chart_type` depicting `chart_data`."""
    return ChartExXmlWriter(chart_type, chart_data).xml_bytes()


def _self_check_parses(chart_type: "XL_CHART_EX_TYPE", chart_data: "ChartExData"):
    """Parse the produced XML to catch malformed output early (used in tests)."""
    return parse_xml(new_chartex_xml(chart_type, chart_data))
