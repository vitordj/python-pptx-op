"""ChartEx part object (``ppt/charts/chartEx[N].xml``) and its style/color sidecars."""

from __future__ import annotations

import os
from typing import TYPE_CHECKING

from pptx.chartex.xmlwriter import new_chartex_xml
from pptx.opc.constants import CONTENT_TYPE as CT
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import XmlPart

if TYPE_CHECKING:
    from pptx.chartex.data import ChartExData
    from pptx.enum.chart import XL_CHART_EX_TYPE
    from pptx.package import Package

_TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), "templates")


def _template_bytes(name: str) -> bytes:
    with open(os.path.join(_TEMPLATES_DIR, name), "rb") as f:
        return f.read()


class ChartExStylePart(XmlPart):
    """The ``ppt/charts/style[N].xml`` part required alongside a chartEx chart.

    PowerPoint refuses to open a presentation whose chartEx chart lacks the related
    chartStyle and chartColorStyle parts, so these are created automatically. The style
    is the generic Office default; it is identical for every chartEx chart type.
    """

    partname_template = "/ppt/charts/style%d.xml"

    @classmethod
    def new(cls, package: "Package") -> "ChartExStylePart":
        return cls.load(
            package.next_partname(cls.partname_template),
            CT.OFC_CHART_STYLE,
            package,
            _template_bytes("style.xml"),
        )


class ChartExColorsPart(XmlPart):
    """The ``ppt/charts/colors[N].xml`` (chartColorStyle) sidecar for a chartEx chart."""

    partname_template = "/ppt/charts/colors%d.xml"

    @classmethod
    def new(cls, package: "Package") -> "ChartExColorsPart":
        return cls.load(
            package.next_partname(cls.partname_template),
            CT.OFC_CHART_COLORS,
            package,
            _template_bytes("colors.xml"),
        )


class ChartExPart(XmlPart):
    """A chartEx (2014 Chart Extension) part.

    Corresponds to parts having partnames matching ``ppt/charts/chartEx[1-9][0-9]*.xml``.
    Unlike :class:`~pptx.parts.chart.ChartPart`, this part caches its data inline in the
    chart XML and does not require an embedded Excel workbook to render. It does, however,
    require related chartStyle and chartColorStyle parts (created in :meth:`new`).
    """

    partname_template = "/ppt/charts/chartEx%d.xml"

    @classmethod
    def new(
        cls, chart_type: "XL_CHART_EX_TYPE", chart_data: "ChartExData", package: "Package"
    ) -> "ChartExPart":
        """Return new |ChartExPart| of `chart_type` depicting `chart_data`, added to `package`."""
        chartex_part = cls.load(
            package.next_partname(cls.partname_template),
            CT.OFC_CHART_EX,
            package,
            new_chartex_xml(chart_type, chart_data),
        )
        # -- PowerPoint requires these sidecar parts or it won't open the file --
        chartex_part.relate_to(ChartExStylePart.new(package), RT.CHART_STYLE)
        chartex_part.relate_to(ChartExColorsPart.new(package), RT.CHART_COLOR_STYLE)
        return chartex_part
