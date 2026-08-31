"""Support for the 2014 Chart Extension ("chartEx") chart family.

These are the "modern" Office 2016 charts (waterfall, treemap, sunburst, box &
whisker, histogram/Pareto, funnel, region-map). They use the ``cx:`` schema rather
than the legacy ``c:`` chart schema, so they are implemented as a small parallel
module instead of being grafted onto :mod:`pptx.chart`.

See ``docs/ext/charts.rst`` for usage and the list of supported chart types.
"""

from __future__ import annotations

from pptx.chartex.data import ChartExData

__all__ = ["ChartExData"]
