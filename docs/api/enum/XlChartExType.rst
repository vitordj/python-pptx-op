.. _XL_CHART_EX_TYPE:

``XL_CHART_EX_TYPE``
====================

Specifies a chart type from the 2014 Chart Extension ("chartEx") family — the "modern"
charts introduced with Office 2016.

These use a separate XML schema (namespace ``...drawing/2014/chartex``, prefix ``cx:``)
rather than the legacy ``c:`` chart schema. A single chart type is distinguished by the
``layoutId`` attribute of its ``cx:series`` element, and the XML value of each member
below is exactly that ``layoutId``.

See :ref:`ext_charts` for usage.

Example::

    from pptx.enum.chart import XL_CHART_EX_TYPE

    slide.shapes.add_chartex(XL_CHART_EX_TYPE.WATERFALL, x, y, cx, cy, data)

----

WATERFALL
    Waterfall (bridge) chart.

FUNNEL
    Funnel chart.

TREEMAP
    Treemap (hierarchical area) chart. Requires hierarchical categories.

SUNBURST
    Sunburst (hierarchical ring) chart. Requires hierarchical categories.

BOX_WHISKER
    Box & Whisker (statistical) chart.

CLUSTERED_COLUMN
    Clustered-column chart, the layout used by Histogram and Pareto.

PARETO_LINE
    Pareto line, overlaid on a histogram. Defined by the schema but not yet emitted by
    the XML writer; requesting it raises ``NotImplementedError``.

REGION_MAP
    Filled region (geographic) map. Defined by the schema but not yet emitted by the XML
    writer; requesting it raises ``NotImplementedError``.
