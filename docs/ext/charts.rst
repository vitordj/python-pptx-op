.. _ext_charts:

Charts
======

Three additions: the **chartEx** family of modern Office 2016 charts, **doughnut**
geometry, and **date-axis time units**.


chartEx — modern charts
-----------------------

Waterfall, funnel, treemap, sunburst, box & whisker and histogram charts were introduced
in Office 2016 under a separate XML schema — namespace
``.../drawing/2014/chartex``, prefix ``cx:`` — rather than the legacy ``c:`` chart
schema. Because the two schemas share almost nothing, they are implemented here as a
small parallel module, ``pptx.chartex``, instead of being grafted onto ``pptx.chart``.

Adding one
~~~~~~~~~~

::

    from pptx.chartex import ChartExData
    from pptx.enum.chart import XL_CHART_EX_TYPE
    from pptx.util import Cm

    data = ChartExData()
    data.categories = ["Opening", "Q1", "Q2", "Q3", "Closing"]
    data.add_series("Cash flow", (100, 25, -12, 18, 131))

    slide.shapes.add_chartex(
        XL_CHART_EX_TYPE.WATERFALL, Cm(2), Cm(2), Cm(20), Cm(11), data
    )

``add_chartex()`` takes the same position and size arguments as ``add_chart()`` and
returns the |GraphicFrame| holding the chart.

.. important::
   chartEx charts are **write-only**. There is no ``Chart``-style read/write API for
   them: no ``.plots``, no ``.series``, no re-reading a chart from an existing deck.
   You describe the chart and the library emits it. Everything else — the visual style —
   comes from the chart's own theme parts, as it does in PowerPoint.

Chart types
~~~~~~~~~~~

``XL_CHART_EX_TYPE`` members map onto the ``layoutId`` attribute of the ``cx:series``
element, which is how the schema distinguishes one chart type from another.

.. csv-table::
   :header: "Member", "Chart", "Categories", "Emitted"
   :widths: 25, 30, 22, 13

   "``WATERFALL``", "Waterfall (bridge)", "flat", "yes"
   "``FUNNEL``", "Funnel", "flat", "yes"
   "``TREEMAP``", "Treemap", "hierarchical", "yes"
   "``SUNBURST``", "Sunburst", "hierarchical", "yes"
   "``BOX_WHISKER``", "Box & whisker", "flat", "yes"
   "``CLUSTERED_COLUMN``", "Clustered column (histogram layout)", "flat", "yes"
   "``PARETO_LINE``", "Pareto line", "flat", "not yet"
   "``REGION_MAP``", "Filled region map", "flat", "not yet"

The last two are defined in the enumeration because they are part of the schema, but the
XML writer does not emit them yet; asking for one raises ``NotImplementedError``.

Data
~~~~

``ChartExData`` holds a category dimension and exactly **one** series — which is what
this chart family models. Calling ``add_series()`` twice replaces the first definition
rather than adding a second series.

Flat categories are a list of strings::

    data = ChartExData()
    data.categories = ["Opening", "Q1", "Q2", "Closing"]
    data.add_series("Cash flow", (100, 25, -12, 113))

Treemap and sunburst need **hierarchical** categories instead: one tuple per data point,
ordered root-first, all of the same depth::

    data = ChartExData()
    data.categories = [
        ("Europe", "France", "Paris"),
        ("Europe", "France", "Lyon"),
        ("Europe", "Spain", "Madrid"),
        ("Americas", "Brazil", "São Paulo"),
    ]
    data.add_series("Revenue", (22, 12, 18, 30))

    slide.shapes.add_chartex(
        XL_CHART_EX_TYPE.TREEMAP, Cm(2), Cm(2), Cm(14), Cm(10), data
    )

Each tuple is a full path from the root to a leaf; the intermediate levels are what the
chart groups by. A flat list of strings works for a treemap too — it just produces a
single, ungrouped level.

The data is checked when the chart is built: ``add_chartex()`` raises ``ValueError``,
and adds nothing to the slide, when the category and value counts differ, a value is
``nan`` or infinite, hierarchical categories have unequal depths, or one of the
formatting options below does not fit the data.

Waterfall totals
~~~~~~~~~~~~~~~~

In a waterfall, a *total* bar starts at zero instead of floating on the running sum. By
default the first and last categories are totals, which fits the common "opening …
closing" shape. ``subtotal_idxs`` marks exactly the points you list instead::

    data.categories = ["Q1", "Q2", "Q3", "Q4", "Year"]
    data.add_series("Result", (12, -4, 7, 3, 18))
    data.subtotal_idxs = [4]        # only "Year" is a total
    data.subtotal_idxs = []         # no totals at all
    data.subtotal_idxs = None       # back to the default (first and last)

Formatting
~~~~~~~~~~

``ChartExData`` also carries the formatting the chart is written with. Every option is
``None`` (or empty) by default, which leaves the chart style's own look::

    data.point_colors = ["8CC640", "8CC640", "FF0000", None, "1F4E79"]
    data.number_format = "0.00%"
    data.font_name = "Montserrat"
    data.font_size = 9

.. csv-table::
   :header: "Property", "Applies to", "Notes"
   :widths: 18, 32, 50

   "``point_colors``", "Fill of each bar, slice or box", "One colour per point, in order: an ``RGBColor`` or a hex string such as ``92D050`` or ``#92D050``. ``None`` entries and points past the end of the list keep the theme colour. More colours than values is an error"
   "``number_format``", "Data labels, value-axis tick labels and the cached values", "Any Excel format code, e.g. ``0.0%`` or ``#,##0.00``"
   "``font_name``", "Data labels and tick labels of every axis", "Latin typeface"
   "``font_size``", "Data labels and tick labels of every axis", "Points, 1 to 4000"

A misspelled property name raises ``AttributeError`` rather than being silently ignored.

A typical green-up / red-down waterfall::

    values = (0.10, 0.03, -0.02, 0.04, 0.15)
    data = ChartExData()
    data.categories = ["Previous", "Equities", "Credit", "FX", "Total"]
    data.add_series("Attribution", values)
    data.point_colors = [
        "7F7F7F" if i in (0, len(values) - 1) else ("8CC640" if v >= 0 else "E5413B")
        for i, v in enumerate(values)
    ]
    data.number_format = "0.00%"


Doughnut geometry
-----------------

`Upstream issue #493 <https://github.com/scanny/python-pptx/issues/493>`_.

::

    plot = chart.plots[0]
    plot.hole_size = 60          # percent of the diameter, 1 to 90
    plot.first_slice_angle = 90  # degrees clockwise from 12 o'clock, 0 to 360

``hole_size`` is the diameter of the hole as a percentage of the chart diameter; the
PowerPoint default is 10. ``first_slice_angle`` rotates the whole chart, which is how you
control where the first category starts. Both are read/write and validated against their
schema ranges.


Date-axis time units
--------------------

A date (time-scale) category axis groups its data points into units of time. Three
properties expose that grouping, each taking a member of ``XL_TIME_UNIT`` — ``DAYS``,
``MONTHS`` or ``YEARS``::

    from pptx.enum.chart import XL_TIME_UNIT

    date_axis = chart.category_axis
    date_axis.base_time_unit = XL_TIME_UNIT.MONTHS   # group daily points into months
    date_axis.major_time_unit = XL_TIME_UNIT.YEARS   # major tick every year
    date_axis.minor_time_unit = XL_TIME_UNIT.MONTHS

``base_time_unit`` sets the granularity the axis groups by. ``major_time_unit`` and
``minor_time_unit`` combine with the existing ``major_unit`` and ``minor_unit`` counts to
place tick marks — ``major_unit = 3`` with ``major_time_unit = MONTHS`` puts a major tick
every quarter.

All three return ``None`` when unset, which is the 'Auto' setting in the PowerPoint UI;
assigning ``None`` removes the element and restores Auto.

.. note::
   These properties live on |DateAxis|. An axis is only a date axis when the chart's
   category data is dates; on a plain category axis the properties are not present.
