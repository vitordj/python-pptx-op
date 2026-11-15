.. _XlTimeUnit:

``XL_TIME_UNIT``
================

Specifies the unit of time for chart axes and data series.

Used for the base, major, and minor time units of a date (time-scale) category axis.

Example::

    from pptx.enum.chart import XL_TIME_UNIT

    date_axis = chart.category_axis
    date_axis.base_time_unit = XL_TIME_UNIT.MONTHS
    date_axis.major_time_unit = XL_TIME_UNIT.YEARS

----

DAYS
    Days

MONTHS
    Months

YEARS
    Years
