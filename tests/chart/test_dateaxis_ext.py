"""Unit tests for the DateAxis base/major/minor unit properties."""

import datetime
import io
import itertools

import pytest

from pptx import Presentation
from pptx.chart.axis import DateAxis
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_TIME_UNIT
from pptx.oxml.ns import qn
from pptx.util import Inches

from ..unitutil.cxml import element

SCHEMA_ORDER = [
    "lblOffset",
    "baseTimeUnit",
    "majorUnit",
    "majorTimeUnit",
    "minorUnit",
    "minorTimeUnit",
    "extLst",
]

# (property name, element tag, is_time_unit)
PROPS = [
    ("base_time_unit", "c:baseTimeUnit", True),
    ("major_unit", "c:majorUnit", False),
    ("major_time_unit", "c:majorTimeUnit", True),
    ("minor_unit", "c:minorUnit", False),
    ("minor_time_unit", "c:minorTimeUnit", True),
]
TIME_PROPS = [p for p in PROPS if p[2]]
UNIT_PROPS = [p for p in PROPS if not p[2]]


def _child_names(dateAx):
    return [child.tag.split("}")[1] for child in dateAx]


def _axis(cxml="c:dateAx"):
    return DateAxis(element(cxml))


class DescribeDateAxisUnits:
    @pytest.mark.parametrize("prop", [p[0] for p in PROPS])
    def it_returns_None_when_the_element_is_absent(self, prop):
        assert getattr(_axis(), prop) is None

    @pytest.mark.parametrize(("prop", "tag"), [(p[0], p[1]) for p in TIME_PROPS])
    @pytest.mark.parametrize("member", list(XL_TIME_UNIT))
    def it_reads_each_time_unit(self, prop, tag, member):
        axis = _axis("c:dateAx/%s{val=%s}" % (tag, member.xml_value))
        assert getattr(axis, prop) == member

    @pytest.mark.parametrize(("prop", "tag"), [(p[0], p[1]) for p in TIME_PROPS])
    def it_reads_DAYS_for_a_time_unit_element_without_val(self, prop, tag):
        axis = _axis("c:dateAx/%s" % tag)
        assert getattr(axis, prop) == XL_TIME_UNIT.DAYS

    @pytest.mark.parametrize(("prop", "tag"), [(p[0], p[1]) for p in UNIT_PROPS])
    def it_reads_major_and_minor_unit_as_float(self, prop, tag):
        value = getattr(_axis("c:dateAx/%s{val=3}" % tag), prop)
        assert value == 3.0
        assert isinstance(value, float)

    @pytest.mark.parametrize(("prop", "tag"), [(p[0], p[1]) for p in TIME_PROPS])
    @pytest.mark.parametrize("member", list(XL_TIME_UNIT))
    def it_writes_each_time_unit(self, prop, tag, member):
        axis = _axis()
        setattr(axis, prop, member)
        elms = axis._element.findall(qn(tag))
        assert len(elms) == 1
        assert getattr(axis, prop) == member
        if member == XL_TIME_UNIT.DAYS:
            # -- default value is not written; the schema default is "days" --
            assert elms[0].get("val") is None
        else:
            assert elms[0].get("val") == member.xml_value

    def it_reads_back_DAYS_when_assigned_DAYS(self):
        axis = _axis()
        axis.base_time_unit = XL_TIME_UNIT.DAYS
        # -- OptionalAttribute with default DAYS omits `val`, so `<c:baseTimeUnit/>` --
        assert axis._element.baseTimeUnit is not None
        assert axis._element.baseTimeUnit.attrib == {}
        assert axis.base_time_unit == XL_TIME_UNIT.DAYS

    @pytest.mark.parametrize(("prop", "tag"), [(p[0], p[1]) for p in UNIT_PROPS])
    def it_writes_major_and_minor_unit(self, prop, tag):
        axis = _axis()
        setattr(axis, prop, 3)
        assert axis._element.find(qn(tag)).get("val") == "3.0"
        assert getattr(axis, prop) == 3.0
        setattr(axis, prop, 2.5)
        assert getattr(axis, prop) == 2.5

    @pytest.mark.parametrize(("prop", "tag", "is_time"), PROPS)
    def it_removes_the_element_when_assigned_None(self, prop, tag, is_time):
        value = XL_TIME_UNIT.YEARS if is_time else 4
        axis = _axis()
        setattr(axis, prop, value)
        assert axis._element.find(qn(tag)) is not None
        setattr(axis, prop, None)
        assert axis._element.find(qn(tag)) is None
        assert getattr(axis, prop) is None

    def it_tolerates_assigning_None_when_already_absent(self):
        axis = _axis()
        for prop, _, _ in PROPS:
            setattr(axis, prop, None)
        assert _child_names(axis._element) == []

    @pytest.mark.parametrize(("prop", "tag", "is_time"), PROPS)
    def it_replaces_rather_than_duplicates_on_reassignment(self, prop, tag, is_time):
        first, second = (XL_TIME_UNIT.MONTHS, XL_TIME_UNIT.YEARS) if is_time else (1, 6)
        axis = _axis()
        setattr(axis, prop, first)
        setattr(axis, prop, second)
        assert len(axis._element.findall(qn(tag))) == 1
        assert getattr(axis, prop) == second

    def it_leaves_the_other_properties_alone(self):
        axis = _axis()
        axis.major_unit = 2
        axis.major_time_unit = XL_TIME_UNIT.MONTHS
        axis.base_time_unit = XL_TIME_UNIT.YEARS
        axis.major_time_unit = None
        assert axis.major_unit == 2.0
        assert axis.base_time_unit == XL_TIME_UNIT.YEARS
        assert axis.minor_unit is None
        assert axis.minor_time_unit is None


class DescribeDateAxisChildOrder:
    _VALUES = {
        "base_time_unit": XL_TIME_UNIT.MONTHS,
        "major_unit": 3,
        "major_time_unit": XL_TIME_UNIT.YEARS,
        "minor_unit": 1,
        "minor_time_unit": XL_TIME_UNIT.DAYS,
    }
    _NAMES = list(_VALUES)

    @pytest.mark.parametrize(
        "order",
        [
            _NAMES,
            list(reversed(_NAMES)),
            ["major_time_unit", "minor_unit", "base_time_unit", "minor_time_unit", "major_unit"],
            ["minor_time_unit", "base_time_unit", "major_unit", "minor_unit", "major_time_unit"],
        ],
    )
    def it_keeps_schema_order_whatever_the_assignment_order(self, order):
        axis = _axis("c:dateAx/(c:lblOffset{val=100},c:extLst)")
        for prop in order:
            setattr(axis, prop, self._VALUES[prop])
        assert _child_names(axis._element) == SCHEMA_ORDER

    def it_keeps_schema_order_for_every_permutation(self):
        for order in itertools.permutations(self._NAMES):
            axis = _axis("c:dateAx/(c:lblOffset{val=100},c:extLst)")
            for prop in order:
                setattr(axis, prop, self._VALUES[prop])
            assert _child_names(axis._element) == SCHEMA_ORDER, order

    def it_keeps_schema_order_when_replacing_and_removing(self):
        axis = _axis("c:dateAx/(c:lblOffset{val=100},c:extLst)")
        for prop in reversed(self._NAMES):
            setattr(axis, prop, self._VALUES[prop])
        axis.major_unit = None
        axis.major_unit = 7
        axis.base_time_unit = XL_TIME_UNIT.YEARS
        assert _child_names(axis._element) == SCHEMA_ORDER


class DescribeDateAxisRoundTrip:
    def it_round_trips_units_through_a_saved_presentation(self):
        data = CategoryChartData()
        data.categories = [datetime.date(2024, 1, 1), datetime.date(2024, 2, 1)]
        data.add_series("S", (1, 2))
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.LINE, Inches(1), Inches(1), Inches(5), Inches(4), data
        ).chart
        axis = chart.category_axis
        assert isinstance(axis, DateAxis)

        axis.base_time_unit = XL_TIME_UNIT.MONTHS
        axis.major_unit = 3
        axis.major_time_unit = XL_TIME_UNIT.MONTHS
        axis.minor_unit = 1
        axis.minor_time_unit = XL_TIME_UNIT.YEARS

        stream = io.BytesIO()
        prs.save(stream)
        stream.seek(0)
        reloaded = Presentation(stream).slides[0].shapes[-1].chart.category_axis

        assert isinstance(reloaded, DateAxis)
        assert reloaded.base_time_unit == XL_TIME_UNIT.MONTHS
        assert reloaded.major_unit == 3.0
        assert reloaded.major_time_unit == XL_TIME_UNIT.MONTHS
        assert reloaded.minor_unit == 1.0
        assert reloaded.minor_time_unit == XL_TIME_UNIT.YEARS
        names = _child_names(reloaded._element)
        tail = names[names.index("lblOffset") :]
        assert tail == [n for n in SCHEMA_ORDER if n in tail]
        assert tail[-5:] == SCHEMA_ORDER[1:6]
