"""Unit tests for ColorFormat.transparency (a:alpha), issue #62."""

import pytest

from pptx import Presentation
from pptx.dml.color import ColorFormat, RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Inches

from ..unitutil.cxml import element


def _color_format(cxml="a:solidFill/a:srgbClr{val=ABCDEF}"):
    return ColorFormat.from_colorchoice_parent(element(cxml))


class DescribeColorTransparency:
    def it_defaults_to_zero(self):
        cf = _color_format()
        assert cf.transparency == 0.0

    def it_can_set_transparency_writing_alpha(self):
        cf = _color_format()
        cf.transparency = 0.25
        srgbClr = cf._xFill.find(qn("a:srgbClr"))
        assert srgbClr.find(qn("a:alpha")).get("val") == "75000"
        assert cf.transparency == 0.25

    def it_removes_alpha_when_set_back_to_zero(self):
        cf = _color_format()
        cf.transparency = 0.4
        cf.transparency = 0
        assert cf._xFill.find(qn("a:srgbClr") + "/" + qn("a:alpha")) is None
        assert cf.transparency == 0.0

    def it_works_on_a_theme_color(self):
        cf = _color_format("a:solidFill/a:schemeClr{val=accent1}")
        cf.transparency = 0.5
        assert cf.type == MSO_THEME_COLOR.NOT_THEME_COLOR or cf.theme_color is not None
        assert cf._xFill.find(qn("a:schemeClr") + "/" + qn("a:alpha")).get("val") == "50000"

    @pytest.mark.parametrize("bad", [-0.1, 1.1])
    def it_rejects_out_of_range_values(self, bad):
        cf = _color_format()
        with pytest.raises(ValueError):
            cf.transparency = bad

    def it_rejects_setting_transparency_on_an_undefined_color(self):
        cf = _color_format("a:solidFill")
        with pytest.raises(ValueError):
            cf.transparency = 0.5

    def it_survives_a_round_trip_on_a_real_shape(self):
        from pptx import Presentation
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.util import Inches
        import io

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, *([Inches(1)] * 4))
        shp.fill.solid()
        shp.fill.fore_color.rgb = RGBColor(0x12, 0x34, 0x56)
        shp.fill.fore_color.transparency = 0.3
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        shp2 = prs2.slides[0].shapes[0]
        assert round(shp2.fill.fore_color.transparency, 5) == 0.3


_EXACT_VALUES = [0.1, 0.25, 0.3, 0.5, 0.7, 0.9, 0.99999]


def _rectangle():
    """A rectangle autoshape on a slide of a new presentation."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    return slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, *([Inches(1)] * 4))


class DescribeColorTransparency_noColor:
    def it_reads_zero_when_no_color_is_defined(self):
        cf = _color_format("a:solidFill")
        assert cf.type is None
        assert cf.transparency == 0.0

    def it_reads_zero_for_a_fill_fore_color_with_no_color(self):
        shp = _rectangle()
        shp.fill.solid()
        assert shp.fill.fore_color.type is None
        assert shp.fill.fore_color.transparency == 0.0

    def it_reads_zero_for_a_line_color_with_no_color(self):
        shp = _rectangle()
        shp.line.fill.solid()
        assert shp.line.color.type is None
        assert shp.line.color.transparency == 0.0

    def but_it_still_rejects_setting_transparency_in_that_state(self):
        shp = _rectangle()
        shp.fill.solid()
        with pytest.raises(ValueError):
            shp.fill.fore_color.transparency = 0.5
        shp.line.fill.solid()
        with pytest.raises(ValueError):
            shp.line.color.transparency = 0.5


class DescribeColorTransparency_exactness:
    @pytest.mark.parametrize("value", _EXACT_VALUES)
    def it_reads_back_exactly_what_was_set_on_an_rgb_color(self, value):
        cf = _color_format()
        cf.transparency = value
        assert cf.transparency == value

    @pytest.mark.parametrize("value", _EXACT_VALUES)
    def it_reads_back_exactly_what_was_set_on_a_theme_color(self, value):
        cf = _color_format("a:solidFill/a:schemeClr{val=accent1}")
        cf.transparency = value
        assert cf.transparency == value

    def it_writes_70000_to_the_alpha_val_for_0_3(self):
        cf = _color_format()
        cf.transparency = 0.3
        alpha = cf._xFill.find(qn("a:srgbClr") + "/" + qn("a:alpha"))
        assert alpha.get("val") == "70000"
