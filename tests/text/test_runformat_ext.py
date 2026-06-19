"""Unit tests for the run character-format extensions: spacing, caps, highlight."""

import pytest

from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.text.text import Font
from pptx.util import Pt


def _font():
    return Font(parse_xml(f"<a:rPr {nsdecls('a')}/>"))


class DescribeFontSpacing:
    def it_defaults_to_None(self):
        assert _font().spacing is None

    def it_can_set_positive_and_negative_spacing(self):
        f = _font()
        f.spacing = Pt(1.5)
        assert f._rPr.get("spc") == "150"  # 1.5pt = 150 centipoints
        assert f.spacing == Pt(1.5)
        f.spacing = Pt(-0.5)
        assert f._rPr.get("spc") == "-50"
        f.spacing = None
        assert f._rPr.get("spc") is None


class DescribeFontCaps:
    def it_defaults_to_None(self):
        assert _font().caps is None

    @pytest.mark.parametrize("value", ["all", "small", "none"])
    def it_can_set_caps(self, value):
        f = _font()
        f.caps = value
        assert f._rPr.get("cap") == value
        assert f.caps == value

    def it_rejects_invalid_caps(self):
        with pytest.raises(ValueError):
            _font().caps = "BOGUS"


class DescribeFontHighlight:
    def it_defaults_to_None(self):
        assert _font().highlight_color is None

    def it_can_set_and_clear_highlight(self):
        f = _font()
        f.highlight_color = RGBColor(0xFF, 0xFF, 0x00)
        highlight = f._rPr.find(qn("a:highlight"))
        assert highlight.find(qn("a:srgbClr")).get("val") == "FFFF00"
        assert f.highlight_color == RGBColor(0xFF, 0xFF, 0x00)
        f.highlight_color = None
        assert f._rPr.find(qn("a:highlight")) is None

    def it_keeps_rPr_children_in_schema_order(self):
        # highlight must sort before a:latin within a:rPr
        f = _font()
        f.name = "Arial"  # adds a:latin
        f.highlight_color = RGBColor(0x00, 0xFF, 0x00)
        tags = [c.tag for c in f._rPr]
        assert tags.index(qn("a:highlight")) < tags.index(qn("a:latin"))
