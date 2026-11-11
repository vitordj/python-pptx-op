"""Tests for _Paragraph.margin_left/margin_right/first_line_indent (paraindent branch)."""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.text.text import _Paragraph
from pptx.util import Emu, Inches, Pt

from ..unitutil.cxml import element


def _para(cxml="a:p"):
    from pptx.text.text import _Paragraph

    p = element(cxml)
    return _Paragraph(p, None)


class DescribeParagraph_margin_left:
    def it_returns_None_when_no_pPr(self):
        assert _para().margin_left is None

    def it_returns_None_when_pPr_has_no_marL(self):
        assert _para("a:p/a:pPr").margin_left is None

    def it_can_set_and_read_margin_left(self):
        para = _para()
        para.margin_left = Pt(36)
        assert para.margin_left == Pt(36)

    def it_can_clear_margin_left_by_assigning_None(self):
        para = _para()
        para.margin_left = Pt(18)
        para.margin_left = None
        assert para.margin_left is None
        assert para._p.pPr.get("marL") is None


class DescribeParagraph_margin_right:
    def it_returns_None_when_no_pPr(self):
        assert _para().margin_right is None

    def it_can_set_and_read_margin_right(self):
        para = _para()
        para.margin_right = Pt(12)
        assert para.margin_right == Pt(12)

    def it_can_clear_margin_right_by_assigning_None(self):
        para = _para()
        para.margin_right = Pt(12)
        para.margin_right = None
        assert para.margin_right is None


class DescribeParagraph_first_line_indent:
    def it_returns_None_when_no_pPr(self):
        assert _para().first_line_indent is None

    def it_can_set_a_positive_first_line_indent(self):
        para = _para()
        para.first_line_indent = Pt(18)
        assert para.first_line_indent == Pt(18)

    def it_can_set_a_negative_indent_for_hanging_effect(self):
        para = _para()
        para.first_line_indent = Pt(-18)
        assert para.first_line_indent == Pt(-18)

    def it_can_clear_first_line_indent_by_assigning_None(self):
        para = _para()
        para.first_line_indent = Pt(18)
        para.first_line_indent = None
        assert para.first_line_indent is None

    def it_survives_a_round_trip_with_all_three_properties(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shp = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
        tf = shp.text_frame
        tf.text = "indented paragraph"
        para = tf.paragraphs[0]
        para.margin_left = Pt(36)
        para.margin_right = Pt(12)
        para.first_line_indent = Pt(-18)

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        para2 = prs2.slides[0].shapes[0].text_frame.paragraphs[0]

        assert para2.margin_left == Pt(36)
        assert para2.margin_right == Pt(12)
        assert para2.first_line_indent == Pt(-18)


class DescribeParagraph_indent_xml_units:
    """The XML attributes are plain integer EMU, not centipoints."""

    def it_writes_EMU_to_the_XML_attributes(self):
        para = _para()
        para.margin_left = Pt(36)
        para.margin_right = Pt(12)
        para.first_line_indent = Pt(-18)
        pPr = para._p.pPr
        assert pPr.get("marL") == "457200"
        assert pPr.get("marR") == "152400"
        assert pPr.get("indent") == "-228600"

    def it_reads_PowerPoint_authored_EMU_values(self):
        p = parse_xml(
            '<a:p %s><a:pPr marL="342900" indent="-342900"/></a:p>' % nsdecls("a")
        )
        para = _Paragraph(p, None)
        assert para.margin_left == Inches(0.375)
        assert para.first_line_indent == -Inches(0.375)
        assert para.margin_left.pt == pytest.approx(27.0)
        assert isinstance(para.margin_left, Emu)

    def it_agrees_with_the_units_PowerPoint_wrote_in_the_default_template(self):
        # -- ground truth for the unit: the slide master in the default template was authored
        # -- by PowerPoint and indents level-1 body text by 0.375", which is 342900 EMU.
        master = Presentation().slide_masters[0]
        lvl1 = master._element.xpath("p:txStyles/p:bodyStyle/a:lvl1pPr")[0]
        assert lvl1.get("marL") == str(Inches(0.375)) == "342900"
        assert lvl1.get("indent") == "-342900"

    @pytest.mark.parametrize("prop", ["margin_left", "margin_right"])
    def it_rejects_negative_margins(self, prop):
        para = _para()
        with pytest.raises(ValueError):
            setattr(para, prop, Pt(-1))

    def it_accepts_a_negative_first_line_indent(self):
        para = _para()
        para.first_line_indent = Emu(-51206400)
        assert para.first_line_indent == Emu(-51206400)

    @pytest.mark.parametrize("prop", ["margin_left", "margin_right", "first_line_indent"])
    def it_rejects_values_above_the_maximum(self, prop):
        para = _para()
        with pytest.raises(ValueError):
            setattr(para, prop, Emu(51206401))

    def it_accepts_the_maximum_value(self):
        para = _para()
        para.margin_left = Emu(51206400)
        assert para._p.pPr.get("marL") == "51206400"

    def it_rejects_a_first_line_indent_below_the_minimum(self):
        para = _para()
        with pytest.raises(ValueError):
            para.first_line_indent = Emu(-51206401)

    def it_removes_the_attributes_when_assigned_None(self):
        para = _para()
        para.margin_left = Pt(36)
        para.margin_right = Pt(12)
        para.first_line_indent = Pt(-18)
        para.margin_left = None
        para.margin_right = None
        para.first_line_indent = None
        pPr = para._p.pPr
        assert pPr.get("marL") is None
        assert pPr.get("marR") is None
        assert pPr.get("indent") is None
