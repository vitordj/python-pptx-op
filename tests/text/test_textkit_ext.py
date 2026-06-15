"""Unit-test suite for the run/paragraph formatting extensions (textkit).

Covers Font.strikethrough / .superscript / .subscript and _Paragraph.bullet. Uses the public
API and asserts against the produced XML so the tests double as documentation of the OOXML
emitted.
"""

import pytest

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.text.text import Font, _Paragraph


def _font():
    return Font(parse_xml(f"<a:rPr {nsdecls('a')}/>"))


def _paragraph():
    return _Paragraph(parse_xml(f"<a:p {nsdecls('a')}/>"), None)


class DescribeFontStrikeAndBaseline:
    def it_defaults_to_None(self):
        f = _font()
        assert f.strikethrough is None
        assert f.superscript is None
        assert f.subscript is None

    @pytest.mark.parametrize(
        "value, expected",
        [(True, "sngStrike"), (False, "noStrike"), ("dblStrike", "dblStrike")],
    )
    def it_can_set_strikethrough(self, value, expected):
        f = _font()
        f.strikethrough = value
        assert f._rPr.get("strike") == expected
        assert f.strikethrough is (expected != "noStrike")

    def it_can_set_superscript_and_subscript_exclusively(self):
        f = _font()
        f.superscript = True
        assert f._rPr.get("baseline") == "30000"
        assert f.superscript is True and f.subscript is False
        f.subscript = True
        assert f._rPr.get("baseline") == "-25000"
        assert f.subscript is True and f.superscript is False
        f.subscript = False
        assert f._rPr.get("baseline") is None


class DescribeParagraphBullet:
    def it_can_suppress_the_bullet(self):
        p = _paragraph()
        p.bullet.none()
        assert p.bullet.type == "none"
        assert p._pPr.find(qn("a:buNone")) is not None

    def it_can_set_a_character_bullet_with_styling(self):
        p = _paragraph()
        p.bullet.character("-", font="Arial", color="FF0000", size_pct=80)
        b = p.bullet
        assert b.type == "character" and b.char == "-"
        pPr = p._pPr
        assert pPr.find(qn("a:buFont")).get("typeface") == "Arial"
        assert pPr.find(qn("a:buSzPct")).get("val") == "80000"
        assert pPr.find(qn("a:buClr") + "/" + qn("a:srgbClr")).get("val") == "FF0000"

    def it_keeps_bullet_children_in_schema_order(self):
        p = _paragraph()
        p.bullet.character("-", font="Arial", color="FF0000", size_pct=80)
        tags = [c.tag for c in p._pPr]
        wanted = [qn(t) for t in ("a:buClr", "a:buSzPct", "a:buFont", "a:buChar")]
        present = [t for t in wanted if t in tags]
        assert [tags.index(t) for t in present] == sorted(tags.index(t) for t in present)

    def it_can_set_an_auto_numbered_bullet(self):
        p = _paragraph()
        p.bullet.auto_number("arabicParenR", start_at=3)
        b = p.bullet
        assert b.type == "auto_number" and b.number_type == "arabicParenR"
        assert p._pPr.find(qn("a:buAutoNum")).get("startAt") == "3"

    def it_makes_the_three_bullet_kinds_mutually_exclusive(self):
        p = _paragraph()
        p.bullet.character("-")
        p.bullet.auto_number("arabicPeriod")
        assert p._pPr.find(qn("a:buChar")) is None
        assert p._pPr.find(qn("a:buAutoNum")) is not None
