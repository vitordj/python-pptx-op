"""Unit-test suite for the run/paragraph formatting extensions (textkit).

Covers Font.strikethrough / .superscript / .subscript and _Paragraph.bullet. Uses the public
API and asserts against the produced XML so the tests double as documentation of the OOXML
emitted.
"""

import pytest
from lxml import etree

from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.text.text import Font, _Bullet, _Paragraph


def _font():
    return Font(parse_xml(f"<a:rPr {nsdecls('a')}/>"))


def _pPr(inner=""):
    return parse_xml(f"<a:pPr {nsdecls('a')}>{inner}</a:pPr>")


def _local_tags(pPr):
    return [child.tag.split("}")[1] for child in pPr]


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

    @pytest.mark.parametrize("value", [True, False, None, "sngStrike", "dblStrike", "noStrike"])
    def it_accepts_every_valid_strikethrough_value(self, value):
        f = _font()
        f.strikethrough = value
        expected = {True: "sngStrike", False: "noStrike", None: None}.get(value, value)
        assert f._rPr.get("strike") == expected

    @pytest.mark.parametrize("value", ["bogus", "single", "", 1])
    def it_rejects_an_invalid_strikethrough_value(self, value):
        f = _font()
        with pytest.raises(ValueError):
            f.strikethrough = value
        assert f._rPr.get("strike") is None

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


class DescribeBulletChoiceGroups:
    @pytest.mark.parametrize(
        ("existing", "kwargs", "expected"),
        [
            ("<a:buClrTx/>", {"color": "FF0000"}, ["buClr", "buChar"]),
            ("<a:buSzTx/>", {"size_pct": 80}, ["buSzPct", "buChar"]),
            ("<a:buSzPts val='1200'/>", {"size_pct": 80}, ["buSzPct", "buChar"]),
            ("<a:buFontTx/>", {"font": "Arial"}, ["buFont", "buChar"]),
            ("<a:buBlip/>", {}, ["buChar"]),
            ("<a:buNone/>", {}, ["buChar"]),
            ("<a:buAutoNum type='arabicPeriod'/>", {}, ["buChar"]),
            (
                "<a:buClrTx/><a:buSzTx/><a:buFontTx/><a:buBlip/>",
                {"color": "FF0000", "size_pct": 80, "font": "Arial"},
                ["buClr", "buSzPct", "buFont", "buChar"],
            ),
        ],
    )
    def it_keeps_each_choice_group_exclusive_and_ordered(self, existing, kwargs, expected):
        pPr = _pPr(existing)
        _Bullet(pPr).character("-", **kwargs)
        assert _local_tags(pPr) == expected

    def it_replaces_buClr_buSzPct_and_buFont_with_their_alternates_too(self):
        pPr = _pPr("<a:buClr><a:srgbClr val='00FF00'/></a:buClr><a:buSzPct val='50000'/>")
        _Bullet(pPr).character("-", color="FF0000", size_pct=80)
        assert _local_tags(pPr) == ["buClr", "buSzPct", "buChar"]

    def it_handles_the_buFontTx_buBlip_repro(self):
        pPr = _pPr("<a:buFontTx/><a:buBlip/>")
        _Bullet(pPr).character("•", font="Arial")
        assert _local_tags(pPr) == ["buFont", "buChar"]

    @pytest.mark.parametrize("method", ["none", "character", "auto_number"])
    def it_removes_a_pre_existing_buBlip(self, method):
        pPr = _pPr("<a:buBlip/>")
        getattr(_Bullet(pPr), method)()
        assert pPr.find(qn("a:buBlip")) is None
        assert len(_local_tags(pPr)) == 1

    def it_reports_a_picture_bullet_type(self):
        assert _Bullet(_pPr("<a:buBlip/>")).type == "picture"

    def it_does_not_duplicate_elements_when_character_is_called_twice(self):
        pPr = _pPr()
        b = _Bullet(pPr)
        b.character("-", font="Arial", color="FF0000", size_pct=80)
        b.character("-", font="Arial", color="FF0000", size_pct=80)
        assert _local_tags(pPr) == ["buClr", "buSzPct", "buFont", "buChar"]


class DescribeBulletValidation:
    @pytest.mark.parametrize("color", ["FF0000", "#ff0000", RGBColor(0xFF, 0, 0)])
    def it_normalises_valid_colors(self, color):
        pPr = _pPr()
        _Bullet(pPr).character("-", color=color)
        assert pPr.find(qn("a:buClr") + "/" + qn("a:srgbClr")).get("val") == "FF0000"

    @pytest.mark.parametrize("color", ["FF00", "GGGGGG", "#FF0000AA", "", "##FF0000"])
    @pytest.mark.parametrize("method", ["character", "auto_number"])
    def it_rejects_invalid_colors_without_mutating(self, color, method):
        pPr = _pPr("<a:buFontTx/><a:buBlip/>")
        before = etree.tostring(pPr)
        with pytest.raises(ValueError):
            getattr(_Bullet(pPr), method)(font="Arial", color=color, size_pct=80)
        assert etree.tostring(pPr) == before

    @pytest.mark.parametrize("size", [25, 400])
    def it_accepts_boundary_sizes(self, size):
        pPr = _pPr()
        _Bullet(pPr).character("-", size_pct=size)
        assert pPr.find(qn("a:buSzPct")).get("val") == str(size * 1000)

    @pytest.mark.parametrize("size", [24, 401, 0, -5])
    @pytest.mark.parametrize("method", ["character", "auto_number"])
    def it_rejects_out_of_range_sizes_without_mutating(self, size, method):
        pPr = _pPr("<a:buSzTx/><a:buBlip/>")
        before = etree.tostring(pPr)
        with pytest.raises(ValueError):
            getattr(_Bullet(pPr), method)(font="Arial", color="FF0000", size_pct=size)
        assert etree.tostring(pPr) == before
