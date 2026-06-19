"""DrawingML oxml element classes for visual effects (shadow, soft-edge)."""

from __future__ import annotations

from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import (
    ST_PositiveCoordinate,
    ST_PositiveFixedAngle,
    XsdBoolean,
    XsdString,
)
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    Choice,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
    ZeroOrOneChoice,
)


class CT_EffectList(BaseOxmlElement):
    """`a:effectLst` custom element class.

    Container for shape effects: shadow, glow, soft-edge, etc.
    Only the two most-requested effects are wired here; others remain
    as unregistered (plain lxml) elements.
    """

    _tag_seq = (
        "a:blur",
        "a:fillOverlay",
        "a:glow",
        "a:innerShdw",
        "a:outerShdw",
        "a:prstShdw",
        "a:reflection",
        "a:softEdge",
    )
    outerShdw: CT_OuterShadow | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:outerShdw", successors=_tag_seq[5:]
    )
    softEdge: CT_SoftEdge | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:softEdge", successors=()
    )
    del _tag_seq

    def _new_outerShdw(self):
        """Default outer-shadow matching PowerPoint's 'Offset: Bottom-Right' preset."""
        from pptx.oxml import parse_xml

        return parse_xml(
            "<a:outerShdw %s"
            ' blurRad="40000" dist="23000" dir="5400000"'
            ' algn="ctr" rotWithShape="0">\n'
            '  <a:srgbClr val="000000">\n'
            '    <a:alpha val="63000"/>\n'
            "  </a:srgbClr>\n"
            "</a:outerShdw>\n" % nsdecls("a")
        )


class CT_OuterShadow(BaseOxmlElement):
    """`a:outerShdw` custom element class."""

    eg_colorChoice = ZeroOrOneChoice(
        (
            Choice("a:scrgbClr"),
            Choice("a:srgbClr"),
            Choice("a:hslClr"),
            Choice("a:sysClr"),
            Choice("a:schemeClr"),
            Choice("a:prstClr"),
        ),
        successors=(),
    )
    blurRad: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "blurRad", ST_PositiveCoordinate
    )
    dist: int | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dist", ST_PositiveCoordinate
    )
    dir: float | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "dir", ST_PositiveFixedAngle
    )
    algn: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "algn", XsdString
    )
    rotWithShape: bool | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "rotWithShape", XsdBoolean
    )


class CT_SoftEdge(BaseOxmlElement):
    """`a:softEdge` custom element class."""

    rad: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "rad", ST_PositiveCoordinate
    )
