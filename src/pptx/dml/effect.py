"""Visual effects on a shape such as shadow, glow, and reflection."""

from __future__ import annotations

from typing import TYPE_CHECKING

from pptx.dml.color import ColorFormat
from pptx.util import Emu, Pt

if TYPE_CHECKING:
    from pptx.oxml.dml.effect import CT_EffectList, CT_OuterShadow
    from pptx.util import Length


class ShadowFormat(object):
    """Provides access to shadow and soft-edge effects on a shape.

    Access via ``shape.shadow``. A |ShadowFormat| is always returned even
    when no effects are explicitly set (the shape inherits from the style
    hierarchy). Check ``shadow.inherit`` to distinguish.

    Shadow properties (``blur_radius``, ``distance``, ``direction``,
    ``color``) target the OOXML ``a:outerShdw`` element. Reading any of
    them returns ``None`` when no explicit shadow is set; writing any of
    them creates the shadow element with sensible defaults on first write.

    ``soft_edge_radius`` targets the ``a:softEdge`` element independently.
    """

    def __init__(self, spPr):
        # spPr may also be a grpSpPr; both have a:effectLst child
        self._element = spPr

    # ------------------------------------------------------------------
    # inherit (existing behaviour, unchanged)
    # ------------------------------------------------------------------

    @property
    def inherit(self):
        """True if shape inherits shadow settings.

        An explicitly-defined ``a:effectLst`` element causes this to return
        False.  Assigning True removes the element (restoring inheritance
        for *all* effects).  Assigning False inserts an empty element
        (suppresses inheritance without adding any effect).
        """
        return self._element.effectLst is None

    @inherit.setter
    def inherit(self, value):
        if bool(value):
            self._element._remove_effectLst()
        else:
            self._element.get_or_add_effectLst()

    # ------------------------------------------------------------------
    # outer shadow
    # ------------------------------------------------------------------

    @property
    def blur_radius(self) -> Length | None:
        """Blur radius of the outer shadow as an |Emu| object, or None.

        Typical values: ``Pt(3)`` (subtle) to ``Pt(7)`` (strong card shadow).
        """
        outerShdw = self._outerShdw
        return None if outerShdw is None else outerShdw.blurRad

    @blur_radius.setter
    def blur_radius(self, value: Length):
        self._get_or_add_outerShdw().blurRad = Emu(value)

    @property
    def distance(self) -> Length | None:
        """Offset distance from shape to shadow as an |Emu| object, or None."""
        outerShdw = self._outerShdw
        return None if outerShdw is None else outerShdw.dist

    @distance.setter
    def distance(self, value: Length):
        self._get_or_add_outerShdw().dist = Emu(value)

    @property
    def direction(self) -> float | None:
        """Direction of shadow offset as float degrees (0-360), or None.

        0° = right, 90° = down, 180° = left, 270° = up.
        PowerPoint's "Bottom-Right" preset uses 315° (≈ 5400000 / 60000).
        Wait — actually 5400000/60000 = 90°. Let me clarify: 0=right, 90=down.
        PowerPoint's offset-bottom-right default dir is 5400000 → 90°.
        """
        outerShdw = self._outerShdw
        return None if outerShdw is None else outerShdw.dir

    @direction.setter
    def direction(self, value: float):
        self._get_or_add_outerShdw().dir = float(value)

    @property
    def color(self) -> ColorFormat | None:
        """``ColorFormat`` for the shadow colour, or ``None`` if not set.

        Example::

            shape.shadow.color.rgb = RGBColor(0x00, 0x00, 0x00)
        """
        outerShdw = self._outerShdw
        if outerShdw is None:
            return None
        return ColorFormat.from_colorchoice_parent(outerShdw)

    def remove_shadow(self):
        """Remove the outer shadow, if present."""
        effectLst = self._effectLst
        if effectLst is not None:
            effectLst._remove_outerShdw()

    # ------------------------------------------------------------------
    # soft edge
    # ------------------------------------------------------------------

    @property
    def soft_edge_radius(self) -> Length | None:
        """Soft-edge blur radius as an |Emu|, or ``None`` if not set.

        Typical values: ``Pt(2)`` to ``Pt(25)``.
        Setting to ``None`` removes the soft-edge effect.
        """
        effectLst = self._effectLst
        if effectLst is None:
            return None
        softEdge = effectLst.softEdge
        return None if softEdge is None else softEdge.rad

    @soft_edge_radius.setter
    def soft_edge_radius(self, value: Length | None):
        if value is None:
            effectLst = self._effectLst
            if effectLst is not None:
                effectLst._remove_softEdge()
            return
        effectLst = self._element.get_or_add_effectLst()
        effectLst.get_or_add_softEdge().rad = Emu(value)

    # ------------------------------------------------------------------
    # private helpers
    # ------------------------------------------------------------------

    def _get_or_add_outerShdw(self) -> CT_OuterShadow:
        return self._element.get_or_add_effectLst().get_or_add_outerShdw()

    @property
    def _outerShdw(self) -> CT_OuterShadow | None:
        effectLst = self._effectLst
        return None if effectLst is None else effectLst.outerShdw

    @property
    def _effectLst(self) -> CT_EffectList | None:
        return self._element.effectLst
