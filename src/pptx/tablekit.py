"""Batteries-included helpers for building richly-formatted tables.

The core table API (``slide.shapes.add_table`` / |Table| / |_Cell|) exposes the raw
primitives -- per-cell ``text``, ``fill``, ``borders`` (issue #71), ``merge`` and
margins -- but driving them cell-by-cell for a real report is verbose and easy to get
wrong. ``tablekit`` layers reusable *cell styles* and *range helpers* on top of those
primitives so an application can recreate a complex, fully-formatted table natively,
instead of round-tripping through Excel + a COM paste.

Quick start::

    from pptx.tablekit import CellStyle, BorderSpec, Edge, StyledTable

    HEADER = CellStyle(font_name="Montserrat", size=9, bold=True, color="FFFFFF",
                       fill="62993D", align="center", anchor="middle",
                       borders=BorderSpec.box(Edge("FFFFFF", 0.5)))
    BODY = CellStyle(font_name="Montserrat", size=9, color="000000", fill="FFFFFF",
                     align="center", anchor="middle",
                     borders=BorderSpec(bottom=Edge("D3D3D3", 0.5)))

    st = StyledTable.add(slide.shapes, rows=7, cols=5,
                         x=Cm(13.33), y=Cm(3.73), cx=Cm(19), cy=Cm(4.6))
    st.fill_rows(data, header_style=HEADER, body_style=BODY)   # data: list[list[str]]
    st.fit(width_cm=19)                                        # exact total width
    st.autosize_rows(font_pt=9)                                # row height = rendered height

Why this exists (the limitations it papers over):

* python-pptx has no high-level border or "apply table style" support
  (issues #71, #27, #203, #573) -- ``borders`` is the primitive, ``CellStyle`` /
  ``BorderSpec`` make it reusable.
* ``add_table`` injects a default *table style* (``firstRow``/``bandRow`` + a
  ``tableStyleId``) whose banding fights explicit per-cell fills. ``StyledTable``
  strips it so explicit fills win (matching what a PowerPoint paste produces).
* A table row height in the XML is only a *minimum*; PowerPoint grows it to fit the
  text at render time but never writes the grown value back (issues #480, #296). So a
  table built and saved purely from python looks too short until opened+saved in real
  PowerPoint. ``autosize_rows`` sets the stored height to the height PowerPoint *would*
  render (~1.2x the font size per text line + margins), removing the need for that
  round-trip.
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from typing import TYPE_CHECKING, Iterable, Sequence

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_VERTICAL_ANCHOR, PP_ALIGN
from pptx.util import Cm, Emu, Pt

if TYPE_CHECKING:
    from pptx.enum.dml import MSO_LINE_DASH_STYLE
    from pptx.shapes.shapetree import _BaseGroupShapes  # noqa: F401
    from pptx.table import Table, _Cell
    from pptx.util import Length

Color = "str | RGBColor"  # "62993D" hex string or an RGBColor

_ALIGN = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "centre": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
    "justify": PP_ALIGN.JUSTIFY,
}
_ANCHOR = {
    "top": MSO_VERTICAL_ANCHOR.TOP,
    "middle": MSO_VERTICAL_ANCHOR.MIDDLE,
    "center": MSO_VERTICAL_ANCHOR.MIDDLE,
    "bottom": MSO_VERTICAL_ANCHOR.BOTTOM,
}


def _rgb(color) -> RGBColor:
    """Coerce a hex string (``"62993D"``) or |RGBColor| to |RGBColor|."""
    if isinstance(color, RGBColor):
        return color
    return RGBColor.from_string(str(color).lstrip("#"))


# ---------------------------------------------------------------------------
# style value objects (immutable, composable)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Edge:
    """One border edge: a color, a width in points, and an optional dash style."""

    color: "str | RGBColor" = "000000"
    pt: float = 0.5
    dash: "MSO_LINE_DASH_STYLE | None" = None


class _NoLine:
    """Sentinel: draw *explicitly no line* on an edge (emits ``<a:lnX><a:noFill/></a:lnX>``).

    This is different from leaving an edge as ``None``. PowerPoint draws its *default*
    (dark) cell border on an edge that has no explicit ``a:lnX`` element; setting the
    edge to :data:`NO_LINE` suppresses that default. This is exactly what a PowerPoint
    clipboard paste does on the left/right edges of body cells -- without it you get
    stray vertical black lines between columns.
    """

    __slots__ = ()

    def __repr__(self):  # pragma: no cover - cosmetic
        return "NO_LINE"


NO_LINE = _NoLine()


@dataclass(frozen=True)
class BorderSpec:
    """The (up to four) edges of a cell.

    Each edge is one of:

    * ``None``         -- leave the edge alone (PowerPoint may draw its default border);
    * an :class:`Edge` -- draw that line;
    * :data:`NO_LINE`  -- *explicitly* draw nothing (suppresses the default border).
    """

    left: "Edge | _NoLine | None" = None
    right: "Edge | _NoLine | None" = None
    top: "Edge | _NoLine | None" = None
    bottom: "Edge | _NoLine | None" = None

    @classmethod
    def box(cls, edge: Edge) -> "BorderSpec":
        """All four edges drawn with the same `edge`."""
        return cls(left=edge, right=edge, top=edge, bottom=edge)

    @classmethod
    def underline(cls, edge: Edge, suppress_sides: bool = True) -> "BorderSpec":
        """A bottom "row separator" line, with the side edges suppressed by default.

        Suppressing the sides (``NO_LINE``) is what keeps PowerPoint from drawing stray
        vertical borders between columns -- the common gotcha when only a bottom line is
        set.
        """
        side = NO_LINE if suppress_sides else None
        return cls(left=side, right=side, top=None, bottom=edge)

    @classmethod
    def hlines(cls, edge: Edge, suppress_sides: bool = True) -> "BorderSpec":
        """Horizontal lines top *and* bottom, side edges suppressed by default."""
        side = NO_LINE if suppress_sides else None
        return cls(left=side, right=side, top=edge, bottom=edge)

    @classmethod
    def topline(cls, edge: Edge, suppress_sides: bool = True) -> "BorderSpec":
        """A single top line (e.g. a section/subtotal separator), sides suppressed."""
        side = NO_LINE if suppress_sides else None
        return cls(left=side, right=side, top=edge, bottom=NO_LINE if suppress_sides else None)


@dataclass(frozen=True)
class CellStyle:
    """A reusable bundle of cell formatting.

    Every field is optional; ``None`` means "don't touch this aspect", so styles
    *compose*: apply a base style, then a :meth:`derive` of it that overrides only the
    fill, and the rest is preserved. This replaces ad-hoc "conditional format" reuse
    (e.g. one xlsxwriter format per row type) with explicit, named, derivable styles.
    """

    font_name: "str | None" = None
    size: "float | None" = None
    bold: "bool | None" = None
    italic: "bool | None" = None
    color: "str | RGBColor | None" = None
    fill: "str | RGBColor | None" = None  # hex/RGBColor; the literal "none" -> no fill
    align: "str | None" = None  # left|center|right|justify
    anchor: "str | None" = None  # top|middle|bottom
    borders: "BorderSpec | None" = None
    # cell text margins in cm, as (left, right, top, bottom); None entries keep default
    margins_cm: "tuple[float | None, float | None, float | None, float | None] | None" = None

    def derive(self, **overrides) -> "CellStyle":
        """Return a copy of this style with the given fields overridden."""
        return replace(self, **overrides)


# ---------------------------------------------------------------------------
# applying styles to live cells
# ---------------------------------------------------------------------------


def _set_edge(line_format, edge) -> None:
    """Apply one edge spec to a |LineFormat|.

    ``None`` leaves it untouched; :data:`NO_LINE` writes an explicit ``a:noFill`` (which
    suppresses PowerPoint's default border); an :class:`Edge` draws a solid line.
    """
    if edge is None:
        return
    if edge is NO_LINE:
        line_format.fill.background()  # -> <a:ln><a:noFill/></a:ln>
        return
    line_format.width = Pt(edge.pt)
    line_format.color.rgb = _rgb(edge.color)
    if edge.dash is not None:
        line_format.dash_style = edge.dash


def apply_borders(cell: "_Cell", borders: BorderSpec) -> None:
    """Draw the edges described by `borders` on `cell` (leaving `None` edges untouched)."""
    _set_edge(cell.borders.left, borders.left)
    _set_edge(cell.borders.right, borders.right)
    _set_edge(cell.borders.top, borders.top)
    _set_edge(cell.borders.bottom, borders.bottom)


def apply_cell_style(cell: "_Cell", style: CellStyle) -> None:
    """Apply every *set* (non-``None``) aspect of `style` to `cell`.

    Font attributes are applied to the cell's existing run(s); set the cell's text
    *before* calling so there is a run to format.
    """
    # -- fill --
    if style.fill is not None:
        if str(style.fill).lower() == "none":
            cell.fill.background()
        else:
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(style.fill)

    # -- vertical anchor --
    if style.anchor is not None:
        cell.vertical_anchor = _ANCHOR[style.anchor]

    # -- margins --
    if style.margins_cm is not None:
        ml, mr, mt, mb = style.margins_cm
        if ml is not None:
            cell.margin_left = Cm(ml)
        if mr is not None:
            cell.margin_right = Cm(mr)
        if mt is not None:
            cell.margin_top = Cm(mt)
        if mb is not None:
            cell.margin_bottom = Cm(mb)

    # -- paragraph + run font --
    for paragraph in cell.text_frame.paragraphs:
        if style.align is not None:
            paragraph.alignment = _ALIGN[style.align]
        # paragraph-level defaults cover empty cells; run-level covers existing text
        targets = [paragraph.font] + [run.font for run in paragraph.runs]
        for font in targets:
            if style.font_name is not None:
                font.name = style.font_name
            if style.size is not None:
                font.size = Pt(style.size)
            if style.bold is not None:
                font.bold = style.bold
            if style.italic is not None:
                font.italic = style.italic
            if style.color is not None:
                font.color.rgb = _rgb(style.color)

    # -- borders --
    if style.borders is not None:
        apply_borders(cell, style.borders)


# ---------------------------------------------------------------------------
# size helpers
# ---------------------------------------------------------------------------


def _distribute(total: int, weights: Sequence[int]) -> "list[int]":
    """Split `total` across len(weights) buckets proportional to `weights`.

    Uses integer math and dumps the rounding remainder into the last bucket, so the
    returned list always sums to exactly `total` (no drift between the grid and the
    frame extent).
    """
    wsum = sum(weights)
    if wsum <= 0:
        # equal split
        base = total // len(weights)
        out = [base] * len(weights)
        out[-1] = total - base * (len(weights) - 1)
        return out
    out = [int(round(total * w / wsum)) for w in weights]
    out[-1] = total - sum(out[:-1])
    return out


def min_row_height_for_font(font_pt: float, lines: int = 1, pad_cm: float = 0.0) -> "Length":
    """Return the height PowerPoint renders for `lines` lines of `font_pt` text.

    PowerPoint lays out a text line at ~1.2x the font size. A table row's stored ``h``
    is only a *minimum* (issues #480/#296): if it is smaller than this, PowerPoint grows
    the row at render time but never writes the value back, so a python-only build looks
    too short. Setting the stored height to this value keeps the file self-consistent
    without a PowerPoint round-trip.
    """
    emu = int(Pt(font_pt * 1.2) * lines) + int(Cm(pad_cm))
    return Emu(emu)


# ---------------------------------------------------------------------------
# StyledTable -- the ergonomic wrapper
# ---------------------------------------------------------------------------


class StyledTable:
    """Thin wrapper over a python-pptx |Table| with styling + sizing conveniences."""

    def __init__(self, table: "Table", graphic_frame=None):
        self._table = table
        self._graphic_frame = graphic_frame if graphic_frame is not None else table._graphic_frame

    # -- construction --------------------------------------------------------

    @classmethod
    def add(
        cls,
        shapes: "_BaseGroupShapes",
        rows: int,
        cols: int,
        x: "Length",
        y: "Length",
        cx: "Length",
        cy: "Length",
        strip_style: bool = True,
    ) -> "StyledTable":
        """Create a table on `shapes` and wrap it.

        By default the auto-injected table style is stripped (``strip_style=True``) so
        explicit per-cell fills are authoritative -- the same clean slate a PowerPoint
        clipboard paste yields.
        """
        gf = shapes.add_table(rows, cols, x, y, cx, cy)
        st = cls(gf.table, gf)
        if strip_style:
            st.strip_table_style()
        return st

    # -- pass-through accessors ---------------------------------------------

    @property
    def table(self) -> "Table":
        return self._table

    @property
    def graphic_frame(self):
        return self._graphic_frame

    def cell(self, row: int, col: int) -> "_Cell":
        return self._table.cell(row, col)

    # -- table-level style ---------------------------------------------------

    def strip_table_style(self) -> "StyledTable":
        """Remove the default table style + banding so explicit fills win.

        Leaves an empty ``<a:tblPr/>`` (what a PowerPoint paste produces).
        """
        tbl = self._table._tbl
        tblPr = tbl.tblPr
        if tblPr is not None:
            for child in list(tblPr):
                tblPr.remove(child)  # drop <a:tableStyleId> etc.
            for attr in ("firstRow", "firstCol", "lastRow", "lastCol", "bandRow", "bandCol"):
                if attr in tblPr.attrib:
                    del tblPr.attrib[attr]
        return self

    # -- content + styling ---------------------------------------------------

    def set(self, row: int, col: int, text, style: "CellStyle | None" = None) -> "_Cell":
        """Set a cell's text (coercing to ``str``) and optionally apply `style`."""
        cell = self._table.cell(row, col)
        cell.text = "" if text is None else str(text)
        if style is not None:
            apply_cell_style(cell, style)
        return cell

    def style_cell(self, row: int, col: int, style: CellStyle) -> "_Cell":
        cell = self._table.cell(row, col)
        apply_cell_style(cell, style)
        return cell

    def style_range(
        self, r0: int, c0: int, r1: int, c1: int, style: CellStyle
    ) -> "StyledTable":
        """Apply `style` to every cell in the inclusive rectangle (r0,c0)-(r1,c1)."""
        for r in range(min(r0, r1), max(r0, r1) + 1):
            for c in range(min(c0, c1), max(c0, c1) + 1):
                apply_cell_style(self._table.cell(r, c), style)
        return self

    def style_row(self, row: int, style: CellStyle, cols: "Iterable[int] | None" = None):
        ncols = len(self._table.columns)
        for c in (range(ncols) if cols is None else cols):
            apply_cell_style(self._table.cell(row, c), style)
        return self

    def style_col(self, col: int, style: CellStyle, rows: "Iterable[int] | None" = None):
        nrows = len(self._table.rows)
        for r in (range(nrows) if rows is None else rows):
            apply_cell_style(self._table.cell(r, col), style)
        return self

    def fill_rows(
        self,
        data: "Sequence[Sequence]",
        header_style: "CellStyle | None" = None,
        body_style: "CellStyle | None" = None,
        first_row_is_header: bool = True,
    ) -> "StyledTable":
        """Populate the table from a 2D sequence and style header/body rows.

        `data[0]` is treated as the header row when `first_row_is_header` is True.
        """
        for r, row_vals in enumerate(data):
            is_header = first_row_is_header and r == 0
            style = header_style if is_header else body_style
            for c, val in enumerate(row_vals):
                self.set(r, c, val, style)
        return self

    # -- merging -------------------------------------------------------------

    def merge_equal_runs(
        self,
        col: int,
        rows: "Iterable[int] | None" = None,
        skip=("", "Compromissada"),
    ) -> "StyledTable":
        """Merge vertically-adjacent cells in `col` that share the same text.

        Native replacement for the spreadsheet ``mescla_coluna`` trick. Values in
        `skip` are never merged (e.g. blanks, or a label that legitimately repeats).
        """
        nrows = len(self._table.rows)
        row_list = list(range(nrows) if rows is None else rows)
        i = 0
        while i < len(row_list) - 1:
            r = row_list[i]
            text = self._table.cell(r, col).text
            j = 1
            while (
                i + j < len(row_list)
                and self._table.cell(row_list[i + j], col).text == text
                and text not in skip
            ):
                j += 1
            if j > 1:
                top = self._table.cell(r, col)
                bottom = self._table.cell(row_list[i + j - 1], col)
                top.merge(bottom)
            i += j
        return self

    # -- sizing --------------------------------------------------------------

    def col_widths(self, widths_cm: "Sequence[float]") -> "StyledTable":
        """Set each column width from a list of cm values (frame width auto-syncs)."""
        cols = self._table.columns
        for c, w in zip(cols, widths_cm):
            c.width = Cm(w)
        return self

    def fit(self, width_cm: float, height_cm: "float | None" = None) -> "StyledTable":
        """Scale columns (and optionally rows) so the totals hit exact targets.

        Distributes proportionally with integer math so the grid sum equals the frame
        extent exactly -- the per-column-width-then-resize-again dance becomes a single
        deterministic pass. Column setters already keep ``graphic_frame.width`` in sync.
        """
        cols = self._table.columns
        cur = [c.width for c in cols]
        for c, w in zip(cols, _distribute(int(Cm(width_cm)), cur)):
            c.width = Emu(w)
        if height_cm is not None:
            rows = self._table.rows
            cur_h = [r.height for r in rows]
            for r, h in zip(rows, _distribute(int(Cm(height_cm)), cur_h)):
                r.height = Emu(h)
        return self

    def autosize_rows(
        self, font_pt: float, lines: int = 1, pad_cm: float = 0.0, header_pt: "float | None" = None
    ) -> "StyledTable":
        """Set every row height to the height PowerPoint will render (issues #480/#296).

        `header_pt` lets the first row use a different font size than the body.
        """
        rows = self._table.rows
        for i, row in enumerate(rows):
            pt = header_pt if (header_pt is not None and i == 0) else font_pt
            row.height = min_row_height_for_font(pt, lines=lines, pad_cm=pad_cm)
        return self

    def set_uniform_row_height(self, cm: float) -> "StyledTable":
        for row in self._table.rows:
            row.height = Cm(cm)
        return self


# ---------------------------------------------------------------------------
# DataFrame convenience (pandas optional)
# ---------------------------------------------------------------------------


def dataframe_to_table(
    shapes: "_BaseGroupShapes",
    df,
    x: "Length",
    y: "Length",
    cx: "Length",
    cy: "Length",
    header_style: "CellStyle | None" = None,
    body_style: "CellStyle | None" = None,
    include_header: bool = True,
    index: bool = False,
) -> "StyledTable":
    """Build a fully-styled table from a pandas DataFrame in one call.

    Returns the |StyledTable| so the caller can still ``fit``/``autosize_rows``/merge.
    """
    cols = ([df.index.name or ""] + list(df.columns)) if index else list(df.columns)
    rows = []
    if include_header:
        rows.append([str(c) for c in cols])
    for tup in df.itertuples(index=index):
        rows.append(["" if v is None else str(v) for v in tup])

    nrows, ncols = len(rows), len(cols)
    st = StyledTable.add(shapes, nrows, ncols, x, y, cx, cy)
    st.fill_rows(
        rows, header_style=header_style, body_style=body_style, first_row_is_header=include_header
    )
    return st
