.. _ext_tables:

Tables
======

Two additions: **cell borders**, the primitive that upstream never exposed, and
**``pptx.tablekit``**, a small layer on top of it for building fully-formatted tables
without writing four nested loops.


Cell borders
------------

`Upstream issue #71 <https://github.com/scanny/python-pptx/issues/71>`_.

A table cell exposes its six border lines through ``cell.borders``. Each one is an
ordinary |LineFormat|, identical to the ``.line`` of any shape, so width, colour, dash
style and fill all work the way you already know::

    from pptx.dml.color import RGBColor
    from pptx.enum.dml import MSO_LINE_DASH_STYLE
    from pptx.util import Pt

    cell = table.cell(0, 0)

    cell.borders.bottom.width = Pt(1.5)
    cell.borders.bottom.color.rgb = RGBColor(0x8C, 0xC6, 0x40)
    cell.borders.left.dash_style = MSO_LINE_DASH_STYLE.DASH

The six edges are:

``left``, ``right``, ``top``, ``bottom``
    The four sides, backed by ``a:lnL``, ``a:lnR``, ``a:lnT`` and ``a:lnB``.

``tl_to_br``, ``bl_to_tr``
    The two diagonals, backed by ``a:lnTlToBr`` and ``a:lnBlToTr``.

.. note::
   PowerPoint applies borders **cell by cell**, not to the table as a whole. To outline
   a block of cells you set the outward-facing edge of each cell on the perimeter. Two
   adjacent cells each have their own edge between them, so a line drawn on the right of
   one and the left of its neighbour is drawn twice.

Suppressing a border
~~~~~~~~~~~~~~~~~~~~

An edge with no explicit ``a:lnX`` element is not "no border" — PowerPoint draws its
*default* border there. To get genuinely no line, write an explicit no-fill::

    cell.borders.left.fill.background()   # -> <a:lnL><a:noFill/></a:lnL>

This is the usual cause of stray vertical lines appearing between columns when you only
meant to draw a bottom rule. ``tablekit`` handles it for you with :data:`NO_LINE`, below.


tablekit
--------

``pptx.tablekit`` turns the per-cell primitives into reusable, composable *styles*. It
is entirely optional — it is built on the public API and adds no new XML behaviour.

Quick start
~~~~~~~~~~~

::

    from pptx.tablekit import BorderSpec, CellStyle, Edge, StyledTable
    from pptx.util import Cm

    HEADER = CellStyle(
        font_name="Montserrat", size=9, bold=True, color="FFFFFF", fill="62993D",
        align="center", anchor="middle",
        borders=BorderSpec.box(Edge("FFFFFF", 0.5)),
    )
    BODY = CellStyle(
        font_name="Montserrat", size=9, color="000000", fill="FFFFFF",
        align="center", anchor="middle",
        borders=BorderSpec.underline(Edge("D3D3D3", 0.5)),
    )

    st = StyledTable.add(slide.shapes, rows=7, cols=5,
                         x=Cm(13.33), y=Cm(3.73), cx=Cm(19), cy=Cm(4.6))
    st.fill_rows(data, header_style=HEADER, body_style=BODY)  # data: list[list[str]]
    st.fit(width_cm=19)          # column widths sum to exactly 19 cm
    st.autosize_rows(font_pt=9)  # row heights match what PowerPoint renders

Colours are given as hex strings (``"62993D"``, with or without a leading ``#``) or as
|RGBColor|; either is accepted anywhere a colour is expected.


Styles
~~~~~~

``Edge(color="000000", pt=0.5, dash=None)``
    One border line: a colour, a width in points, and an optional
    ``MSO_LINE_DASH_STYLE``.

``BorderSpec(left=None, right=None, top=None, bottom=None)``
    The edges of a cell. Each edge is one of three things, and the distinction matters:

    * ``None`` — leave the edge alone; PowerPoint may draw its default border.
    * an ``Edge`` — draw that line.
    * ``NO_LINE`` — *explicitly* draw nothing, suppressing the default border.

    Four constructors cover the common cases::

        BorderSpec.box(edge)        # all four edges
        BorderSpec.underline(edge)  # bottom rule, sides suppressed
        BorderSpec.hlines(edge)     # top and bottom rules, sides suppressed
        BorderSpec.topline(edge)    # a single top rule, e.g. above a subtotal

    ``underline``, ``hlines`` and ``topline`` take ``suppress_sides=True`` by default,
    which is what stops the stray vertical lines described above. Pass
    ``suppress_sides=False`` to leave the side edges untouched.

``CellStyle(...)``
    A reusable bundle of cell formatting. Every field is optional and ``None`` means
    "don't touch this aspect", so styles compose rather than overwrite:

    .. csv-table::
       :header: "Field", "Meaning"
       :widths: 30, 70

       "``font_name``", "Typeface name"
       "``size``", "Font size in points"
       "``bold``, ``italic``", "Booleans"
       "``color``", "Text colour"
       "``fill``", "Cell fill colour; the literal string ``none`` means no fill"
       "``align``", "``left`` | ``center`` | ``right`` | ``justify``"
       "``anchor``", "``top`` | ``middle`` | ``bottom``"
       "``borders``", "A ``BorderSpec``"
       "``margins_cm``", "``(left, right, top, bottom)`` in cm; ``None`` entries keep the default"

    ``CellStyle`` is immutable. Use ``derive()`` to make a variant, which is how you
    express conditional formatting::

        TOTAL = BODY.derive(bold=True, fill="EEEEEE")
        NEGATIVE = BODY.derive(color="C00000")

Applying a style directly
~~~~~~~~~~~~~~~~~~~~~~~~~

If you already have a table and just want the styling helpers::

    from pptx.tablekit import apply_cell_style, apply_borders

    apply_cell_style(cell, BODY)
    apply_borders(cell, BorderSpec.box(Edge("000000", 1.0)))

.. note::
   Font attributes are applied to the runs that exist at the time of the call, so set a
   cell's text **before** styling it. ``StyledTable.set()`` and ``fill_rows()`` do this
   in the right order for you.


StyledTable
~~~~~~~~~~~

``StyledTable`` wraps a |Table| and returns ``self`` from its mutators, so calls chain.

**Creating or wrapping**

``StyledTable.add(shapes, rows, cols, x, y, cx, cy, strip_style=True)``
    Add a table to a shape collection and wrap it.

``StyledTable(table)``
    Wrap a table you already have (for example ``graphic_frame.table``).

``.table`` and ``.graphic_frame`` give you the underlying objects back at any point, so
you never lose access to the full python-pptx API.

**Content and styling**

.. csv-table::
   :header: "Method", "Does"
   :widths: 45, 55

   "``set(row, col, text, style=None)``", "Set one cell's text, optionally styling it"
   "``fill_rows(data, header_style=, body_style=, first_row_is_header=True)``", "Populate from a 2-D sequence, styling the header and body rows"
   "``style_cell(row, col, style)``", "Style one cell"
   "``style_range(r0, c0, r1, c1, style)``", "Style an inclusive rectangle"
   "``style_row(row, style, cols=None)``", "Style a whole row, or selected columns of it"
   "``style_col(col, style, rows=None)``", "Style a whole column, or selected rows of it"

**Sizing**

.. csv-table::
   :header: "Method", "Does"
   :widths: 45, 55

   "``col_widths(widths_cm)``", "Set each column width from a list of cm values"
   "``fit(width_cm, height_cm=None)``", "Scale columns (and optionally rows) so the totals hit exact targets"
   "``autosize_rows(font_pt, lines=1, pad_cm=0.0, header_pt=None)``", "Set row heights to what PowerPoint will actually render"
   "``set_uniform_row_height(cm)``", "Give every row the same height"

**Merging**

``merge_equal_runs(col, rows=None, skip=("",))``
    Merge vertically-adjacent cells in a column that share the same text — the usual way
    to present a grouped category label once instead of repeated down the column. Values
    listed in ``skip`` are never merged; by default only blank cells are protected. Pass
    e.g. ``skip=("", "Subtotal")`` to also keep a label that legitimately repeats.

**Table style**

``strip_table_style()``
    Remove the table style and banding that ``add_table`` injects. See below.


Two PowerPoint behaviours tablekit works around
-----------------------------------------------

These are worth understanding even if you never use ``tablekit``, because they explain
results that otherwise look like bugs.

Banding fights explicit fills
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

``shapes.add_table()`` injects a default *table style* — a ``tableStyleId`` plus the
``firstRow`` and ``bandRow`` flags. That style applies its own banded fills, which
compete with per-cell fills you set explicitly. ``StyledTable.add()`` strips it by
default (``strip_style=True``), leaving an empty ``<a:tblPr/>``, which is exactly what
pasting a table from the PowerPoint clipboard produces. Pass ``strip_style=False`` to
keep the injected style.

Row heights are only a minimum
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

`Upstream issues #480 <https://github.com/scanny/python-pptx/issues/480>`_ and
`#296 <https://github.com/scanny/python-pptx/issues/296>`_.

The row height stored in the XML is a *minimum*. PowerPoint grows a row at render time
to fit its text but never writes the grown value back to the file. A deck built purely
from Python therefore looks too short until it has been opened and saved once in real
PowerPoint — which is impossible in a headless pipeline, and makes the file inconsistent
for any other consumer.

``autosize_rows()`` sets the stored height to the height PowerPoint *would* render,
computed as roughly 1.2 × the font size per line of text, plus margins::

    st.autosize_rows(font_pt=9)                  # every row, single line of 9 pt text
    st.autosize_rows(font_pt=9, header_pt=11)    # a taller first row
    st.autosize_rows(font_pt=9, lines=2)         # allow for two lines
    st.autosize_rows(font_pt=9, pad_cm=0.1)      # extra breathing room

The same calculation is available on its own::

    from pptx.tablekit import min_row_height_for_font

    row.height = min_row_height_for_font(font_pt=9, lines=2, pad_cm=0.1)


From a pandas DataFrame
-----------------------

``dataframe_to_table()`` builds a styled table in a single call. pandas is **not** a
dependency of this library; the function simply accepts anything with the DataFrame
interface it uses::

    from pptx.tablekit import dataframe_to_table

    st = dataframe_to_table(
        slide.shapes, df,
        x=Cm(2), y=Cm(2), cx=Cm(24), cy=Cm(8),
        header_style=HEADER, body_style=BODY,
        include_header=True, index=False,
    )
    st.fit(width_cm=24).autosize_rows(font_pt=9)

Every value is converted with ``str()``, so format numbers and dates in the DataFrame
before handing it over.
