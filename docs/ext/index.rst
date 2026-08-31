.. _ext_index:

Fork extensions
===============

This section documents everything ``python-pptx-op`` adds on top of upstream
python-pptx 1.0.2. The rest of this documentation describes the upstream API, all of
which still applies unchanged — the additions below are extra properties on objects you
already use, plus two new modules (``pptx.tablekit`` and ``pptx.chartex``).

Nothing upstream was changed or removed. If a page in the User Guide contradicts a page
here, the page here describes an *addition*, not a replacement.


.. toctree::
   :maxdepth: 2

   tables
   charts
   text
   effects


At a glance
-----------

Tables — :ref:`ext_tables`
~~~~~~~~~~~~~~~~~~~~~~~~~~

.. csv-table::
   :header: "Capability", "API"
   :widths: 40, 60

   "Per-cell borders, four edges plus two diagonals", "``cell.borders.left`` … (each a |LineFormat|)"
   "Reusable cell styles", "``pptx.tablekit.CellStyle`` / ``BorderSpec`` / ``Edge``"
   "Styled table builder, sizing and merging", "``pptx.tablekit.StyledTable``"
   "Build a table straight from a DataFrame", "``pptx.tablekit.dataframe_to_table()``"

Charts — :ref:`ext_charts`
~~~~~~~~~~~~~~~~~~~~~~~~~~

.. csv-table::
   :header: "Capability", "API"
   :widths: 40, 60

   "Office 2016 'modern' charts", "``shapes.add_chartex()``, ``ChartExData``, ``XL_CHART_EX_TYPE``"
   "Doughnut hole size and rotation", "``plot.hole_size``, ``plot.first_slice_angle``"
   "Date-axis time units", "``axis.base_time_unit`` / ``.major_time_unit`` / ``.minor_time_unit``"

Text — :ref:`ext_text`
~~~~~~~~~~~~~~~~~~~~~~

.. csv-table::
   :header: "Capability", "API"
   :widths: 40, 60

   "Strikethrough, superscript, subscript", "``font.strikethrough``, ``font.superscript``, ``font.subscript``"
   "Character spacing, caps, highlight", "``font.spacing``, ``font.caps``, ``font.highlight_color``"
   "Bullet control", "``paragraph.bullet.none()`` / ``.character()`` / ``.auto_number()``"
   "Paragraph indents", "``paragraph.margin_left`` / ``.margin_right`` / ``.first_line_indent``"
   "Vertical and rotated text", "``text_frame.text_direction``"
   "Font-file lookup on Linux", "``PYTHON_PPTX_FONT_DIRECTORY`` environment variable"

Shapes, colour and effects — :ref:`ext_effects`
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

.. csv-table::
   :header: "Capability", "API"
   :widths: 40, 60

   "Colour transparency (fill and line)", "``color.transparency``"
   "Picture transparency", "``picture.transparency``"
   "Shadow parameters and soft edges", "``shape.shadow.blur_radius`` / ``.distance`` / ``.direction`` / ``.color`` / ``.soft_edge_radius``"
   "Line arrow heads and end caps", "``line.begin_arrow``, ``line.end_arrow``, ``line.cap``"
   "Alt text on any shape", "``shape.alt_text``"


Design notes
------------

A few conventions hold across all of the additions.

**Everything is additive.** No upstream signature changed. Existing code produces
byte-identical output to upstream python-pptx unless it opts into a new property.

**"Not set" is ``None``.** Every new property returns ``None`` when the underlying XML
attribute or element is absent, meaning the value is inherited from the placeholder,
layout, master or theme. Assigning ``None`` removes the setting again and restores
inheritance, rather than writing an explicit "off" value — which matters, because an
explicit value blocks the inheritance a template relies on.

**Values use the library's existing types.** Lengths are |Length| objects (``Pt``,
``Cm``, ``Emu``), colours are |RGBColor|, and enumerated values are members of the
enumerations in ``pptx.enum``. Where the OOXML schema stores an odd unit — centipoints
for character spacing, sixtieths of a degree for shadow direction, percent-times-1000
for opacity — the property converts, so you always work in the familiar unit.
