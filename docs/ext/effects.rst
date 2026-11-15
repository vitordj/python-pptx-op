.. _ext_effects:

Shapes, colour and effects
==========================

Transparency, shadows, line arrow heads, and alt text.


Colour transparency
-------------------

`Upstream issue #62 <https://github.com/scanny/python-pptx/issues/62>`_.

Any |ColorFormat| — a shape fill, a line colour, a font colour, a shadow colour — takes a
``transparency`` value, a float from ``0.0`` (fully opaque, the default) to ``1.0``
(fully transparent)::

    from pptx.dml.color import RGBColor

    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0x1F, 0x77, 0xB4)
    shape.fill.fore_color.transparency = 0.35    # 35% transparent

    shape.line.color.rgb = RGBColor(0, 0, 0)
    shape.line.color.transparency = 0.5

It works with theme colours as well as explicit RGB, since it is stored as an
``<a:alpha>`` child of whatever colour element is in use.

.. warning::
   Set a colour **before** its transparency. Assigning ``transparency`` when no colour is
   defined raises ``ValueError`` — there is no colour element to attach the alpha value
   to. Values outside ``0.0``–``1.0`` also raise ``ValueError``.

.. note::
   OOXML stores *opacity* (``alpha``), not transparency, so the file contains
   ``1.0 - transparency``. The property does the conversion, which is why the number you
   assign matches the "Transparency" slider in the PowerPoint UI.


Picture transparency
--------------------

Pictures use a different mechanism from fills, so they have their own property with the
same meaning and range::

    picture.transparency = 0.5   # 50% transparent
    picture.transparency = 0.0   # fully opaque (the default; removes the element)

Reading returns ``0.0`` for a picture with no transparency set. Out-of-range values raise
``ValueError``. This is the "washout" effect used for background images behind text.


Shadows
-------

Upstream exposes only ``shadow.inherit``. This fork adds the shadow's actual parameters,
backed by ``a:effectLst/a:outerShdw``::

    from pptx.dml.color import RGBColor
    from pptx.util import Pt

    shadow = shape.shadow
    shadow.inherit = False        # stop inheriting from the theme first

    shadow.blur_radius = Pt(8)    # softness
    shadow.distance = Pt(4)       # offset distance from the shape
    shadow.direction = 45.0       # degrees; 45° casts to the lower right
    shadow.color.rgb = RGBColor(0, 0, 0)

.. csv-table::
   :header: "Property", "Type", "Meaning"
   :widths: 25, 20, 55

   "``blur_radius``", "|Length|", "Blur softness. ``Pt(3)`` is subtle, ``Pt(7)`` a strong card shadow"
   "``distance``", "|Length|", "How far the shadow is offset from the shape"
   "``direction``", "``float``", "Offset angle in degrees, clockwise from east"
   "``color``", "|ColorFormat|", "Shadow colour; combine with ``transparency`` above for a soft shadow"
   "``soft_edge_radius``", "|Length|", "Feathers the *shape's own* edges (``a:softEdge``)"

Direction runs clockwise from east, matching the OOXML convention:

.. csv-table::
   :header: "Degrees", "Shadow falls"
   :widths: 20, 80

   "0", "right"
   "45", "lower right — the familiar drop shadow"
   "90", "straight down"
   "180", "left"
   "270", "up"

**Reading.** Every property returns ``None`` when the shape has no explicit shadow, so
``shadow.color`` is ``None`` rather than a ``ColorFormat`` in that case — check before
dereferencing it.

**Seeding.** Setting any one property on a shape with no shadow creates one first, using
a plausible default (a soft black shadow offset downwards at 63% opacity), so a single
assignment gives a visible result instead of a zero-distance artifact. Assign the other
properties to take full control.

**Removing.** ``shadow.remove_shadow()`` deletes the outer shadow. That is different from
``shadow.inherit = True``, which restores inheritance from the theme.

Soft edges
~~~~~~~~~~

``soft_edge_radius`` feathers the shape's own outline rather than its shadow::

    shape.shadow.soft_edge_radius = Pt(6)
    shape.shadow.soft_edge_radius = None   # remove

It lives on the same effect list, which is why it is reached through ``shadow``.


Line arrow heads and caps
-------------------------

`Upstream issue #375 <https://github.com/scanny/python-pptx/issues/375>`_.

``line.begin_arrow`` and ``line.end_arrow`` describe the two ends of any line —
a connector, a shape outline, a chart line::

    connector.line.end_arrow.type = "triangle"
    connector.line.end_arrow.width = "lg"
    connector.line.end_arrow.length = "lg"

    connector.line.begin_arrow.type = "oval"
    connector.line.begin_arrow.remove()   # drop the element entirely

.. csv-table::
   :header: "Property", "Values"
   :widths: 20, 80

   "``type``", "``none``, ``triangle``, ``stealth``, ``diamond``, ``oval``, ``arrow``"
   "``width``", "``sm``, ``med``, ``lg``"
   "``length``", "``sm``, ``med``, ``lg``"

Each returns ``None`` when not set. ``begin_arrow`` maps to ``a:headEnd`` and
``end_arrow`` to ``a:tailEnd``; "begin" and "end" refer to the direction the line was
drawn in.

The line's end cap is a property of the line itself::

    connector.line.cap = "rnd"    # "rnd" | "sq" | "flat"

``rnd`` rounds the ends, ``sq`` extends a square cap past the end point, and ``flat``
stops exactly at it.


Alt text
--------

Every shape — auto shape, picture, table, chart, connector, group — carries an
accessibility description::

    picture.alt_text = "Quarterly revenue by region, bar chart"
    shape.alt_text = ""      # clears the attribute

Reading returns ``""`` when no description is set, and assigning ``""`` removes the
attribute rather than storing an empty one, which is what PowerPoint does.

This is the "Description" field in PowerPoint's *Alt Text* pane, stored as
``p:cNvPr@descr``. Screen readers announce it, and it is what accessibility checkers look
for — worth setting on every non-decorative image in a generated deck.
