.. _ext_text:

Text
====

Run-level character formatting, paragraph bullets and indents, text direction, and
font-file discovery.


Character formatting
--------------------

These are properties of |Font|, so they work anywhere a font does — on a run
(``run.font``), on a paragraph's defaults (``paragraph.font``), or on a shape's text
frame.

Strikethrough
~~~~~~~~~~~~~

`Upstream issues #339 <https://github.com/scanny/python-pptx/issues/339>`_ and
`#574 <https://github.com/scanny/python-pptx/issues/574>`_.

::

    run.font.strikethrough = True          # single strikethrough
    run.font.strikethrough = "dblStrike"   # double
    run.font.strikethrough = False         # explicitly none
    run.font.strikethrough = None          # inherit (removes the attribute)

Reading returns ``True`` for either strikethrough kind, ``False`` for an explicit none,
and ``None`` when unset.

Superscript and subscript
~~~~~~~~~~~~~~~~~~~~~~~~~

::

    run.font.superscript = True
    run.font.subscript = True

Both are backed by the same ``a:rPr@baseline`` attribute, so they are mutually
exclusive — setting one clears the other. Assigning ``None`` returns the run to the
normal baseline.

Character spacing
~~~~~~~~~~~~~~~~~

Inter-character tracking, as a |Length|::

    from pptx.util import Pt

    run.font.spacing = Pt(1.5)    # loose
    run.font.spacing = Pt(-0.5)   # condensed; negative values are valid
    run.font.spacing = None       # inherit

Capitalization
~~~~~~~~~~~~~~

::

    run.font.caps = "all"     # ALL CAPS
    run.font.caps = "small"   # Sᴍᴀʟʟ Cᴀᴘs
    run.font.caps = "none"    # explicitly no caps effect
    run.font.caps = None      # inherit

Any other value raises ``ValueError``. Note that this is a *rendering* effect: the
underlying text is unchanged, so extracting ``run.text`` returns the original casing.

Highlight colour
~~~~~~~~~~~~~~~~

The text marker/highlight, as an |RGBColor| or ``None`` to remove it::

    from pptx.dml.color import RGBColor

    run.font.highlight_color = RGBColor(0xFF, 0xFF, 0x00)
    run.font.highlight_color = None

This is the highlighter pen, which is distinct from the run's fill colour
(``font.color``) and from a shape's fill.


Bullets
-------

`Upstream issues #100 <https://github.com/scanny/python-pptx/issues/100>`_ and
`#114 <https://github.com/scanny/python-pptx/issues/114>`_.

``paragraph.bullet`` gives access to the paragraph's bullet setting. The three kinds are
mutually exclusive — setting one removes any other.

::

    # no bullet at all
    paragraph.bullet.none()

    # a glyph bullet
    paragraph.bullet.character()                    # "•", the default
    paragraph.bullet.character("▪", font="Arial", color="62993D", size_pct=80)

    # an automatically-numbered bullet
    paragraph.bullet.auto_number()                          # 1.  2.  3.
    paragraph.bullet.auto_number("romanLcPeriod", start_at=4)   # iv.  v.  vi.

Both ``character()`` and ``auto_number()`` accept the same three styling arguments,
which control the bullet mark itself rather than the paragraph text:

``font``
    Typeface for the bullet glyph — commonly ``"Arial"`` or ``"Wingdings"``.

``color``
    An ``RRGGBB`` hex string.

``size_pct``
    Bullet size relative to the text, e.g. ``80`` for 80% (range 25–400).

``number_type`` is an OOXML auto-number scheme. Common ones are ``"arabicPeriod"``
(``1.``), ``"arabicParenR"`` (``1)``), ``"romanLcPeriod"`` (``i.``), ``"romanUcPeriod"``
(``I.``) and ``"alphaUcPeriod"`` (``A.``).

Reading the current setting::

    paragraph.bullet.type          # "none" | "character" | "auto_number" | None
    paragraph.bullet.char          # the glyph, when a character bullet is set
    paragraph.bullet.number_type   # the scheme, when an auto-number bullet is set

``type`` returns ``None`` when the paragraph has no explicit bullet setting of its own
and inherits one from the layout or master.

.. note::
   A bullet is a paragraph property, not a run property, and PowerPoint indents bulleted
   text using the paragraph indents below. ``bullet.character()`` on its own sets the
   mark; pair it with ``first_line_indent`` for a hanging indent.


Paragraph indents
-----------------

Three read/write |Length| properties on a paragraph::

    from pptx.util import Cm, Pt

    paragraph.margin_left = Cm(1)           # left indent of the whole paragraph
    paragraph.margin_right = Cm(0.5)        # right indent
    paragraph.first_line_indent = Pt(-18)   # negative -> hanging indent

All three return ``None`` when not set explicitly, meaning the value is inherited from
the style hierarchy, and accept ``None`` to remove the setting.

``first_line_indent`` is relative to ``margin_left``: a positive value indents the first
line further, and a **negative** value pulls it back to produce the hanging indent used
for bulleted and numbered lists::

    paragraph.bullet.character("•")
    paragraph.margin_left = Pt(18)          # text block indented
    paragraph.first_line_indent = Pt(-18)   # bullet pulled back to the margin


Text direction
--------------

``text_frame.text_direction`` rotates the flow of text within its frame, which is how
you get vertical column headers in a table or a sidebar label in a text box::

    from pptx.enum.text import MSO_TEXT_DIRECTION

    text_frame.text_direction = MSO_TEXT_DIRECTION.VERTICAL_270

.. csv-table::
   :header: "Member", "Effect"
   :widths: 35, 65

   "``HORIZONTAL``", "Horizontal flow (the default)"
   "``VERTICAL``", "Vertical, characters rotated 90° clockwise"
   "``VERTICAL_270``", "Vertical, characters rotated 270° clockwise (reads bottom-to-top)"
   "``EAST_ASIAN_VERTICAL``", "East Asian vertical flow"
   "``MONGOLIAN_VERTICAL``", "Mongolian vertical flow"
   "``WORD_ART_VERTICAL``", "WordArt-style vertical (stacked characters)"
   "``WORD_ART_VERTICAL_RTL``", "WordArt-style vertical, right-to-left"

Returns ``None`` when unset; assigning ``None`` removes the attribute.

.. tip::
   For a rotated table column header, set the direction on the *cell's* text frame:
   ``cell.text_frame.text_direction = MSO_TEXT_DIRECTION.VERTICAL_270``. Remember to give
   the row enough height for the rotated text.


Font files on Linux
-------------------

Mirrors `upstream PR #1039 <https://github.com/scanny/python-pptx/pull/1039>`_.

``TextFrame.fit_text()`` measures text by loading the actual font file, so it needs to
find the fonts installed on the machine. Upstream only knows the macOS and Windows
locations and raises ``OSError: unsupported operating system`` everywhere else — which
includes most CI runners and containers.

This fork adds the standard Linux locations, searched automatically:

* ``/usr/share/fonts``
* ``/usr/local/share/fonts``
* ``~/.local/share/fonts``
* ``~/.fonts``

Custom directories
~~~~~~~~~~~~~~~~~~

The ``PYTHON_PPTX_FONT_DIRECTORY`` environment variable names additional directories,
delimited like ``PATH`` (``:`` on POSIX, ``;`` on Windows):

.. code-block:: bash

    export PYTHON_PPTX_FONT_DIRECTORY=/opt/brand-fonts:/srv/shared/fonts

They are searched **first**, ahead of the platform defaults, and are *added to* them
rather than replacing them. Two uses:

* bundling brand fonts with an application, without installing them system-wide;
* an escape hatch on any platform whose default locations the library doesn't know —
  setting the variable is enough to make ``fit_text()`` work there.
