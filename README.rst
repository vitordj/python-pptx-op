*python-pptx-op* is a friendly fork of `python-pptx`_, the Python library for creating,
reading, and updating PowerPoint (.pptx) files.

Tribute
-------

python-pptx was created and maintained for over a decade by **Steve Canny** (`@scanny`_).
He is the brilliant mind behind this project: he designed its architecture, wrote its
excellent documentation, and spent countless hours patiently answering the community on
GitHub and Stack Overflow. Everything in this fork stands on his work, and we are deeply
grateful to him. If this library is useful to you, that is first and foremost his merit —
please consider starring the `upstream repository`_.

This fork exists only to ship a handful of additional features on top of upstream
release 1.0.2, keeping the original MIT license and full credit to its author. The
import name remains ``pptx`` and the public API is a strict superset of upstream, so it
is a drop-in replacement.

.. _`python-pptx`: https://github.com/scanny/python-pptx
.. _`@scanny`: https://github.com/scanny
.. _`upstream repository`: https://github.com/scanny/python-pptx

About the library
-----------------

A typical use would be generating a PowerPoint presentation from dynamic content such as
a database query, analytics output, or a JSON payload, perhaps in response to an HTTP
request and downloading the generated PPTX file in response. It runs on any Python
capable platform, including macOS and Linux, and does not require the PowerPoint
application to be installed or licensed.

It can also be used to analyze PowerPoint files from a corpus, perhaps to extract search
indexing text and images.

It can also be used to simply automate the production of a slide or two that would be
tedious to get right by hand, which is how this all got started.

More information is available in the `python-pptx documentation`_, which covers the
entire base API of this fork.

Browse `examples with screenshots`_ to get a quick idea what you can do with
python-pptx.

Installation
------------

::

    pip install python-pptx-op

Because the import package is still ``pptx``, do not install this fork alongside
upstream ``python-pptx`` in the same environment — install one or the other.

Features added in this fork
---------------------------

- ChartEx charts (``pptx.chartex``) and a table-building toolkit (``pptx.tablekit``)
- ``Font``: character spacing, caps effect, and highlight color
- ``_Run``: strikethrough and super/subscript; paragraph bullet control
- ``_Paragraph``: ``margin_left``, ``margin_right``, and ``first_line_indent``
- ``TextFrame.text_direction`` (vertical text)
- ``ColorFormat`` transparency (``a:alpha``)
- ``ShadowFormat``: blur radius, distance, direction, color, and soft-edge radius
- ``LineFormat``: arrow heads and end cap
- ``BaseShape.alt_text`` getter/setter
- ``Picture.transparency``
- Doughnut chart hole size and first-slice angle
- Linux font directories and ``PYTHON_PPTX_FONT_DIRECTORY`` override

.. _`python-pptx documentation`:
   https://python-pptx.readthedocs.org/en/latest/

.. _`examples with screenshots`:
   https://python-pptx.readthedocs.org/en/latest/user/quickstart.html
