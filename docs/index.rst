
python-pptx-op
==============

Release v\ |version| (:ref:`Installation <install>`)

*python-pptx-op* is a drop-in fork of `python-pptx`_ that adds the PowerPoint formatting
features the original never shipped. It tracks upstream python-pptx 1.0.2 and keeps the
same import name and API, so ``import pptx`` and existing code keep working unchanged.

*python-pptx* is a Python library for creating, reading, and updating PowerPoint (.pptx)
files.

A typical use would be generating a PowerPoint presentation from dynamic content such as
a database query, analytics output, or a JSON payload, perhaps in response to an HTTP
request and downloading the generated PPTX file in response. It runs on any Python
capable platform, including macOS and Linux, and does not require the PowerPoint
application to be installed or licensed.

It can also be used to analyze PowerPoint files from a corpus, perhaps to extract search
indexing text and images.

In can also be used to simply automate the production of a slide or two that would be
tedious to get right by hand, which is how this all got started.

**New to this fork?** Start with :ref:`ext_index`, which documents everything added on
top of upstream. Everything else in this documentation describes the upstream API, all
of which still applies.

.. _`python-pptx`: https://github.com/scanny/python-pptx


Philosophy
----------

|pp| aims to broadly support the PowerPoint format (PPTX, PowerPoint 2007 and later),
but its primary commitment is to be *industrial-grade*, that is, suitable for use in a
commercial setting. Maintaining this robustness requires a high engineering standard
which includes a comprehensive two-level (e2e + unit) testing regimen. This discipline
comes at a cost in development effort/time, but we consider reliability to be an
essential requirement.


Feature Support
---------------

|pp| has the following capabilities:

* Round-trip any Open XML presentation (.pptx file) including all its elements
* Add slides
* Populate text placeholders, for example to create a bullet slide
* Add image to slide at arbitrary position and size
* Add textbox to a slide; manipulate text font size and bold
* Add table to a slide
* Add auto shapes (e.g. polygons, flowchart shapes, etc.) to a slide
* Add and manipulate column, bar, line, and pie charts
* Access and change core document properties such as title and subject
* And many others ...

Even with all |pp| does, the PowerPoint document format is very rich and there are still
features |pp| does not support.


This fork adds to that list: table cell borders, the modern Office 2016 chart family
(waterfall, treemap, funnel, sunburst, box & whisker), shadow parameters, colour and
picture transparency, line arrow heads, strikethrough, text highlight, bullet control,
paragraph indents, vertical text, and alt text. See :ref:`ext_index`.


Fork extensions
---------------

.. toctree::
   :maxdepth: 2

   ext/index


User Guide
----------

.. toctree::
   :maxdepth: 1

   user/intro
   user/install
   user/quickstart
   user/presentations
   user/slides
   user/understanding-shapes
   user/autoshapes
   user/placeholders-understanding
   user/placeholders-using
   user/text
   user/charts
   user/table
   user/notes
   user/use-cases
   user/concepts


Community Guide
---------------

.. toctree::
   :maxdepth: 1

   community/faq
   community/support
   community/updates


.. _api:

API Documentation
-----------------

.. toctree::
   :maxdepth: 2

   api/presentation
   api/slides
   api/shapes
   api/placeholders
   api/table
   api/chart-data
   api/chart
   api/text
   api/action
   api/dml
   api/image
   api/exc
   api/util
   api/enum/index


Contributor Guide
-----------------

.. toctree::
   :maxdepth: 1

   dev/runtests
   dev/xmlchemy
   dev/development_practices
   dev/philosophy
   dev/analysis/index
   dev/resources/index
