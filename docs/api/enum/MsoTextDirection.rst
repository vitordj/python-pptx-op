.. _MsoTextDirection:

``MSO_TEXT_DIRECTION``
======================

Specifies the direction of text flow within a text frame.

Example::

    from pptx.enum.text import MSO_TEXT_DIRECTION

    cell.text_frame.text_direction = MSO_TEXT_DIRECTION.VERTICAL_270

----

HORIZONTAL
    Horizontal text flow (the default).

VERTICAL
    Vertical text flow, characters rotated 90° clockwise.

VERTICAL_270
    Vertical text flow, characters rotated 270° clockwise.

EAST_ASIAN_VERTICAL
    East Asian vertical text flow.

MONGOLIAN_VERTICAL
    Mongolian vertical text flow.

WORD_ART_VERTICAL
    WordArt-style vertical text flow (stacked characters).

WORD_ART_VERTICAL_RTL
    WordArt-style vertical right-to-left text flow.
