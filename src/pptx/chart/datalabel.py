"""Data label-related objects."""

from __future__ import annotations

from pptx.chart._scatter_label_layout import hybrid_cluster_offsets
from pptx.oxml.ns import qn
from pptx.text.text import Font, TextFrame
from pptx.util import lazyproperty


class DataLabels(object):
    """Provides access to properties of data labels for a plot or a series.

    This is not a collection and does not provide access to individual data
    labels. Access to individual labels is via the |Point| object. The
    properties this object provides control formatting of *all* the data
    labels in its scope.
    """

    def __init__(self, dLbls):
        super(DataLabels, self).__init__()
        self._element = dLbls

    @lazyproperty
    def font(self):
        """
        The |Font| object that provides access to the text properties for
        these data labels, such as bold, italic, etc.
        """
        defRPr = self._element.defRPr
        font = Font(defRPr)
        return font

    @property
    def number_format(self):
        """
        Read/write string specifying the format for the numbers on this set
        of data labels. Returns 'General' if no number format has been set.
        Note that this format string has no effect on rendered data labels
        when :meth:`number_format_is_linked` is |True|. Assigning a format
        string to this property automatically sets
        :meth:`number_format_is_linked` to |False|.
        """
        numFmt = self._element.numFmt
        if numFmt is None:
            return "General"
        return numFmt.formatCode

    @number_format.setter
    def number_format(self, value):
        self._element.get_or_add_numFmt().formatCode = value
        self.number_format_is_linked = False

    @property
    def number_format_is_linked(self):
        """
        Read/write boolean specifying whether number formatting should be
        taken from the source spreadsheet rather than the value of
        :meth:`number_format`.
        """
        numFmt = self._element.numFmt
        if numFmt is None:
            return True
        souceLinked = numFmt.sourceLinked
        if souceLinked is None:
            return True
        return numFmt.sourceLinked

    @number_format_is_linked.setter
    def number_format_is_linked(self, value):
        numFmt = self._element.get_or_add_numFmt()
        numFmt.sourceLinked = value

    @property
    def position(self):
        """
        Read/write :ref:`XlDataLabelPosition` enumeration value specifying
        the position of the data labels with respect to their data point, or
        |None| if no position is specified. Assigning |None| causes
        PowerPoint to choose the default position, which varies by chart
        type.
        """
        dLblPos = self._element.dLblPos
        if dLblPos is None:
            return None
        return dLblPos.val

    @position.setter
    def position(self, value):
        if value is None:
            self._element._remove_dLblPos()
            return
        self._element.get_or_add_dLblPos().val = value

    @property
    def show_category_name(self):
        """Read/write. True when name of category should appear in label."""
        return self._element.get_or_add_showCatName().val

    @show_category_name.setter
    def show_category_name(self, value):
        self._element.get_or_add_showCatName().val = bool(value)

    @property
    def show_legend_key(self):
        """Read/write. True when data label displays legend-color swatch."""
        return self._element.get_or_add_showLegendKey().val

    @show_legend_key.setter
    def show_legend_key(self, value):
        self._element.get_or_add_showLegendKey().val = bool(value)

    @property
    def show_percentage(self):
        """Read/write. True when data label displays percentage.

        This option is not operative on all chart types. Percentage appears
        on polar charts such as pie and donut.
        """
        return self._element.get_or_add_showPercent().val

    @show_percentage.setter
    def show_percentage(self, value):
        self._element.get_or_add_showPercent().val = bool(value)

    @property
    def show_series_name(self):
        """Read/write. True when data label displays series name."""
        return self._element.get_or_add_showSerName().val

    @show_series_name.setter
    def show_series_name(self, value):
        self._element.get_or_add_showSerName().val = bool(value)

    @property
    def show_value(self):
        """Read/write. True when label displays numeric value of datapoint."""
        return self._element.get_or_add_showVal().val

    @show_value.setter
    def show_value(self, value):
        self._element.get_or_add_showVal().val = bool(value)


class XyDataLabels(DataLabels):
    """Data labels for an XY (scatter) series.

    Extends |DataLabels| with cell-linked labels (Office 2013+): each label's
    text is driven from the chart's worksheet, and any displaced label gets a
    native leader line -- no manual text boxes or connectors on the slide.
    """

    def apply_hybrid_scatter_offsets(
        self, labels, *, workbook_column=3, show_leader_lines=True, **offset_kwargs
    ):
        """Compute hybrid cluster offsets for this series and apply them.

        Convenience wrapper that reads this series' X/Y values from the chart,
        calls :func:`hybrid_cluster_offsets` (forwarding keywords such as
        ``magnitude`` or ``cluster_eps`` in *offset_kwargs*), and hands the
        result to :meth:`set_values_from_cells`.
        """
        ser = self._scatter_ser()
        xs = self._point_values(ser.xVal)
        ys = self._point_values(ser.yVal)
        offsets = hybrid_cluster_offsets(xs, ys, labels, **offset_kwargs)
        self.set_values_from_cells(
            labels,
            workbook_column=workbook_column,
            layout_offsets=offsets,
            show_leader_lines=show_leader_lines,
        )

    def set_values_from_cells(
        self, labels, *, workbook_column=3, layout_offsets=None, show_leader_lines=True
    ):
        """Show *labels* as cell-linked scatter data labels.

        *labels* must have one entry per data point. *layout_offsets*, when
        provided, is a same-length sequence in which each item is either |None|
        (let PowerPoint place that label automatically) or an ``(dx, dy)``
        factorial offset; only the offset points receive an individual `c:dLbl`
        with a leader line. Positive *dy* moves a label downward. *workbook_column*
        (1-based) selects the worksheet column referenced by the label range.
        """
        labels = [str(label) for label in labels]
        if not labels:
            raise ValueError("labels must contain at least one label")
        ser = self._scatter_ser()
        point_count = ser.yVal_ptCount_val
        if point_count and len(labels) != point_count:
            raise ValueError(
                "labels has %d entries but series has %d points" % (len(labels), point_count)
            )
        if layout_offsets is not None and len(layout_offsets) != len(labels):
            raise ValueError("layout_offsets must be the same length as labels")

        self._element.use_cell_range_labels(show_leader_lines=show_leader_lines)
        ser.set_datalabels_range(self._range_ref(workbook_column, len(labels)), labels)

        if layout_offsets is None:
            return
        for idx, offset in enumerate(layout_offsets):
            if offset is None:
                continue
            x, y = float(offset[0]), float(offset[1])
            if abs(x) < 1e-9 and abs(y) < 1e-9:
                continue
            self._element.add_cell_range_dLbl(idx, x, y, labels[idx])

    @staticmethod
    def _point_values(num_data_source):
        """List of the float values cached in a `c:xVal`/`c:yVal` numeric source."""
        if num_data_source is None:
            return []
        return [num_data_source.pt_v(idx) for idx in range(num_data_source.ptCount_val)]

    @staticmethod
    def _range_ref(column, count):
        """Worksheet reference like ``Sheet1!$C$2:$C${count+1}`` for *column*."""
        letters = ""
        col = column
        while col:
            col, rem = divmod(col - 1, 26)
            letters = chr(ord("A") + rem) + letters
        return "Sheet1!$%s$2:$%s$%d" % (letters, letters, count + 1)

    def _scatter_ser(self):
        """The `c:ser` element owning these labels; guard against non-scatter."""
        ser = self._element.getparent()
        xChart = None if ser is None else ser.getparent()
        if xChart is None:
            raise ValueError("data labels are not attached to a series")
        if xChart.tag != qn("c:scatterChart"):
            raise NotImplementedError("cell-linked labels are only supported on XY scatter charts")
        return ser


class DataLabel(object):
    """
    The data label associated with an individual data point.
    """

    def __init__(self, ser, idx):
        super(DataLabel, self).__init__()
        self._ser = self._element = ser
        self._idx = idx

    @lazyproperty
    def font(self):
        """The |Font| object providing text formatting for this data label.

        This font object is used to customize the appearance of automatically
        inserted text, such as the data point value. The font applies to the
        entire data label. More granular control of the appearance of custom
        data label text is controlled by a font object on runs in the text
        frame.
        """
        txPr = self._get_or_add_txPr()
        text_frame = TextFrame(txPr, self)
        paragraph = text_frame.paragraphs[0]
        return paragraph.font

    @property
    def has_text_frame(self):
        """
        Return |True| if this data label has a text frame (implying it has
        custom data label text), and |False| otherwise. Assigning |True|
        causes a text frame to be added if not already present. Assigning
        |False| causes any existing text frame to be removed along with any
        text contained in the text frame.
        """
        dLbl = self._dLbl
        if dLbl is None:
            return False
        if dLbl.xpath("c:tx/c:rich"):
            return True
        return False

    @has_text_frame.setter
    def has_text_frame(self, value):
        if bool(value) is True:
            self._get_or_add_tx_rich()
        else:
            self._remove_tx_rich()

    @property
    def position(self):
        """
        Read/write :ref:`XlDataLabelPosition` member specifying the position
        of this data label with respect to its data point, or |None| if no
        position is specified. Assigning |None| causes PowerPoint to choose
        the default position, which varies by chart type.
        """
        dLbl = self._dLbl
        if dLbl is None:
            return None
        dLblPos = dLbl.dLblPos
        if dLblPos is None:
            return None
        return dLblPos.val

    @position.setter
    def position(self, value):
        if value is None:
            dLbl = self._dLbl
            if dLbl is None:
                return
            dLbl._remove_dLblPos()
            return
        dLbl = self._get_or_add_dLbl()
        dLbl.get_or_add_dLblPos().val = value

    @property
    def text_frame(self):
        """
        |TextFrame| instance for this data label, containing the text of the
        data label and providing access to its text formatting properties.
        """
        rich = self._get_or_add_rich()
        return TextFrame(rich, self)

    @property
    def _dLbl(self):
        """
        Return the |CT_DLbl| instance referring specifically to this
        individual data label (having the same index value), or |None| if not
        present.
        """
        return self._ser.get_dLbl(self._idx)

    def _get_or_add_dLbl(self):
        """
        The ``CT_DLbl`` instance referring specifically to this individual
        data label, newly created if not yet present in the XML.
        """
        return self._ser.get_or_add_dLbl(self._idx)

    def _get_or_add_rich(self):
        """
        Return the `c:rich` element representing the text frame for this data
        label, newly created with its ancestors if not present.
        """
        dLbl = self._get_or_add_dLbl()

        # having a c:spPr or c:txPr when a c:tx is present causes the "can't
        # save" bug on bubble charts. Remove c:spPr and c:txPr when present.
        dLbl._remove_spPr()
        dLbl._remove_txPr()

        return dLbl.get_or_add_rich()

    def _get_or_add_tx_rich(self):
        """
        Return the `c:tx` element for this data label, with its `c:rich`
        child and descendants, newly created if not yet present.
        """
        dLbl = self._get_or_add_dLbl()

        # having a c:spPr or c:txPr when a c:tx is present causes the "can't
        # save" bug on bubble charts. Remove c:spPr and c:txPr when present.
        dLbl._remove_spPr()
        dLbl._remove_txPr()

        return dLbl.get_or_add_tx_rich()

    def _get_or_add_txPr(self):
        """Return the `c:txPr` element for this data label.

        The `c:txPr` element and its parent `c:dLbl` element are created if
        not yet present.
        """
        dLbl = self._get_or_add_dLbl()
        return dLbl.get_or_add_txPr()

    def _remove_tx_rich(self):
        """
        Remove any `c:tx/c:rich` child of the `c:dLbl` element for this data
        label. Do nothing if that element is not present.
        """
        dLbl = self._dLbl
        if dLbl is None:
            return
        dLbl.remove_tx_rich()
