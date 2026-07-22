"""Chart data-label related oxml objects."""

from __future__ import annotations

from uuid import uuid4
from xml.sax.saxutils import escape

from pptx.enum.chart import XL_DATA_LABEL_POSITION
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.text import CT_TextBody
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

# -- URIs of the Office 2013 (`c15`) chart extensions used for cell-linked
#    scatter data labels; the `c15` namespace is declared in `ns.py`. The
#    cache-only design tradeoff is documented on `set_datalabels_range`. --
_LABEL_RANGE_URI = "{02D57815-91ED-43cb-92C2-25804820EDAC}"
_SHOW_RANGE_URI = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}"


def replace_ext(extLst, uri, *children):
    """Replace the `c:ext` having *uri* under *extLst* with a fresh one.

    Any `c:ext` whose `uri` differs is left untouched, so unknown extensions
    survive a round-trip. The new `c:ext` receives *children* in order.
    """
    for ext in extLst.findall(qn("c:ext")):
        if ext.get("uri") == uri:
            extLst.remove(ext)
    ext = parse_xml('<c:ext %s uri="%s"/>' % (nsdecls("c"), uri))
    for child in children:
        ext.append(child)
    extLst.append(ext)
    return ext


class CT_DLbl(BaseOxmlElement):
    """
    ``<c:dLbl>`` element specifying the properties of the data label for an
    individual data point.
    """

    _tag_seq = (
        "c:idx",
        "c:layout",
        "c:tx",
        "c:numFmt",
        "c:spPr",
        "c:txPr",
        "c:dLblPos",
        "c:showLegendKey",
        "c:showVal",
        "c:showCatName",
        "c:showSerName",
        "c:showPercent",
        "c:showBubbleSize",
        "c:separator",
        "c:extLst",
    )
    idx = OneAndOnlyOne("c:idx")
    tx = ZeroOrOne("c:tx", successors=_tag_seq[3:])
    spPr = ZeroOrOne("c:spPr", successors=_tag_seq[5:])
    txPr = ZeroOrOne("c:txPr", successors=_tag_seq[6:])
    dLblPos = ZeroOrOne("c:dLblPos", successors=_tag_seq[7:])
    extLst = ZeroOrOne("c:extLst", successors=())
    del _tag_seq

    def get_or_add_rich(self):
        """
        Return the `c:rich` descendant representing the text frame of the
        data label, newly created if not present. Any existing `c:strRef`
        element is removed along with its contents.
        """
        tx = self.get_or_add_tx()
        tx._remove_strRef()
        return tx.get_or_add_rich()

    def get_or_add_tx_rich(self):
        """
        Return the `c:tx[c:rich]` subtree, newly created if not present.
        """
        tx = self.get_or_add_tx()
        tx._remove_strRef()
        tx.get_or_add_rich()
        return tx

    @property
    def idx_val(self):
        """
        The integer value of the `val` attribute on the required `c:idx`
        child.
        """
        return self.idx.val

    @classmethod
    def new_dLbl(cls):
        """Return a newly created "loose" `c:dLbl` element.

        The `c:dLbl` element contains the same (fairly extensive) default
        subtree added by PowerPoint when an individual data label is
        customized in the UI. Note that the idx value must be set by the
        client. Failure to set the idx value will likely result in any
        changes not being visible and may result in a repair error on open.
        """
        return parse_xml(
            "<c:dLbl %s>\n"
            '  <c:idx val="666"/>\n'
            "  <c:spPr/>\n"
            "  <c:txPr>\n"
            "    <a:bodyPr/>\n"
            "    <a:lstStyle/>\n"
            "    <a:p>\n"
            "      <a:pPr>\n"
            "        <a:defRPr/>\n"
            "      </a:pPr>\n"
            "    </a:p>\n"
            "  </c:txPr>\n"
            '  <c:showLegendKey val="0"/>\n'
            '  <c:showVal val="1"/>\n'
            '  <c:showCatName val="0"/>\n'
            '  <c:showSerName val="0"/>\n'
            '  <c:showPercent val="0"/>\n'
            '  <c:showBubbleSize val="0"/>\n'
            "</c:dLbl>" % nsdecls("c", "a")
        )

    @classmethod
    def new_cell_range_dLbl(cls, idx, x, y, text):
        """Return a loose `c:dLbl` for a displaced cell-linked scatter label.

        The label is manually positioned at factorial offset (*x*, *y*) and its
        text comes from a `CELLRANGE` field (linked to the worksheet), matching
        what PowerPoint writes when a data label is dragged. The caller is
        responsible for adding the `c15:showDataLabelsRange` extension.
        """
        return parse_xml(
            "<c:dLbl %s>\n"
            '  <c:idx val="%d"/>\n'
            "  <c:layout>\n"
            "    <c:manualLayout>\n"
            '      <c:x val="%s"/>\n'
            '      <c:y val="%s"/>\n'
            "    </c:manualLayout>\n"
            "  </c:layout>\n"
            "  <c:tx>\n"
            "    <c:rich>\n"
            "      <a:bodyPr/>\n"
            "      <a:lstStyle/>\n"
            "      <a:p>\n"
            '        <a:fld id="{%s}" type="CELLRANGE">\n'
            '          <a:rPr lang="en-US"/>\n'
            "          <a:pPr/>\n"
            "          <a:t>%s</a:t>\n"
            "        </a:fld>\n"
            '        <a:endParaRPr lang="en-US"/>\n'
            "      </a:p>\n"
            "    </c:rich>\n"
            "  </c:tx>\n"
            '  <c:showLegendKey val="0"/>\n'
            '  <c:showVal val="0"/>\n'
            '  <c:showCatName val="0"/>\n'
            '  <c:showSerName val="0"/>\n'
            '  <c:showPercent val="0"/>\n'
            '  <c:showBubbleSize val="0"/>\n'
            "</c:dLbl>" % (nsdecls("c", "a"), idx, x, y, str(uuid4()).upper(), escape(str(text)))
        )

    def remove_tx_rich(self):
        """
        Remove any `c:tx[c:rich]` child, or do nothing if not present.
        """
        matches = self.xpath("c:tx[c:rich]")
        if not matches:
            return
        tx = matches[0]
        self.remove(tx)

    def _new_txPr(self):
        return CT_TextBody.new_txPr()


class CT_DLblPos(BaseOxmlElement):
    """
    ``<c:dLblPos>`` element specifying the positioning of a data label with
    respect to its data point.
    """

    val = RequiredAttribute("val", XL_DATA_LABEL_POSITION)


class CT_DLbls(BaseOxmlElement):
    """`c:dLbls` element specifying properties for a set of data labels."""

    _tag_seq = (
        "c:dLbl",
        "c:numFmt",
        "c:spPr",
        "c:txPr",
        "c:dLblPos",
        "c:showLegendKey",
        "c:showVal",
        "c:showCatName",
        "c:showSerName",
        "c:showPercent",
        "c:showBubbleSize",
        "c:separator",
        "c:showLeaderLines",
        "c:leaderLines",
        "c:extLst",
    )
    dLbl = ZeroOrMore("c:dLbl", successors=_tag_seq[1:])
    numFmt = ZeroOrOne("c:numFmt", successors=_tag_seq[2:])
    txPr = ZeroOrOne("c:txPr", successors=_tag_seq[4:])
    dLblPos = ZeroOrOne("c:dLblPos", successors=_tag_seq[5:])
    showLegendKey = ZeroOrOne("c:showLegendKey", successors=_tag_seq[6:])
    showVal = ZeroOrOne("c:showVal", successors=_tag_seq[7:])
    showCatName = ZeroOrOne("c:showCatName", successors=_tag_seq[8:])
    showSerName = ZeroOrOne("c:showSerName", successors=_tag_seq[9:])
    showPercent = ZeroOrOne("c:showPercent", successors=_tag_seq[10:])
    showBubbleSize = ZeroOrOne("c:showBubbleSize", successors=_tag_seq[11:])
    showLeaderLines = ZeroOrOne("c:showLeaderLines", successors=_tag_seq[13:])
    extLst = ZeroOrOne("c:extLst", successors=())
    del _tag_seq

    @property
    def defRPr(self):
        """
        ``<a:defRPr>`` great-great-grandchild element, added with its
        ancestors if not present.
        """
        txPr = self.get_or_add_txPr()
        defRPr = txPr.defRPr
        return defRPr

    def get_dLbl_for_point(self, idx):
        """
        Return the `c:dLbl` child representing the label for the data point
        at index *idx*.
        """
        matches = self.xpath('c:dLbl[c:idx[@val="%d"]]' % idx)
        if matches:
            return matches[0]
        return None

    def get_or_add_dLbl_for_point(self, idx):
        """
        Return the `c:dLbl` element representing the label of the point at
        index *idx*.
        """
        matches = self.xpath('c:dLbl[c:idx[@val="%d"]]' % idx)
        if matches:
            return matches[0]
        return self._insert_dLbl_in_sequence(idx)

    def add_cell_range_dLbl(self, idx, x, y, text):
        """Add a displaced cell-linked `c:dLbl` for point *idx*, in `c:idx` order.

        The new `c:dLbl` carries a `manualLayout` at factorial offset (*x*, *y*),
        a `CELLRANGE` field showing *text*, and the `c15:dlblFieldTable` /
        `c15:showDataLabelsRange` extension PowerPoint expects on such a label.
        """
        dLbl = CT_DLbl.new_cell_range_dLbl(idx, x, y, text)
        replace_ext(
            dLbl.get_or_add_extLst(),
            _SHOW_RANGE_URI,
            parse_xml("<c15:dlblFieldTable %s/>" % nsdecls("c15")),
            parse_xml('<c15:showDataLabelsRange %s val="1"/>' % nsdecls("c15")),
        )
        siblings = self.dLbl_lst
        for existing in siblings:
            if existing.idx_val > idx:
                existing.addprevious(dLbl)
                return dLbl
        if siblings:
            siblings[-1].addnext(dLbl)
        else:
            self.insert(0, dLbl)
        return dLbl

    def use_cell_range_labels(self, show_leader_lines=True):
        """Configure this `c:dLbls` for cell-linked scatter labels.

        Turns off every built-in label content, sets leader lines per
        *show_leader_lines*, and installs the `c15:showDataLabelsRange` /
        `c15:showLeaderLines` extension. Existing per-point `c:dLbl` children
        are removed so the caller starts from a clean slate.
        """
        for dLbl in self.dLbl_lst:
            self.remove(dLbl)
        self.get_or_add_showLegendKey().val = False
        self.get_or_add_showVal().val = False
        self.get_or_add_showCatName().val = False
        self.get_or_add_showSerName().val = False
        self.get_or_add_showPercent().val = False
        self.get_or_add_showBubbleSize().val = False
        self.get_or_add_showLeaderLines().val = bool(show_leader_lines)
        leader = "1" if show_leader_lines else "0"
        replace_ext(
            self.get_or_add_extLst(),
            _SHOW_RANGE_URI,
            parse_xml('<c15:showDataLabelsRange %s val="1"/>' % nsdecls("c15")),
            parse_xml('<c15:showLeaderLines %s val="%s"/>' % (nsdecls("c15"), leader)),
        )

    @classmethod
    def new_dLbls(cls):
        """Return a newly created "loose" `c:dLbls` element."""
        return parse_xml(
            "<c:dLbls %s>\n"
            '  <c:showLegendKey val="0"/>\n'
            '  <c:showVal val="0"/>\n'
            '  <c:showCatName val="0"/>\n'
            '  <c:showSerName val="0"/>\n'
            '  <c:showPercent val="0"/>\n'
            '  <c:showBubbleSize val="0"/>\n'
            '  <c:showLeaderLines val="1"/>\n'
            "</c:dLbls>" % nsdecls("c")
        )

    def _insert_dLbl_in_sequence(self, idx):
        """
        Return a newly created `c:dLbl` element having `c:idx` child of *idx*
        and inserted in numeric sequence among the `c:dLbl` children of this
        element.
        """
        new_dLbl = self._new_dLbl()
        new_dLbl.idx.val = idx

        dLbl = None
        for dLbl in self.dLbl_lst:
            if dLbl.idx_val > idx:
                dLbl.addprevious(new_dLbl)
                return new_dLbl
        if dLbl is not None:
            dLbl.addnext(new_dLbl)
        else:
            self.insert(0, new_dLbl)
        return new_dLbl

    def _new_dLbl(self):
        return CT_DLbl.new_dLbl()

    def _new_showCatName(self):
        """Return a new `c:showCatName` with value initialized.

        This method is called by the metaclass-generated code whenever a new
        `c:showCatName` element is required. In this case, it defaults to
        `val=true`, which is not what we need so we override to make val
        explicitly False.
        """
        return parse_xml('<c:showCatName %s val="0"/>' % nsdecls("c"))

    def _new_showLegendKey(self):
        return parse_xml('<c:showLegendKey %s val="0"/>' % nsdecls("c"))

    def _new_showPercent(self):
        return parse_xml('<c:showPercent %s val="0"/>' % nsdecls("c"))

    def _new_showSerName(self):
        return parse_xml('<c:showSerName %s val="0"/>' % nsdecls("c"))

    def _new_showVal(self):
        return parse_xml('<c:showVal %s val="0"/>' % nsdecls("c"))

    def _new_txPr(self):
        return CT_TextBody.new_txPr()
