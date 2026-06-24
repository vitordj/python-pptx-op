"""Data model for chartEx charts.

A chartEx chart has a single series whose values map onto a category dimension. The
category dimension may be *flat* (a list of strings) or *hierarchical* (a list of
tuples, root-first), the latter being used by treemap and sunburst.
"""

from __future__ import annotations

from typing import Iterable, Sequence, Union

CategoryValue = Union[str, Sequence[str]]


class ChartExData:
    """Categories and a single series of values for a chartEx chart.

    Categories may be flat::

        data = ChartExData()
        data.categories = ["Anterior", "Compra A", "Compra B", "Total"]
        data.add_series("Atribuição", (0.13, 0.02, -0.01, 0.14))

    or hierarchical (root-first tuples), as needed by treemap/sunburst::

        data.categories = [("Branch 1", "Stem 1", "Leaf 1"),
                           ("Branch 1", "Stem 1", "Leaf 2")]
        data.add_series("Tamanho", (22, 12))
    """

    def __init__(self) -> None:
        self._categories: list[CategoryValue] = []
        self._series_name: str | None = None
        self._values: list[float] = []
        # -- optional styling (consumed by the XML writer; None = writer default) --
        #: per-point fill colors as hex strings ("92D050"); index-aligned with values.
        self.point_colors: "list[str] | None" = None
        #: indices marked as totals/subtotals (waterfall). None -> default (first+last).
        self.subtotal_idxs: "list[int] | None" = None
        #: value-axis number format code (e.g. "0.00%").
        self.number_format: "str | None" = None
        #: font applied to axis tick labels and data labels.
        self.font_name: "str | None" = None
        #: font size in points for axis tick labels and data labels.
        self.font_size: "float | None" = None

    @property
    def categories(self) -> list[CategoryValue]:
        return self._categories

    @categories.setter
    def categories(self, value: Iterable[CategoryValue]) -> None:
        self._categories = list(value)

    @property
    def series_name(self) -> str:
        return self._series_name or "Series 1"

    @property
    def values(self) -> list[float]:
        return self._values

    def add_series(self, name: str, values: Iterable[float]) -> None:
        """Define the (single) series of this chart.

        chartEx charts modeled here carry exactly one series; calling this more than
        once replaces the previous definition.
        """
        self._series_name = name
        self._values = [float(v) for v in values]

    # -- internal helpers used by the XML writer ------------------------------

    @property
    def is_hierarchical(self) -> bool:
        """True when categories are tuples/lists rather than plain strings."""
        return bool(self._categories) and not isinstance(self._categories[0], str)

    @property
    def depth(self) -> int:
        """Number of category levels (1 for flat categories)."""
        if not self._categories:
            return 1
        if isinstance(self._categories[0], str):
            return 1
        return len(self._categories[0])

    def levels(self) -> list[list[str]]:
        """Return category levels ordered *finest-first* (leaf level first).

        This matches the chartEx ``cx:lvl`` ordering, where the first ``cx:lvl`` holds
        the leaf labels and subsequent levels are progressively coarser. For flat
        categories this returns a single level.
        """
        if not self._categories:
            return [[]]
        if isinstance(self._categories[0], str):
            return [[str(c) for c in self._categories]]
        depth = self.depth
        # category tuples are root-first; emit finest (last element) first
        return [[str(cat[depth - 1 - lvl]) for cat in self._categories] for lvl in range(depth)]
