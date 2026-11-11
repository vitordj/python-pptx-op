"""Data model for chartEx charts.

A chartEx chart has a single series whose values map onto a category dimension. The
category dimension may be *flat* (a list of strings) or *hierarchical* (a list of
tuples, root-first), the latter being used by treemap and sunburst.
"""

from __future__ import annotations

import math
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

    def validate(self) -> None:
        """Raise |ValueError| if this data cannot be written as a valid chartEx part.

        Checks that the category and value counts agree, that every value is finite, and
        that hierarchical categories all have the same depth (and are not mixed with plain
        strings). Returns |None| when the data is fine; never modifies the data.
        """
        n_cats, n_vals = len(self._categories), len(self._values)
        if n_cats != n_vals:
            raise ValueError(
                f"chartEx data has {n_cats} categories but {n_vals} values; "
                "they must be the same length"
            )

        for idx, value in enumerate(self._values):
            if not math.isfinite(value):
                raise ValueError(
                    f"chartEx value at index {idx} is {value!r}; "
                    "values must be finite (no nan or inf)"
                )

        if not self._categories:
            return
        first_is_str = isinstance(self._categories[0], str)
        for idx, cat in enumerate(self._categories):
            if isinstance(cat, str) != first_is_str:
                raise ValueError(
                    f"chartEx categories mix plain strings and tuples (category {idx} "
                    f"is {'a string' if isinstance(cat, str) else 'a tuple'}, category 0 "
                    f"is {'a string' if first_is_str else 'a tuple'}); use one or the other"
                )
        if first_is_str:
            return
        first_depth = len(self._categories[0])
        for idx, cat in enumerate(self._categories):
            if len(cat) != first_depth:
                raise ValueError(
                    f"chartEx category {idx} has depth {len(cat)} but the first category "
                    f"has depth {first_depth}; all hierarchical categories must have the "
                    "same depth"
                )

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
