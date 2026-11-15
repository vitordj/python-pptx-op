"""Data model for chartEx charts.

A chartEx chart has a single series whose values map onto a category dimension. The
category dimension may be *flat* (a list of strings) or *hierarchical* (a list of
tuples, root-first), the latter being used by treemap and sunburst.
"""

from __future__ import annotations

import math
import re
from typing import Iterable, Sequence, Union

CategoryValue = Union[str, Sequence[str]]

_HEX_COLOR = re.compile(r"[0-9A-F]{6}")


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

    # -- a data holder: `__slots__` turns a misspelled attribute into an immediate
    # -- AttributeError instead of a silently-ignored setting.
    __slots__ = (
        "_categories",
        "_series_name",
        "_values",
        "_point_colors",
        "_subtotal_idxs",
        "_number_format",
        "_font_name",
        "_font_size",
    )

    def __init__(self) -> None:
        self._categories: list[CategoryValue] = []
        self._series_name: str | None = None
        self._values: list[float] = []
        self._point_colors: list[str | None] = []
        self._subtotal_idxs: list[int] | None = None
        self._number_format: str | None = None
        self._font_name: str | None = None
        self._font_size: float | None = None

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

        Checks that the category and value counts agree, that every value is finite, that
        subtotal indexes and point colors fit the values, and that hierarchical categories
        all have the same depth (and are not mixed with plain strings). Returns |None| when
        the data is fine; never modifies the data.
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

        if self._subtotal_idxs is not None:
            for idx in self._subtotal_idxs:
                if not 0 <= idx < n_vals:
                    raise ValueError(f"subtotal index {idx} is out of range for {n_vals} values")

        if len(self._point_colors) > n_vals:
            raise ValueError(
                f"{len(self._point_colors)} point colors given for {n_vals} values; "
                "give at most one color per value"
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

    # -- optional formatting ---------------------------------------------------

    @property
    def point_colors(self) -> list[str | None]:
        """Per-data-point fill colors as ``RRGGBB`` hex strings.

        A shorter list (or a ``None`` entry) leaves the remaining points with the
        theme color assigned by the chart color style. Used e.g. to paint a waterfall
        green for positive and red for negative contributions.
        """
        return self._point_colors

    @point_colors.setter
    def point_colors(self, value: Iterable[str | None] | None) -> None:
        if value is None:
            self._point_colors = []
            return
        colors: list[str | None] = []
        for idx, c in enumerate(value):
            if c is None:
                colors.append(None)
                continue
            hex_ = str(c).lstrip("#").upper()
            if not _HEX_COLOR.fullmatch(hex_):
                raise ValueError(
                    f"point color at index {idx} is {c!r}; expected an RGBColor or a"
                    " six-digit hex string such as '92D050' or '#92D050'"
                )
            colors.append(hex_)
        self._point_colors = colors

    @property
    def subtotal_idxs(self) -> list[int] | None:
        """Indexes of the waterfall points drawn as totals (bars starting at zero).

        ``None`` (the default) keeps the convention of treating the first and last
        categories as totals; set it explicitly to mark only some of them — e.g.
        ``[len(values) - 1]`` when only the closing "Total" bar is a total.
        """
        return self._subtotal_idxs

    @subtotal_idxs.setter
    def subtotal_idxs(self, value: Iterable[int] | None) -> None:
        self._subtotal_idxs = None if value is None else sorted({int(i) for i in value})

    @property
    def number_format(self) -> str | None:
        """Number format code for the values, e.g. ``"0.00%"``.

        Applied to the cached values, to the value axis tick labels and to the data
        labels. ``None`` means ``"General"``.
        """
        return self._number_format

    @number_format.setter
    def number_format(self, value: str | None) -> None:
        self._number_format = None if value is None else str(value)

    @property
    def font_name(self) -> str | None:
        """Latin typeface for axis tick labels and data labels (``None`` = inherit)."""
        return self._font_name

    @font_name.setter
    def font_name(self, value: str | None) -> None:
        self._font_name = None if value is None else str(value)

    @property
    def font_size(self) -> float | None:
        """Font size *in points* for axis tick labels and data labels."""
        return self._font_size

    @font_size.setter
    def font_size(self, value: float | None) -> None:
        if value is None:
            self._font_size = None
            return
        size = float(value)
        if not (1.0 <= size <= 4000.0):
            raise ValueError(f"font_size must be between 1 and 4000 points, got {value!r}")
        self._font_size = size

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
