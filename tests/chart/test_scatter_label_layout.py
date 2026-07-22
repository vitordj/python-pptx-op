"""Unit tests for the pure pptx.chart._scatter_label_layout module."""

import math

import pytest

from pptx.chart._scatter_label_layout import (
    cluster_scatter_points,
    hybrid_cluster_offsets,
)

# --- fixtures (embedded, no PPTX required) ----------------------------------

LABELS_PORTFOLIO = [
    "Itau BOVV11 FIC de FIA",
    "Ibovespa",
    "4UM Marlim Dividendos FIA",
    "Fund Core FIA",
    "Fund Core FIA I",
    "Fund Core FIC FIA",
    "Atmos Institucional FIC FIA",
    "Fund Core FIA II",
    "Tarpon GT",
]
XS_PORTFOLIO = [0.199, 0.192, 0.183, 0.197, 0.209, 0.196, 0.225, 0.246, 0.240]
YS_PORTFOLIO = [0.014, 0.068, 0.060, 0.071, 0.046, 0.022, 0.009, -0.011, -0.110]

LABELS_LONG = [
    "Fundo Institucional de Ações Dividendos Premium FIC FIA",
    "Fundo Institucional de Ações Value Long Bias FIC FIA",
    "Fundo Macro Multimercado Crédito Privado FIC FIM",
    "Ibovespa",
    "CDI",
]
XS_LONG = [0.198, 0.1985, 0.199, 0.240, 0.185]
YS_LONG = [0.040, 0.041, 0.0395, -0.080, 0.070]

LABELS_CHAIN = [
    "Elo A",
    "Elo B intermediário",
    "Elo C intermediário longo",
    "Elo D",
    "Isolado Curto",
]
XS_CHAIN = [0.195, 0.198, 0.201, 0.204, 0.245]
YS_CHAIN = [0.028, 0.031, 0.034, 0.037, -0.090]


def _data_angle(offset):
    """Recover the data-space angle from a PPT offset (undo the Y flip)."""
    ox, oy = offset
    return math.atan2(-oy, ox)


class DescribeHybridClusterOffsets:
    def it_returns_empty_for_no_points(self):
        assert hybrid_cluster_offsets([], []) == []

    def it_leaves_isolated_points_on_automatic_layout(self):
        offsets = hybrid_cluster_offsets(XS_PORTFOLIO, YS_PORTFOLIO, LABELS_PORTFOLIO)
        active = [o for o in offsets if o is not None]
        assert 2 <= len(active) <= 6
        assert any(o is None for o in offsets)

    def it_keeps_chain_members_on_their_centroid_side(self):
        offsets = hybrid_cluster_offsets(XS_CHAIN, YS_CHAIN, LABELS_CHAIN)
        assert offsets[4] is None
        assert offsets[0][0] < 0
        assert offsets[1][0] < 0
        assert offsets[2][0] > 0
        assert offsets[3][0] > 0

    def it_sends_the_longest_coincident_label_most_vertical(self):
        offsets = hybrid_cluster_offsets(XS_LONG, YS_LONG, LABELS_LONG)
        active_idx = [i for i, o in enumerate(offsets) if o is not None]
        assert len(active_idx) == 3
        sins = {i: abs(math.sin(_data_angle(offsets[i]))) for i in active_idx}
        longest = max(active_idx, key=lambda i: len(LABELS_LONG[i]))
        assert sins[longest] == max(sins.values())

    def it_honors_a_fixed_magnitude(self):
        offsets = hybrid_cluster_offsets(
            XS_PORTFOLIO, YS_PORTFOLIO, LABELS_PORTFOLIO, magnitude=0.065
        )
        mags = {round(math.hypot(*o), 6) for o in offsets if o is not None}
        assert mags == {0.065}

    def it_offsets_every_point_in_a_collapsed_cluster(self):
        labels = ["Point %d" % i for i in range(5)]
        offsets = hybrid_cluster_offsets([0.205] * 5, [0.030] * 5, labels)
        assert all(o is not None for o in offsets)
        angles = sorted(_data_angle(o) % (2 * math.pi) for o in offsets)
        gaps = [b - a for a, b in zip(angles, angles[1:])]
        assert all(g > 0.5 for g in gaps)  # roughly 2π/5 apart, none collapsed

    def it_raises_when_labels_length_mismatches(self):
        with pytest.raises(ValueError):
            hybrid_cluster_offsets([0.1, 0.2], [0.1, 0.2], ["only one"])

    def it_flips_y_for_powerpoint(self):
        # in the elongated chain, the point below the centroid is pushed down
        # (oy > 0) and the one above is pushed up (oy < 0), per PPT's Y axis.
        offsets = hybrid_cluster_offsets(XS_CHAIN, YS_CHAIN, LABELS_CHAIN)
        assert offsets[0][1] > 0  # lowest Y → positive (downward) offset
        assert offsets[3][1] < 0  # highest Y → negative (upward) offset


class DescribeClusterScatterPoints:
    def it_groups_near_points_and_splits_far_ones(self):
        clusters = cluster_scatter_points([0.10, 0.11, 0.90], [0.10, 0.11, 0.90])
        sizes = sorted(len(c) for c in clusters)
        assert sizes == [1, 2]

    def it_returns_empty_for_no_points(self):
        assert cluster_scatter_points([], []) == []
