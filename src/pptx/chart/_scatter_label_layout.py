"""Pure scatter label layout helpers (no pptx/lxml/openpyxl imports).

This module is deliberately free of any dependency on the rest of the library so
its clustering and offset math can be unit-tested in isolation. It computes
per-point factorial offsets for scatter data labels: isolated points are left to
PowerPoint's automatic placement (``None``), while clustered points fan out
radially from their cluster centroid.
"""

from __future__ import annotations

import math
from typing import TYPE_CHECKING, Sequence

if TYPE_CHECKING:
    from typing_extensions import TypeAlias

_DEFAULT_CLUSTER_EPS = 0.12
_DEFAULT_MIN_CLUSTER = 2
_DEFAULT_MAG_BASE = 0.036
_DEFAULT_MAG_MAX = 0.135
_COINCIDENT_RADIUS = 0.03
_MIN_ANGLE_SEP = 0.50
_REF_LABEL_LEN = 12

Offset2D: "TypeAlias" = "tuple[float, float]"
OffsetOrAuto: "TypeAlias" = "Offset2D | None"


def _normalize_xy(
    xs: Sequence[float],
    ys: Sequence[float],
) -> tuple[list[float], list[float]]:
    x_arr = [float(v) for v in xs]
    y_arr = [float(v) for v in ys]
    if not x_arr:
        return [], []
    x0, y0 = min(x_arr), min(y_arr)
    x_span = max(max(x_arr) - x0, 1e-12)
    y_span = max(max(y_arr) - y0, 1e-12)
    return [(x - x0) / x_span for x in x_arr], [(y - y0) / y_span for y in y_arr]


def cluster_scatter_points(
    xs: Sequence[float],
    ys: Sequence[float],
    *,
    cluster_eps: float = _DEFAULT_CLUSTER_EPS,
) -> list[list[int]]:
    """Connected components by Euclidean distance in normalized [0,1]²."""
    n = len(xs)
    if n == 0:
        return []

    xn, yn = _normalize_xy(xs, ys)
    seen = [False] * n
    clusters: list[list[int]] = []

    for i in range(n):
        if seen[i]:
            continue
        stack = [i]
        seen[i] = True
        members = [i]
        while stack:
            u = stack.pop()
            for v in range(n):
                if seen[v]:
                    continue
                if math.hypot(xn[u] - xn[v], yn[u] - yn[v]) <= cluster_eps:
                    seen[v] = True
                    stack.append(v)
                    members.append(v)
        clusters.append(members)
    return clusters


def _nearest_in_cluster(
    idx: int,
    members: Sequence[int],
    xn: Sequence[float],
    yn: Sequence[float],
) -> float:
    best = float("inf")
    for j in members:
        if j == idx:
            continue
        best = min(best, math.hypot(xn[idx] - xn[j], yn[idx] - yn[j]))
    return 0.0 if math.isinf(best) else best


def _adaptive_magnitude(
    *,
    dist_centroid: float,
    dist_neighbor: float,
    cluster_size: int,
    label_len: int,
    cluster_eps: float,
    magnitude_base: float,
    magnitude_max: float,
) -> float:
    s_prox = 1.0 / (1.0 + dist_centroid / 0.045)
    s_dens = 1.0 / (1.0 + dist_neighbor / max(cluster_eps * 0.35, 0.02))
    s_size = min(1.0, max(0, cluster_size - 1) / 8.0)
    s_lab = min(1.0, max(0, label_len - 8) / 28.0)
    score = 0.34 * s_prox + 0.30 * s_dens + 0.14 * s_size + 0.22 * s_lab
    shaped = math.log1p(3.0 * score) / math.log1p(3.0)
    return magnitude_base + (magnitude_max - magnitude_base) * shaped


def _spread_angles(
    angles: Sequence[float],
    *,
    min_sep_rad: float = _MIN_ANGLE_SEP,
) -> list[float]:
    """Keep circular order; enforce minimum separation. Do NOT rotate the fan."""
    n = len(angles)
    if n <= 1:
        return list(angles)

    order = sorted(range(n), key=angles.__getitem__)
    sorted_angles = [angles[i] for i in order]
    unwrapped = [sorted_angles[0]]
    for a in sorted_angles[1:]:
        cand = a
        while cand < unwrapped[-1]:
            cand += 2.0 * math.pi
        unwrapped.append(cand)

    span = unwrapped[-1] - unwrapped[0]
    needed = (n - 1) * min_sep_rad
    out = [0.0] * n

    if span + 1e-9 >= needed and span > 1e-9:
        placed = [unwrapped[0]]
        for a in unwrapped[1:]:
            prev = placed[-1]
            placed.append(a if (a - prev) >= min_sep_rad else prev + min_sep_rad)
        for k, i in enumerate(order):
            out[i] = placed[k] % (2.0 * math.pi)
        return out

    s = sum(math.sin(a) for a in angles)
    c = sum(math.cos(a) for a in angles)
    mean = math.atan2(s, c)
    half = 0.5 * max(needed, min_sep_rad)
    for k, i in enumerate(order):
        t = 0.0 if n == 1 else k / (n - 1)
        out[i] = (mean - half + 2.0 * half * t) % (2.0 * math.pi)
    return out


def _bbox_magnitude(
    mag: float,
    ang: float,
    label_len: int,
    *,
    magnitude_max: float,
) -> float:
    half_w = min(0.16, 0.0026 * max(label_len, 1))
    half_h = 0.032 if label_len > 24 else 0.018
    need = abs(math.cos(ang)) * half_w + abs(math.sin(ang)) * half_h
    return min(magnitude_max, max(mag, need))


def _obstacle_clearance(
    ang: float,
    *,
    cx: float,
    cy: float,
    mag: float,
    label_len: int,
    obstacles: Sequence[tuple[float, float, float]],
) -> float:
    if not obstacles:
        return 1.0

    half_w = min(0.14, 0.0024 * max(label_len, 1))
    half_h = 0.028 if label_len > 24 else 0.016
    lx = cx + mag * math.cos(ang)
    ly = cy + mag * math.sin(ang)
    ux, uy = math.cos(ang), math.sin(ang)
    px, py = -uy, ux
    samples = (
        (lx, ly),
        (lx + ux * half_w * 0.5, ly + uy * half_h * 0.5),
        (lx + px * half_w, ly + py * half_h),
        (lx - px * half_w, ly - py * half_h),
    )
    best = float("inf")
    for ox, oy, weight in obstacles:
        for sx, sy in samples:
            best = min(best, math.hypot(sx - ox, sy - oy) / max(weight, 0.35))
    return best


def _orient_coincident_fan(
    n: int,
    mags: Sequence[float],
    label_lens: Sequence[int],
    *,
    cx: float,
    cy: float,
    obstacles: Sequence[tuple[float, float, float]],
) -> list[float]:
    """Uniform 2π/n fan for coincident cores; longest labels → most vertical slots."""
    if n <= 0:
        return []

    def assign(slots: Sequence[float]) -> list[float]:
        slot_order = sorted(range(n), key=lambda i: abs(math.sin(slots[i])), reverse=True)
        label_order = sorted(range(n), key=lambda i: label_lens[i], reverse=True)
        out = [0.0] * n
        for slot_i, label_i in zip(slot_order, label_order):
            out[label_i] = slots[slot_i]
        return out

    base = [2.0 * math.pi * k / n for k in range(n)]
    seed = math.pi / (2.0 * n)
    best = assign([(b + seed) % (2.0 * math.pi) for b in base])
    best_score: tuple[float, float, float] | None = None

    for step in range(24):
        rot = seed + step * (2.0 * math.pi / 24.0)
        angs = assign([(b + rot) % (2.0 * math.pi) for b in base])
        horiz_pen = sum(
            (max(0.0, label_lens[i] - 16) / 36.0) * (abs(math.cos(angs[i])) ** 2) for i in range(n)
        )
        long_verts = [abs(math.sin(angs[i])) for i in range(n) if label_lens[i] >= 20]
        min_vert = min(long_verts) if long_verts else min(abs(math.sin(a)) for a in angs)
        clear = (
            min(
                _obstacle_clearance(
                    angs[i],
                    cx=cx,
                    cy=cy,
                    mag=mags[i],
                    label_len=label_lens[i],
                    obstacles=obstacles,
                )
                for i in range(n)
            )
            if obstacles
            else 1.0
        )
        score = (min_vert, -horiz_pen, clear)
        if best_score is None or score > best_score:
            best_score = score
            best = angs
    return best


def hybrid_cluster_offsets(
    xs: Sequence[float],
    ys: Sequence[float],
    labels: Sequence[str] | None = None,
    *,
    cluster_eps: float = _DEFAULT_CLUSTER_EPS,
    min_cluster_size: int = _DEFAULT_MIN_CLUSTER,
    magnitude_base: float = _DEFAULT_MAG_BASE,
    magnitude_max: float = _DEFAULT_MAG_MAX,
    magnitude: float | None = None,
    avoid_outsiders: bool = True,
) -> list[OffsetOrAuto]:
    """Hybrid factorial offsets: radial only on clusters; isolates → None.

    * Elongated / chain clusters: direction = point − centroid (do not rotate fan).
    * Near-coincident cores: uniform fan; long labels on vertical slots.
    * Returned Y is already flipped for PowerPoint (positive = downward).
    """
    n = len(xs)
    offsets: list[OffsetOrAuto] = [None] * n
    if n == 0:
        return offsets

    if magnitude is not None:
        magnitude_base = magnitude_max = float(magnitude)

    if labels is not None and len(labels) != n:
        raise ValueError("labels must have the same length as xs/ys.")
    label_lens = [len(str(labels[i])) if labels is not None else _REF_LABEL_LEN for i in range(n)]

    xn, yn = _normalize_xy(xs, ys)

    for members in cluster_scatter_points(xs, ys, cluster_eps=cluster_eps):
        if len(members) < min_cluster_size:
            continue

        cx = sum(xn[i] for i in members) / len(members)
        cy = sum(yn[i] for i in members) / len(members)
        member_set = set(members)

        obstacles: list[tuple[float, float, float]] = []
        if avoid_outsiders:
            for j in range(n):
                if j in member_set:
                    continue
                weight = 0.55 + min(0.7, label_lens[j] / 40.0)
                obstacles.append((xn[j], yn[j], weight))

        angles: list[float] = []
        dist_c: list[float] = []
        dist_nn: list[float] = []
        for local_i, i in enumerate(members):
            dx, dy = xn[i] - cx, yn[i] - cy
            r = math.hypot(dx, dy)
            dist_c.append(r)
            dist_nn.append(_nearest_in_cluster(i, members, xn, yn))
            if r < 1e-9:
                angles.append((2.0 * math.pi * local_i) / len(members))
            else:
                angles.append(math.atan2(dy, dx))

        coincident = max(dist_c) < _COINCIDENT_RADIUS
        lens_local = [label_lens[i] for i in members]
        mags = [
            _adaptive_magnitude(
                dist_centroid=dist_c[k],
                dist_neighbor=dist_nn[k],
                cluster_size=len(members),
                label_len=lens_local[k],
                cluster_eps=cluster_eps,
                magnitude_base=magnitude_base,
                magnitude_max=magnitude_max,
            )
            for k in range(len(members))
        ]

        if coincident:
            angles = _orient_coincident_fan(
                len(members),
                mags,
                lens_local,
                cx=cx,
                cy=cy,
                obstacles=obstacles,
            )
        else:
            angles = _spread_angles(angles)

        for local_i, i in enumerate(members):
            ang = angles[local_i]
            mag = _bbox_magnitude(
                mags[local_i], ang, lens_local[local_i], magnitude_max=magnitude_max
            )
            # PowerPoint layout Y grows downward.
            offsets[i] = (mag * math.cos(ang), -mag * math.sin(ang))

    return offsets
