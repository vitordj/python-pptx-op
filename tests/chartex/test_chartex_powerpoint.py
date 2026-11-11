"""Check that real PowerPoint opens and recognises python-pptx chartEx charts.

Tests that run without PowerPoint can only prove the generated XML is well-formed and survives
a save/reload inside python-pptx. This module asks the question that matters: does PowerPoint
itself open the file (without repairing it) and report each chart as the expected chart type?

Requirements: Windows, PowerPoint 2016 or later, and `pywin32` (`pip install pywin32`). On any
other machine every test is skipped with a reason.

Run with::

    pytest tests/chartex/test_chartex_powerpoint.py -v -rs

Optional environment variable `PPTX_OP_CHARTEX_OUT`: set it to a directory (created if missing)
and the generated .pptx files and the exported slide .png files are written there so they can
be inspected visually. Otherwise pytest's temporary directories are used.

Each chart type gets one test per stage (opens, one shape, HasChart, ChartType, PNG export) so a
failure shows exactly which stage failed for which chart type. The PNG stage only checks that
PowerPoint drew something other than an empty slide; whether the chart looks right is for a
person to judge from the exported images.
"""

from __future__ import annotations

import gc
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Iterator, Optional

import pytest
from PIL import Image

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.chartex.data import ChartExData
from pptx.enum.chart import XL_CHART_EX_TYPE, XL_CHART_TYPE
from pptx.util import Inches

pytestmark = pytest.mark.skipif(
    sys.platform != "win32",
    reason="needs Windows with PowerPoint 2016+ and pywin32 (COM automation)",
)

_OUT_ENV_VAR = "PPTX_OP_CHARTEX_OUT"
_MSO_TRUE = -1

# -- XlChartType values PowerPoint reports for each chartEx layout --
# waterfall=xlWaterfall(119), funnel=xlFunnel(123), treemap=xlTreemap(117),
# sunburst=xlSunburst(120), boxWhisker=xlBoxwhisker(121),
# clusteredColumn (histogram layout)=xlHistogram(118)
_EXPECTED_CHART_TYPE = {
    "waterfall": 119,
    "funnel": 123,
    "treemap": 117,
    "sunburst": 120,
    "box-whisker": 121,
    "histogram": 118,
}


# ---------------------------------------------------------------------------------------------
# deck building (pure python-pptx; importable and callable without PowerPoint)
# ---------------------------------------------------------------------------------------------


def _waterfall_data() -> ChartExData:
    data = ChartExData()
    # -- first and last are totals; "Custos" is a negative step --
    data.categories = ["Receita inicial", "Vendas", "Servicos", "Custos", "Impostos", "Total"]
    data.add_series("Resultado", (120.0, 45.0, 20.0, -35.0, -18.0, 132.0))
    return data


def _funnel_data() -> ChartExData:
    data = ChartExData()
    data.categories = ["Visitantes", "Leads", "Qualificados", "Propostas", "Clientes"]
    data.add_series("Pessoas", (5000, 2400, 1100, 450, 160))
    return data


def _box_whisker_data() -> ChartExData:
    data = ChartExData()
    data.categories = ["A"] * 8 + ["B"] * 8
    data.add_series("Medidas", (4, 6, 7, 7, 8, 9, 11, 15, 10, 12, 13, 13, 14, 16, 18, 24))
    return data


def _histogram_data() -> ChartExData:
    data = ChartExData()
    values = (1, 2, 2, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 6, 6, 7, 8, 9, 9, 12)
    data.categories = [f"Obs {i + 1}" for i in range(len(values))]
    data.add_series("Frequencia", values)
    return data


def _treemap_data() -> ChartExData:
    data = ChartExData()
    data.categories = [
        ("Renda Fixa", "Pos-fixado"),
        ("Renda Fixa", "Prefixado"),
        ("Renda Fixa", "Inflacao"),
        ("Acoes", "Brasil"),
        ("Acoes", "Exterior"),
        ("Multimercado", "Macro"),
        ("Multimercado", "Quant"),
    ]
    data.add_series("Patrimonio", (40, 15, 20, 30, 12, 18, 9))
    return data


def _sunburst_data() -> ChartExData:
    data = ChartExData()
    data.categories = [
        ("America", "Brasil"),
        ("America", "EUA"),
        ("America", "Canada"),
        ("Europa", "Alemanha"),
        ("Europa", "Franca"),
        ("Asia", "Japao"),
        ("Asia", "India"),
    ]
    data.add_series("Vendas", (50, 80, 20, 35, 25, 40, 30))
    return data


# (id, chart type, data factory)
_CASES: list[tuple[str, XL_CHART_EX_TYPE, Callable[[], ChartExData]]] = [
    ("waterfall", XL_CHART_EX_TYPE.WATERFALL, _waterfall_data),
    ("funnel", XL_CHART_EX_TYPE.FUNNEL, _funnel_data),
    ("treemap", XL_CHART_EX_TYPE.TREEMAP, _treemap_data),
    ("sunburst", XL_CHART_EX_TYPE.SUNBURST, _sunburst_data),
    ("box-whisker", XL_CHART_EX_TYPE.BOX_WHISKER, _box_whisker_data),
    ("histogram", XL_CHART_EX_TYPE.CLUSTERED_COLUMN, _histogram_data),
]
_CASE_IDS = [case[0] for case in _CASES]


def _blank_slide(prs: Any) -> Any:
    return prs.slides.add_slide(prs.slide_layouts[6])


def _build_chartex_deck(path: Path, chart_type: XL_CHART_EX_TYPE, data: ChartExData) -> Path:
    """Save a one-slide deck holding a single chartEx chart of `chart_type` at `path`."""
    prs = Presentation()
    slide = _blank_slide(prs)
    slide.shapes.add_chartex(chart_type, Inches(1), Inches(1), Inches(8), Inches(5), data)
    prs.save(str(path))
    return path


def _build_mixed_deck(path: Path) -> Path:
    """Save one slide with two chartEx charts and one classic clustered-column chart."""
    prs = Presentation()
    slide = _blank_slide(prs)
    slide.shapes.add_chartex(
        XL_CHART_EX_TYPE.WATERFALL,
        Inches(0.25),
        Inches(0.25),
        Inches(4.5),
        Inches(3.5),
        _waterfall_data(),
    )
    slide.shapes.add_chartex(
        XL_CHART_EX_TYPE.FUNNEL,
        Inches(5),
        Inches(0.25),
        Inches(4.5),
        Inches(3.5),
        _funnel_data(),
    )
    classic = CategoryChartData()
    classic.categories = ["T1", "T2", "T3", "T4"]
    classic.add_series("Receita", (10, 14, 9, 17))
    slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.25),
        Inches(4),
        Inches(9.25),
        Inches(3.25),
        classic,
    )
    prs.save(str(path))
    return path


# ---------------------------------------------------------------------------------------------
# PowerPoint plumbing (COM; only reached on Windows)
# ---------------------------------------------------------------------------------------------


@pytest.fixture(scope="session")
def out_root(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Directory receiving every generated .pptx and .png."""
    override = os.environ.get(_OUT_ENV_VAR)
    if override:
        root = Path(override)
        root.mkdir(parents=True, exist_ok=True)
        return root.resolve()
    return tmp_path_factory.mktemp("chartex_powerpoint").resolve()


@pytest.fixture(scope="session")
def powerpoint_app() -> Iterator[Any]:
    """A PowerPoint.Application COM object, started once for the whole session.

    `Dispatch` is used rather than `DispatchEx`: PowerPoint is a single-instance application,
    so `DispatchEx` would attach to the running instance anyway, and `Dispatch` makes it clear
    that we may be sharing an instance the user already had open. That is why `Quit()` is only
    called when no presentations were open when we started.
    """
    pythoncom = pytest.importorskip("pythoncom", reason="pywin32 is not installed")
    client = pytest.importorskip("win32com.client", reason="pywin32 is not installed")

    pythoncom.CoInitialize()
    app: Any = None
    quit_on_exit = False
    try:
        try:
            app = client.Dispatch("PowerPoint.Application")
            quit_on_exit = app.Presentations.Count == 0
        except Exception as exc:  # -- pywintypes.com_error and friends --
            pytest.skip(f"PowerPoint could not be started via COM: {exc}")
        yield app
    finally:
        try:
            if app is not None and quit_on_exit:
                app.Quit()
        except Exception:
            pass
        app = None
        gc.collect()
        pythoncom.CoUninitialize()


@dataclass
class _OpenedDeck:
    """Outcome of trying to open one generated deck in PowerPoint."""

    name: str
    pptx_path: Path
    png_path: Path
    presentation: Optional[Any] = None
    open_error: Optional[str] = None


def _open_deck(app: Any, pptx_path: Path) -> _OpenedDeck:
    deck = _OpenedDeck(
        name=pptx_path.stem, pptx_path=pptx_path, png_path=pptx_path.with_suffix(".png")
    )
    try:
        # Open2007(FileName, ReadOnly, Untitled, WithWindow, OpenAndRepair), passed positionally;
        # OpenAndRepair=msoFalse so a damaged file raises instead of being silently repaired.
        deck.presentation = app.Presentations.Open2007(str(pptx_path), -1, 0, 0, 0)
    except Exception as exc:
        deck.open_error = f"{type(exc).__name__}: {exc}"
    return deck


def _close_deck(deck: _OpenedDeck) -> None:
    presentation, deck.presentation = deck.presentation, None
    if presentation is not None:
        try:
            presentation.Saved = _MSO_TRUE
            presentation.Close()
        except Exception:
            pass


def _png_is_blank(path: Path) -> bool:
    """True when every pixel of the image at `path` has the same colour."""
    with Image.open(str(path)) as image:
        extrema = image.convert("RGB").getextrema()
    return all(low == high for low, high in extrema)


def _require_presentation(deck: _OpenedDeck) -> Any:
    """Return the open presentation, or fail with the earlier COM open error."""
    if deck.presentation is None:
        pytest.fail(f"{deck.name}: deck did not open, so this stage cannot run: {deck.open_error}")
    return deck.presentation


@pytest.fixture(scope="module", params=_CASES, ids=_CASE_IDS)
def opened_deck(
    request: pytest.FixtureRequest, powerpoint_app: Any, out_root: Path
) -> Iterator[tuple[str, _OpenedDeck]]:
    case_id, chart_type, make_data = request.param
    pptx_path = _build_chartex_deck(out_root / f"chartex_{case_id}.pptx", chart_type, make_data())
    deck = _open_deck(powerpoint_app, pptx_path)
    try:
        yield case_id, deck
    finally:
        _close_deck(deck)


# ---------------------------------------------------------------------------------------------
# per-chart-type stages
# ---------------------------------------------------------------------------------------------


def test_opens_without_com_error(opened_deck: tuple[str, _OpenedDeck]) -> None:
    case_id, deck = opened_deck
    assert deck.open_error is None, (
        f"{case_id}: PowerPoint failed to open {deck.pptx_path} (OpenAndRepair off): "
        f"{deck.open_error}"
    )


def test_slide_has_exactly_one_shape(opened_deck: tuple[str, _OpenedDeck]) -> None:
    case_id, deck = opened_deck
    slide = _require_presentation(deck).Slides(1)
    count = slide.Shapes.Count
    assert count == 1, f"{case_id}: slide 1 reports {count} shapes, expected 1"


def test_shape_has_chart(opened_deck: tuple[str, _OpenedDeck]) -> None:
    case_id, deck = opened_deck
    shape = _require_presentation(deck).Slides(1).Shapes(1)
    assert shape.HasChart == _MSO_TRUE, (
        f"{case_id}: shape.HasChart is {shape.HasChart!r}, expected msoTrue (-1); PowerPoint did"
        f" not recognise the graphic frame as a chart (shape.Type is {shape.Type!r}, msoChart"
        " is 3)"
    )


def test_chart_type_is_recognised(opened_deck: tuple[str, _OpenedDeck]) -> None:
    case_id, deck = opened_deck
    shape = _require_presentation(deck).Slides(1).Shapes(1)
    if shape.HasChart != _MSO_TRUE:
        pytest.fail(f"{case_id}: shape.HasChart is false, so ChartType cannot be read")
    expected = _EXPECTED_CHART_TYPE[case_id]
    actual = shape.Chart.ChartType
    assert actual == expected, f"{case_id}: Chart.ChartType is {actual!r}, expected {expected}"


def test_slide_exports_to_png(opened_deck: tuple[str, _OpenedDeck]) -> None:
    case_id, deck = opened_deck
    slide = _require_presentation(deck).Slides(1)
    try:
        slide.Export(str(deck.png_path), "PNG")
    except Exception as exc:
        pytest.fail(f"{case_id}: Slide.Export to PNG raised {type(exc).__name__}: {exc}")
    assert deck.png_path.exists(), f"{case_id}: no PNG written at {deck.png_path}"
    assert deck.png_path.stat().st_size > 0, f"{case_id}: empty PNG at {deck.png_path}"
    assert not _png_is_blank(deck.png_path), (
        f"{case_id}: {deck.png_path} is a single flat colour, so PowerPoint drew nothing on the"
        " slide"
    )


# ---------------------------------------------------------------------------------------------
# two chartEx charts next to a classic chart
# ---------------------------------------------------------------------------------------------


def test_mixed_deck_opens_with_three_charts(powerpoint_app: Any, out_root: Path) -> None:
    """Part names (chartEx1/2, style1/2, colors1/2) must coexist with a classic chart part."""
    pptx_path = _build_mixed_deck(out_root / "chartex_mixed.pptx")
    deck = _open_deck(powerpoint_app, pptx_path)
    try:
        assert deck.open_error is None, (
            f"mixed deck: PowerPoint failed to open {pptx_path} (OpenAndRepair off): "
            f"{deck.open_error}"
        )
        shapes = deck.presentation.Slides(1).Shapes
        assert shapes.Count == 3, f"mixed deck: slide 1 reports {shapes.Count} shapes, expected 3"
        flags = [shapes(i).HasChart for i in range(1, shapes.Count + 1)]
        assert flags == [_MSO_TRUE] * 3, f"mixed deck: HasChart per shape is {flags!r}"
    finally:
        _close_deck(deck)
