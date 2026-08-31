# python-pptx-op

**A drop-in fork of [python-pptx](https://github.com/scanny/python-pptx) that adds the
PowerPoint formatting features the original never shipped.**

`python-pptx-op` tracks upstream **python-pptx 1.0.2** and keeps the same import name and
the same API, so `import pptx` and every line of code you already wrote keep working. On
top of that it adds the things you hit the moment you try to reproduce a real,
brand-styled deck from Python: table cell borders, the modern Office 2016 chart family
(waterfall, treemap, funnel, sunburst, box & whisker), shadows, colour transparency,
arrow heads, strikethrough, text highlight, bullet control, and more.

```python
from pptx import Presentation
from pptx.tablekit import StyledTable, CellStyle, BorderSpec, Edge
from pptx.util import Cm

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])

HEADER = CellStyle(bold=True, color="FFFFFF", fill="62993D", align="center",
                   borders=BorderSpec.box(Edge("FFFFFF", 0.5)))
BODY = CellStyle(align="center", borders=BorderSpec.underline(Edge("D3D3D3", 0.5)))

table = StyledTable.add(slide.shapes, rows=3, cols=3,
                        x=Cm(2), y=Cm(2), cx=Cm(18), cy=Cm(3))
table.fill_rows([["Region", "2025", "2026"],
                 ["North", "1.2M", "1.5M"],
                 ["South", "0.9M", "1.1M"]], header_style=HEADER, body_style=BODY)
table.fit(width_cm=18).autosize_rows(font_pt=12, header_pt=12)

prs.save("report.pptx")
```

---

## Why this fork exists

python-pptx is an excellent library, but it deliberately exposes a conservative slice of
the OOXML presentation format. A number of formatting controls that a designer takes for
granted have been requested upstream for years and are still unimplemented — table cell
borders, fill transparency, shadow parameters, line arrow heads, doughnut hole size.

The usual workaround is to build the artwork in PowerPoint or Excel by hand and paste it
in, which defeats the purpose of generating decks programmatically. This fork implements
those features against the real ECMA-376 schema, with the same test discipline as the
upstream project, so a report can be produced end-to-end from Python.

Every addition is **purely additive**. No upstream API was changed or removed, and the
full test suite — upstream's own tests plus this fork's — passes green at 2795 tests.

---

## What this fork adds

Each entry links to the upstream issue it answers, where one exists.

### Tables

| Feature | API | Upstream |
| --- | --- | --- |
| Per-cell borders (4 edges + 2 diagonals), each a full `LineFormat` | `cell.borders.left.width`, `.color`, `.dash_style` | [#71](https://github.com/scanny/python-pptx/issues/71) |
| Reusable cell styles, range/row/column styling, exact sizing, run merging | [`pptx.tablekit`](docs/ext/tables.rst) — `StyledTable`, `CellStyle`, `BorderSpec` | [#27](https://github.com/scanny/python-pptx/issues/27), [#203](https://github.com/scanny/python-pptx/issues/203), [#573](https://github.com/scanny/python-pptx/issues/573) |
| Row heights that match what PowerPoint actually renders | `StyledTable.autosize_rows()` | [#480](https://github.com/scanny/python-pptx/issues/480), [#296](https://github.com/scanny/python-pptx/issues/296) |

### Charts

| Feature | API | Upstream |
| --- | --- | --- |
| Office 2016 "modern" charts: waterfall, funnel, treemap, sunburst, box & whisker | `shapes.add_chartex()`, `ChartExData`, `XL_CHART_EX_TYPE` | not supported upstream |
| Doughnut hole size and first-slice rotation | `plot.hole_size`, `plot.first_slice_angle` | [#493](https://github.com/scanny/python-pptx/issues/493) |
| Date-axis time units | `axis.base_time_unit`, `.major_time_unit`, `.minor_time_unit` (`XL_TIME_UNIT`) | not supported upstream |

### Text

| Feature | API | Upstream |
| --- | --- | --- |
| Strikethrough, superscript, subscript | `font.strikethrough`, `font.superscript`, `font.subscript` | [#339](https://github.com/scanny/python-pptx/issues/339), [#574](https://github.com/scanny/python-pptx/issues/574) |
| Bullet control: none / character / auto-number, with font, colour and size | `paragraph.bullet.none()`, `.character()`, `.auto_number()` | [#100](https://github.com/scanny/python-pptx/issues/100), [#114](https://github.com/scanny/python-pptx/issues/114) |
| Character spacing (tracking), caps effect, highlight colour | `font.spacing`, `font.caps`, `font.highlight_color` | not supported upstream |
| Paragraph indents, including hanging indent | `paragraph.margin_left`, `.margin_right`, `.first_line_indent` | not supported upstream |
| Vertical / rotated text | `text_frame.text_direction` (`MSO_TEXT_DIRECTION`) | not supported upstream |
| Font-file lookup on Linux, plus a directory override | `PYTHON_PPTX_FONT_DIRECTORY` env var | [PR #1039](https://github.com/scanny/python-pptx/pull/1039) |

### Shapes, colour and effects

| Feature | API | Upstream |
| --- | --- | --- |
| Fill and line colour transparency | `color.transparency` | [#62](https://github.com/scanny/python-pptx/issues/62) |
| Picture transparency | `picture.transparency` | not supported upstream |
| Shadow blur, distance, direction, colour, and soft edges | `shape.shadow.blur_radius`, `.distance`, `.direction`, `.color`, `.soft_edge_radius` | not supported upstream |
| Line arrow heads and end caps | `line.begin_arrow`, `line.end_arrow`, `line.cap` | [#375](https://github.com/scanny/python-pptx/issues/375) |
| Alt text on any shape (accessibility) | `shape.alt_text` | not supported upstream |

Full documentation for every item lives in [`docs/ext/`](docs/ext/index.rst).

---

## Installation

```bash
pip install python-pptx-op
```

Or straight from git, for the development version:

```bash
pip install git+https://github.com/vitordj/python-pptx-op.git
```

Requires Python 3.9+ and the same dependencies as upstream (`lxml`, `Pillow`,
`XlsxWriter`, `typing_extensions`).

> [!WARNING]
> **Do not install this alongside `python-pptx`.** Both distributions provide the same
> `pptx` package, so they overwrite each other's files. pip does *not* detect the
> conflict: it reports both as installed, and you get a silently broken mixture — the
> file-level additions (cell borders, shadows, transparency, bullets) disappear, while
> `pptx.tablekit` and `pptx.chartex` remain importable and then fail at runtime with
> errors like `AttributeError: '_Cell' object has no attribute 'borders'`.
>
> If `pip list | grep pptx` shows both, fix it with:
>
> ```bash
> pip uninstall -y python-pptx python-pptx-op && pip install python-pptx-op
> ```

### Migrating from python-pptx

There is nothing to migrate. The import path, class names and behaviour are unchanged:

```python
from pptx import Presentation   # identical to upstream
from pptx.util import Cm, Pt    # identical to upstream
```

The new capabilities are additional properties on the objects you already use, plus two
new modules (`pptx.tablekit` and `pptx.chartex`) that you only import if you want them.

---

## A tour of the additions

### Table cell borders

Borders are the single most requested table feature upstream. Each edge is an ordinary
`LineFormat`, so it behaves like every other line in the library:

```python
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.util import Pt

cell = table.cell(0, 0)
cell.borders.bottom.width = Pt(1.5)
cell.borders.bottom.color.rgb = RGBColor(0x8C, 0xC6, 0x40)
cell.borders.left.dash_style = MSO_LINE_DASH_STYLE.DASH
```

`pptx.tablekit` builds on this with reusable styles, so a formatted report table is a
handful of declarative lines rather than four nested loops. See
[the tables guide](docs/ext/tables.rst).

### Modern charts

```python
from pptx.chartex import ChartExData
from pptx.enum.chart import XL_CHART_EX_TYPE
from pptx.util import Cm

data = ChartExData()
data.categories = ["Opening", "Q1", "Q2", "Q3", "Closing"]
data.add_series("Cash flow", (100, 25, -12, 18, 131))

slide.shapes.add_chartex(XL_CHART_EX_TYPE.WATERFALL,
                         Cm(2), Cm(2), Cm(20), Cm(11), data)
```

Treemap and sunburst take hierarchical categories (root-first tuples). See
[the charts guide](docs/ext/charts.rst).

### Effects and colour

```python
from pptx.util import Emu, Pt

shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(0x1F, 0x77, 0xB4)
shape.fill.fore_color.transparency = 0.35     # 35% transparent

shape.shadow.inherit = False
shape.shadow.blur_radius = Pt(8)
shape.shadow.distance = Pt(4)
shape.shadow.direction = 45.0                 # degrees clockwise from east
shape.shadow.color.rgb = RGBColor(0, 0, 0)
```

See [the effects guide](docs/ext/effects.rst).

---

## Documentation

| | |
| --- | --- |
| [Fork extensions overview](docs/ext/index.rst) | Everything this fork adds, by area |
| [Tables and tablekit](docs/ext/tables.rst) | Cell borders, styled tables, sizing |
| [Charts](docs/ext/charts.rst) | chartEx family, doughnut, date axes |
| [Text](docs/ext/text.rst) | Runs, bullets, indents, direction, fonts |
| [Shapes and effects](docs/ext/effects.rst) | Transparency, shadows, lines, alt text |
| [python-pptx docs](https://python-pptx.readthedocs.io/en/latest/) | The upstream API, all of which still applies |

The docs are Sphinx sources; build them with `make -C docs html`.

---

## Development

```bash
pip install -e .
pip install -r requirements-test.txt
pytest                       # full suite: upstream + this fork's tests
```

Tests for the added features live alongside the upstream ones and are suffixed `_ext.py`
(for example `tests/dml/test_effects_ext.py`), which keeps the upstream files untouched
and makes the fork's surface easy to review. Behaviour tests use `behave` under
`features/`.

The project keeps upstream's tooling: `black` and `ruff` at a 100-column line length, and
`pyright` in strict mode over `src/pptx` and `tests`.

Publishing a release is documented in [RELEASING.md](RELEASING.md).

---

## Roadmap

- Unit tests for `pptx.tablekit` and `pptx.chartex`, the two modules currently without
  dedicated coverage.
- Remaining chartEx layouts: histogram, Pareto and region map are defined in
  `XL_CHART_EX_TYPE` but not yet emitted by the writer.

---

## Credits and licence

python-pptx was created by [Steve Canny](https://github.com/scanny) and is the foundation
of everything here; this fork is a thin layer of additional features on top of his work.

Released under the MIT licence, the same as upstream. See [LICENSE](LICENSE).
