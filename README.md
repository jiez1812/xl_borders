# xl-borders

A Python library that simplifies applying cell borders to [openpyxl](https://openpyxl.readthedocs.io/) worksheets using Excel-style range strings (e.g. `"A1:D5"`), inspired by VBA's `Range.Borders` API.

## Installation

```bash
pip install xl-borders
```

Requires Python 3.13+ and openpyxl.

## Quick Start

```python
from openpyxl import Workbook
from xl_borders import set_border

wb = Workbook()
ws = wb.active

# Thin borders on all sides (default)
set_border(ws, "A1:D5")

# Thick outer edges, thin inner grid
set_border(ws, "A1:D5", outline="thick", inside="thin")

# Compact numeric shorthand: thick outer, thin inner grid
set_border(ws, "A1:D5", custom=(3, 3, 3, 3, 1, 1))

wb.save("output.xlsx")
```

## API

### `set_border(ws, cell_range, **kwargs)`

Apply borders to a rectangular cell range.

**Positional arguments:**

| Argument     | Type        | Description                                        |
| ------------ | ----------- | -------------------------------------------------- |
| `ws`         | `Worksheet` | The openpyxl worksheet to modify.                  |
| `cell_range` | `str`       | Excel-style range string (e.g. `"A1:D5"`, `"B2"`). |

**Keyword-only arguments (in priority order, lowest to highest):**

| Layer | Parameter          | Type                       | Description                                          |
| ----- | ------------------ | -------------------------- | ---------------------------------------------------- |
| 1     | `style`            | `str`                      | Base border style for all 6 sides. Default: `"thin"`. |
| 2     | `custom`           | `tuple[int, ...] \| None`  | Weight tuple in CSS-like order (see below).           |
| 3     | `outline`          | `str \| None`              | Style for all 4 outer edges.                         |
| 3     | `inside`           | `str \| None`              | Style for inner horizontal and vertical lines.       |
| 4     | `horizontal`       | `str \| None`              | Style for top, bottom, and inner horizontal.         |
| 4     | `vertical`         | `str \| None`              | Style for left, right, and inner vertical.           |
| 5     | `left`             | `Side \| None`             | Explicit `Side` for the left edge.                   |
| 5     | `right`            | `Side \| None`             | Explicit `Side` for the right edge.                  |
| 5     | `top`              | `Side \| None`             | Explicit `Side` for the top edge.                    |
| 5     | `bottom`           | `Side \| None`             | Explicit `Side` for the bottom edge.                 |
| 5     | `inner_horizontal` | `Side \| None`             | Explicit `Side` for inner horizontal lines.          |
| 5     | `inner_vertical`   | `Side \| None`             | Explicit `Side` for inner vertical lines.            |

Higher-layer parameters override lower ones.

### The `custom` tuple

A compact way to set all 6 border positions with numeric weights:

```
(top, right, bottom, left, inner_horizontal, inner_vertical)
```

Weight mapping: `0` = none, `1` = thin, `2` = medium, `3` = thick.

The tuple length must be **4** or **6**. A 4-element tuple sets the outer edges only; inner borders default to `0` (none).

```python
# 6-element: thick outer, thin horizontal grid, medium vertical grid
set_border(ws, "A1:D5", custom=(3, 3, 3, 3, 1, 2))

# 4-element: medium outer edges, no inner grid
set_border(ws, "A1:D5", custom=(2, 2, 2, 2))
```

## Examples

### Box border (outline only, no inner grid)

```python
from openpyxl.styles import Side

set_border(
    ws, "A1:D5",
    outline="medium",
    inner_horizontal=Side(style=None),
    inner_vertical=Side(style=None),
)
```

### Different horizontal and vertical weights

```python
set_border(ws, "A1:D5", horizontal="thin", vertical="thick")
```

### Override a single edge

```python
from openpyxl.styles import Side

set_border(
    ws, "A1:D5",
    outline="medium",
    left=Side(style="double", color="FF0000"),
)
```

### Using `custom` with overrides

Higher-priority layers refine on top of `custom`:

```python
from openpyxl.styles import Side

# Start with custom weights, then override left edge
set_border(
    ws, "A1:D5",
    custom=(3, 2, 3, 2, 1, 1),
    left=Side(style="double"),
)
```

## Available border styles

Any style string accepted by openpyxl's `Side(style=...)`:

`thin`, `medium`, `thick`, `double`, `dashed`, `mediumDashed`, `dotted`, `mediumDashDot`, `dashDot`, `mediumDashDotDot`, `dashDotDot`, `hair`, `slantDashDot`

## Development

```powershell
# Install dependencies
uv sync

# Run tests
uv run pytest Test/ -v
```

## License

See [LICENSE](LICENSE) for details.
