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

# Thick red borders everywhere
set_border(ws, "A1:D5", style="thick", color="FF0000")

# Thick outer edges, thin inner grid
set_border(ws, "A1:D5", outline="thick", inside="thin")

# Compact numeric shorthand: thick outer, thin inner grid
set_border(ws, "A1:D5", custom=(3, 3, 3, 3, 1, 1))

# Override a single edge with shorthand (no Side import needed)
set_border(ws, "A1:D5", outline="medium", left=("double", "FF0000"))

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
| 1     | `color`            | `str \| None`              | Base border color for all sides (hex, e.g. `"FF0000"`). |
| 2     | `custom`           | `tuple[int, ...] \| None`  | Weight tuple in CSS-like order (see below).           |
| 3     | `outline`          | `str \| None`              | Style for all 4 outer edges.                         |
| 3     | `inside`           | `str \| None`              | Style for inner horizontal and vertical lines.       |
| 4     | `horizontal`       | `str \| None`              | Style for top, bottom, and inner horizontal.         |
| 4     | `vertical`         | `str \| None`              | Style for left, right, and inner vertical.           |
| 5     | `left`             | `SideSpec`                 | Left edge (see Side shorthand below).                |
| 5     | `right`            | `SideSpec`                 | Right edge.                                          |
| 5     | `top`              | `SideSpec`                 | Top edge.                                            |
| 5     | `bottom`           | `SideSpec`                 | Bottom edge.                                         |
| 5     | `inner_horizontal` | `SideSpec`                 | Inner horizontal lines.                              |
| 5     | `inner_vertical`   | `SideSpec`                 | Inner vertical lines.                                |

Higher-layer parameters override lower ones. The `color` param is inherited by all layers unless a side specifies its own color.

### Side shorthand (`SideSpec`)

Individual side parameters (`left`, `right`, `top`, `bottom`, `inner_horizontal`, `inner_vertical`) accept three forms:

| Form               | Example                          | Result                                      |
| ------------------ | -------------------------------- | ------------------------------------------- |
| `str`              | `left="thick"`                   | `Side(style="thick")`, inherits base `color` |
| `(str, str)` tuple | `left=("thick", "FF0000")`       | `Side(style="thick", color="FF0000")`        |
| `Side`             | `left=Side(style="thick", ...)` | Used as-is (full control)                    |

This eliminates the need to import `Side` for most use cases.

### The `custom` tuple

A compact way to set all 6 border positions with numeric weights:

```
(top, right, bottom, left, inner_horizontal, inner_vertical)
```

Weight mapping: `0` = none, `1` = thin, `2` = medium, `3` = thick.

The tuple length must be **4** or **6**. A 4-element tuple sets the outer edges only; inner borders default to `0` (none). The base `color` is applied to all sides created by `custom`.

```python
# 6-element: thick outer, thin horizontal grid, medium vertical grid
set_border(ws, "A1:D5", custom=(3, 3, 3, 3, 1, 2))

# 4-element: medium outer edges, no inner grid
set_border(ws, "A1:D5", custom=(2, 2, 2, 2))

# With color: all red
set_border(ws, "A1:D5", custom=(3, 3, 3, 3, 1, 1), color="FF0000")
```

## Examples

### All borders in one color

```python
set_border(ws, "A1:D5", style="thick", color="0000FF")
```

### Box border (outline only, no inner grid)

```python
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

### Override a single edge with shorthand

```python
# String shorthand (inherits base color)
set_border(ws, "A1:D5", outline="medium", color="0000FF", left="double")

# Tuple shorthand (own color overrides base)
set_border(ws, "A1:D5", outline="medium", left=("double", "FF0000"))
```

### Using `custom` with overrides

Higher-priority layers refine on top of `custom`:

```python
set_border(
    ws, "A1:D5",
    custom=(3, 2, 3, 2, 1, 1),
    color="0000FF",
    left=("double", "FF0000"),  # override left: double red
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

MIT - See [LICENSE](LICENSE) for details.
