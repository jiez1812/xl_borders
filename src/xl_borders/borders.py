"""Core border-building functions for openpyxl worksheets."""

from __future__ import annotations

from openpyxl.styles import Border, Side
from openpyxl.utils import range_boundaries
from openpyxl.worksheet.worksheet import Worksheet

THIN = Side(style="thin")
NO_SIDE = Side(style=None)

DEFAULT_SIDE = THIN

WEIGHT_STYLES: dict[int, str | None] = {0: None, 1: "thin", 2: "medium", 3: "thick"}

# Shorthand type accepted for individual side params:
#   Side        - full control
#   str         - style name, inherits base color
#   (str, str)  - (style, color)
type SideSpec = Side | str | tuple[str, str] | None

# Accepted types for cell_range parameter:
#   str                              - "A1:D5" or "B2"
#   (int, int)                       - (row, col) single cell
#   ((int, int), (int, int))         - ((min_row, min_col), (max_row, max_col))
type CellRange = str | tuple[int, int] | tuple[tuple[int, int], tuple[int, int]]


def _merge_side(new: Side, existing: Side | None) -> Side:
    """Merge a new Side with an existing Side on a cell.

    - If *new* has no style, keep *existing* entirely (don't overwrite).
    - If *new* has a style but no color, use the new style with the
      existing side's color (preserve color).
    - Otherwise use *new* as-is.
    """
    if existing is None:
        return new
    if new.style is None:
        return existing
    if new.color is None and existing.color is not None:
        return Side(style=new.style, color=existing.color)
    return new


def _resolve_side(spec: SideSpec, base_color: str | None) -> Side | None:
    """Convert a SideSpec shorthand into a Side object.

    Returns None when *spec* is None (meaning "don't override").
    """
    if spec is None:
        return None
    if isinstance(spec, Side):
        return spec
    if isinstance(spec, str):
        return Side(style=spec, color=base_color)
    if isinstance(spec, tuple):
        style, color = spec
        return Side(style=style, color=color)
    raise TypeError(f"Expected Side, str, tuple, or None; got {type(spec).__name__}")


def _parse_range(cell_range: CellRange) -> tuple[int, int, int, int]:
    """Convert a CellRange into (min_col, min_row, max_col, max_row)."""
    if isinstance(cell_range, str):
        return range_boundaries(cell_range)
    if (
        isinstance(cell_range, tuple)
        and len(cell_range) == 2
        and isinstance(cell_range[0], int)
    ):
        row, col = cell_range
        return col, row, col, row
    if (
        isinstance(cell_range, tuple)
        and len(cell_range) == 2
        and isinstance(cell_range[0], tuple)
    ):
        (min_row, min_col), (max_row, max_col) = cell_range
        return min_col, min_row, max_col, max_row
    raise TypeError(
        f"cell_range must be str, (row, col), or ((row, col), (row, col)); "
        f"got {cell_range!r}"
    )


def set_border(
    ws: Worksheet,
    cell_range: CellRange,
    *,
    left: SideSpec = None,
    right: SideSpec = None,
    top: SideSpec = None,
    bottom: SideSpec = None,
    inner_horizontal: SideSpec = None,
    inner_vertical: SideSpec = None,
    style: str = "thin",
    color: str | None = None,
    custom: tuple[int, ...] | None = None,
    outline: str | None = None,
    inside: str | None = None,
    horizontal: str | None = None,
    vertical: str | None = None,
) -> None:
    """Apply borders to a rectangular cell range.

    Examples::

        # All borders thin (default)
        set_border(ws, "A1:D5")

        # All borders thick and red
        set_border(ws, "A1:D5", style="thick", color="FF0000")

        # Thick outline, thin inner grid
        set_border(ws, "A1:D5", outline="thick", inside="thin")

        # Only horizontal lines
        set_border(ws, "A1:D5", horizontal="thin", vertical=None)

        # Thick bottom edge, everything else thin
        set_border(ws, "A1:D5", bottom="thick")

        # CSS-like shorthand: (top, right, bottom, left)
        set_border(ws, "A1:D5", custom=(3, 1, 3, 1))

        # Using numeric coordinates instead of range string
        set_border(ws, (1, 1))                          # single cell B1
        set_border(ws, ((1, 1), (5, 4)))                # same as "A1:D5"

    Args:
        ws: The openpyxl Worksheet to modify.
        cell_range: Target range. Accepts an Excel-style string (``"A1:D5"``),
            a ``(row, col)`` tuple for a single cell, or a
            ``((min_row, min_col), (max_row, max_col))`` tuple for a range.
        style: Default border style for all sides. Defaults to ``"thin"``.
            Values: ``"thin"``, ``"medium"``, ``"thick"``, ``"dashed"``,
            ``"dotted"``, ``"double"``, etc.
        color: Default border color as a hex string (e.g. ``"FF0000"`` for
            red). Applied to all sides unless overridden.

    Group overrides (override ``style`` for a group of sides):
        outline: Set all 4 outer edges (left, right, top, bottom).
        inside: Set both inner grid lines (inner_horizontal, inner_vertical).
        horizontal: Set top, bottom, and inner_horizontal.
        vertical: Set left, right, and inner_vertical.

    Individual side overrides (highest priority, override everything above):
        left: Left edge only.
        right: Right edge only.
        top: Top edge only.
        bottom: Bottom edge only.
        inner_horizontal: Inner horizontal lines only.
        inner_vertical: Inner vertical lines only.

        Each accepts a style name ``str``, a ``(style, color)`` tuple,
        or an ``openpyxl.styles.Side`` object for full control.

    CSS-like shorthand:
        custom: Tuple of weight integers in ``(top, right, bottom, left)``
            order (4 values, inner lines default to none) or
            ``(top, right, bottom, left, inner_h, inner_v)`` (6 values).
            Weights: 0=none, 1=thin, 2=medium, 3=thick.

    Priority (lowest to highest):
        ``style/color`` < ``custom`` < ``outline/inside`` <
        ``horizontal/vertical`` < individual sides.
    """
    # --- Layer 1: base default (style + color) ---
    sides: dict[str, Side] = {k: Side(style=style, color=color) for k in (
        "left", "right", "top", "bottom", "inner_horizontal", "inner_vertical",
    )}

    # --- Layer 2: custom tuple ---
    if custom is not None:
        if len(custom) not in (4, 6):
            raise ValueError(
                f"custom must have 4 or 6 elements, got {len(custom)}"
            )
        for i, w in enumerate(custom):
            if w not in WEIGHT_STYLES:
                raise ValueError(
                    f"custom[{i}] = {w!r} is not a valid weight "
                    f"(expected one of {sorted(WEIGHT_STYLES)})"
                )
        # Order: top, right, bottom, left, inner_horizontal, inner_vertical
        weights = custom + (0, 0) if len(custom) == 4 else custom
        keys = ("top", "right", "bottom", "left", "inner_horizontal", "inner_vertical")
        for k, w in zip(keys, weights):
            sides[k] = Side(style=WEIGHT_STYLES[w], color=color)

    # --- Layer 3: outline / inside ---
    if outline is not None:
        outline_side = Side(style=outline, color=color)
        for k in ("left", "right", "top", "bottom"):
            sides[k] = outline_side
    if inside is not None:
        inside_side = Side(style=inside, color=color)
        for k in ("inner_horizontal", "inner_vertical"):
            sides[k] = inside_side

    # --- Layer 4: horizontal / vertical ---
    if horizontal is not None:
        h_side = Side(style=horizontal, color=color)
        for k in ("top", "bottom", "inner_horizontal"):
            sides[k] = h_side
    if vertical is not None:
        v_side = Side(style=vertical, color=color)
        for k in ("left", "right", "inner_vertical"):
            sides[k] = v_side

    # --- Layer 5: individual side params (highest priority) ---
    explicit: dict[str, SideSpec] = {
        "left": left, "right": right, "top": top, "bottom": bottom,
        "inner_horizontal": inner_horizontal, "inner_vertical": inner_vertical,
    }
    for k, v in explicit.items():
        resolved = _resolve_side(v, base_color=color)
        if resolved is not None:
            sides[k] = resolved

    s_left = sides["left"]
    s_right = sides["right"]
    s_top = sides["top"]
    s_bottom = sides["bottom"]
    s_ih = sides["inner_horizontal"]
    s_iv = sides["inner_vertical"]

    min_col, min_row, max_col, max_row = _parse_range(cell_range)

    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)
            existing = cell.border

            cell.border = Border(
                left=_merge_side(
                    s_left if col == min_col else s_iv, existing.left,
                ),
                right=_merge_side(
                    s_right if col == max_col else s_iv, existing.right,
                ),
                top=_merge_side(
                    s_top if row == min_row else s_ih, existing.top,
                ),
                bottom=_merge_side(
                    s_bottom if row == max_row else s_ih, existing.bottom,
                ),
            )
