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


def set_border(
    ws: Worksheet,
    cell_range: str,
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
    """Apply borders to a rectangular cell range, similar to Excel VBA Range.Borders.

    Parameters resolve in layers (lowest to highest priority):
        1. ``style`` + ``color`` -- base default for all 6 sides
        2. ``custom`` -- tuple of weight ints overriding all 6 sides
        3. ``outline`` / ``inside`` -- override outer / inner groups
        4. ``horizontal`` / ``vertical`` -- override by orientation
        5. Individual side params -- highest priority

    Individual side params accept shorthand in addition to ``Side``:
        - ``str`` -- style name (inherits base ``color``)
        - ``(str, str)`` -- ``(style, color)`` tuple
        - ``Side`` -- full control

    Args:
        ws: The openpyxl Worksheet to modify.
        cell_range: Excel-style range string (e.g. "A1:D5", "B2", "A1:A1").
        left: Side spec for the left edge.
        right: Side spec for the right edge.
        top: Side spec for the top edge.
        bottom: Side spec for the bottom edge.
        inner_horizontal: Side spec for inner horizontal lines.
        inner_vertical: Side spec for inner vertical lines.
        style: Default border style name applied to all sides.
            Common values: "thin", "thick", "medium", "dashed", "dotted", "double".
        color: Default border color (hex string, e.g. "FF0000") applied to all sides.
        custom: Tuple of weight integers in CSS-like order:
            ``(top, right, bottom, left, inner_horizontal, inner_vertical)``.
            Length must be 4 (outer only; inner defaults to 0) or 6.
            Weight mapping: 0=none, 1=thin, 2=medium, 3=thick.
        outline: Style name applied to all 4 outer edges.
        inside: Style name applied to inner_horizontal and inner_vertical.
        horizontal: Style name applied to top, bottom, and inner_horizontal.
        vertical: Style name applied to left, right, and inner_vertical.
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

    min_col, min_row, max_col, max_row = range_boundaries(cell_range)

    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)

            cell.border = Border(
                left=s_left if col == min_col else s_iv,
                right=s_right if col == max_col else s_iv,
                top=s_top if row == min_row else s_ih,
                bottom=s_bottom if row == max_row else s_ih,
            )
