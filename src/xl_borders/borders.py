"""Core border-building functions for openpyxl worksheets."""

from openpyxl.styles import Border, Side
from openpyxl.utils import range_boundaries
from openpyxl.worksheet.worksheet import Worksheet

THIN = Side(style="thin")
NO_SIDE = Side(style=None)

DEFAULT_SIDE = THIN

WEIGHT_STYLES: dict[int, str | None] = {0: None, 1: "thin", 2: "medium", 3: "thick"}


def set_border(
    ws: Worksheet,
    cell_range: str,
    *,
    left: Side | None = None,
    right: Side | None = None,
    top: Side | None = None,
    bottom: Side | None = None,
    inner_horizontal: Side | None = None,
    inner_vertical: Side | None = None,
    style: str = "thin",
    custom: tuple[int, ...] | None = None,
    outline: str | None = None,
    inside: str | None = None,
    horizontal: str | None = None,
    vertical: str | None = None,
) -> None:
    """Apply borders to a rectangular cell range, similar to Excel VBA Range.Borders.

    Parameters resolve in layers (lowest to highest priority):
        1. ``style`` -- base default for all 6 sides
        2. ``custom`` -- tuple of weight ints overriding all 6 sides
        3. ``outline`` -- overrides outer edges (left, right, top, bottom)
        4. ``inside`` -- overrides inner lines (inner_horizontal, inner_vertical)
        5. ``horizontal`` -- overrides top, bottom, inner_horizontal
        6. ``vertical`` -- overrides left, right, inner_vertical
        7. Individual ``Side`` params -- highest priority

    Args:
        ws: The openpyxl Worksheet to modify.
        cell_range: Excel-style range string (e.g. "A1:D5", "B2", "A1:A1").
        left: Side for the left edge.
        right: Side for the right edge.
        top: Side for the top edge.
        bottom: Side for the bottom edge.
        inner_horizontal: Side for inner horizontal lines.
        inner_vertical: Side for inner vertical lines.
        style: Default border style name applied to all sides.
            Common values: "thin", "thick", "medium", "dashed", "dotted", "double".
        custom: Tuple of weight integers in CSS-like order:
            ``(top, right, bottom, left, inner_horizontal, inner_vertical)``.
            Length must be 4 (outer only; inner defaults to 0) or 6.
            Weight mapping: 0=none, 1=thin, 2=medium, 3=thick.
        outline: Style name applied to all 4 outer edges.
        inside: Style name applied to inner_horizontal and inner_vertical.
        horizontal: Style name applied to top, bottom, and inner_horizontal.
        vertical: Style name applied to left, right, and inner_vertical.
    """
    # --- Layer 1: base default ---
    sides: dict[str, Side] = {k: Side(style=style) for k in (
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
            sides[k] = Side(style=WEIGHT_STYLES[w])

    # --- Layer 3: outline / inside ---
    if outline is not None:
        outline_side = Side(style=outline)
        for k in ("left", "right", "top", "bottom"):
            sides[k] = outline_side
    if inside is not None:
        inside_side = Side(style=inside)
        for k in ("inner_horizontal", "inner_vertical"):
            sides[k] = inside_side

    # --- Layer 4: horizontal / vertical ---
    if horizontal is not None:
        h_side = Side(style=horizontal)
        for k in ("top", "bottom", "inner_horizontal"):
            sides[k] = h_side
    if vertical is not None:
        v_side = Side(style=vertical)
        for k in ("left", "right", "inner_vertical"):
            sides[k] = v_side

    # --- Layer 5: individual Side params (highest priority) ---
    explicit: dict[str, Side | None] = {
        "left": left, "right": right, "top": top, "bottom": bottom,
        "inner_horizontal": inner_horizontal, "inner_vertical": inner_vertical,
    }
    for k, v in explicit.items():
        if v is not None:
            sides[k] = v

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
