"""
Render workbook tables that use Excel merge-and-center within the data block.
"""

from __future__ import annotations

import html
import re
from collections.abc import Callable
from typing import Any

import pandas as pd
from IPython.display import HTML

from scripts.cha_table_styling import (
    CELL_PADDING_PX,
    CHA_FONT_FAMILY,
    EXCEL_INDENT_PX_PER_LEVEL,
)
from scripts.workbook_loader import CellMerge, CellStyle, MergedTableGrid

# Body striping: first data row white, then green, then alternate.
# Header rows always use _HEADER_BG (including multiheader sub-rows).
_ROW_ODD_BG = "#FFFFFF"
_ROW_EVEN_BG = "#EAF5DB"
_HEADER_BG = "#EAF5DB"
_BORDER_COLOR = "#5a8f3c"


def _looks_numeric(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return True
    text = str(value).strip()
    if not text:
        return False
    text = text.replace(",", "").replace("$", "").rstrip("%").strip()
    text = re.sub(r"[*†‡§#]+$", "", text)
    try:
        float(text)
        return True
    except ValueError:
        return False


def _is_label_cell(value: Any, *, is_header: bool) -> bool:
    if is_header:
        return True
    text = str(value).strip() if value is not None else ""
    if not text:
        return True
    if text.lower() in {"count", "rate"}:
        return True
    return not _looks_numeric(value)


def _build_merge_lookup(
    merges: tuple[CellMerge, ...],
    n_rows: int,
    n_cols: int,
) -> tuple[dict[tuple[int, int], CellMerge], set[tuple[int, int]]]:
    anchors: dict[tuple[int, int], CellMerge] = {}
    covered: set[tuple[int, int]] = set()
    for merge in merges:
        anchors[(merge.row, merge.col)] = merge
        for dr in range(merge.rowspan):
            for dc in range(merge.colspan):
                pos = (merge.row + dr, merge.col + dc)
                if dr == 0 and dc == 0:
                    continue
                if 0 <= pos[0] < n_rows and 0 <= pos[1] < n_cols:
                    covered.add(pos)
    return anchors, covered


def _cell_style_props(
    grid: MergedTableGrid,
    row_idx: int,
    col_idx: int,
    *,
    raw_value: Any,
    is_header: bool,
) -> tuple[str, list[str]]:
    """Return font-weight and extra CSS props from Excel styles or heuristics."""
    if grid.styles and row_idx < len(grid.styles):
        row_styles = grid.styles[row_idx]
        if col_idx < len(row_styles):
            style: CellStyle = row_styles[col_idx]
            weight = "bold" if style.bold else "normal"
            extras: list[str] = []
            if style.indent > 0:
                indent_px = style.indent * EXCEL_INDENT_PX_PER_LEVEL
                # padding-left indents every line; text-indent only indents the first.
                extras.append(f"padding-left:{CELL_PADDING_PX + indent_px}px")
                extras.append("text-align:left !important")
            elif col_idx == 0:
                extras.append("text-align:left !important")
            return weight, extras

    label_cell = _is_label_cell(raw_value, is_header=is_header)
    weight = "bold" if label_cell else "normal"
    return weight, []


def _row_format_override(row: tuple[Any, ...]) -> str | None:
    """Count/Rate block rows override per-column format codes from the workbook."""
    if not row:
        return None
    label = str(row[0]).strip().lower() if row[0] is not None else ""
    if label == "count":
        return "integer"
    if label == "rate":
        return "percent1"
    return None


def _header_display_text(value: Any) -> str:
    """Render header cells as plain text (never percent/number formats)."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if pd.isna(value):
            return ""
        if value.is_integer():
            return str(int(value))
        return str(value).strip()
    text = str(value).strip()
    if re.fullmatch(r"-?\d+\.0+", text):
        return text.split(".", 1)[0]
    return text


def _format_grid_value(
    value: Any,
    col_idx: int,
    format_rules_by_col: tuple[str, ...],
    format_fn: Callable[[Any, str], Any],
    row_format_override: str | None = None,
    *,
    is_header: bool = False,
) -> str:
    if value is None:
        return ""
    # Headers are always plain text — never apply percent1/percent2/etc.
    if is_header:
        text = _header_display_text(value)
        return html.escape(text).replace("\n", "<br>") if text else ""

    text = str(value).strip()
    if text == "":
        return ""
    if _is_label_cell(value, is_header=False) and row_format_override is None:
        return html.escape(text).replace("\n", "<br>")

    fmt_name = row_format_override or ""
    if not fmt_name and 0 <= col_idx < len(format_rules_by_col):
        fmt_name = format_rules_by_col[col_idx]
    if fmt_name:
        formatted = format_fn(value, fmt_name)
        text = str(formatted).strip()
    return html.escape(text).replace("\n", "<br>")


def _render_row(
    row_idx: int,
    row: tuple[Any, ...],
    *,
    grid: MergedTableGrid,
    anchors: dict[tuple[int, int], CellMerge],
    covered: set[tuple[int, int]],
    format_fn: Callable[[Any, str], Any],
    body_row_index: int = 0,
) -> str:
    is_header = row_idx < grid.header_rows
    tag = "th" if is_header else "td"
    cells_html: list[str] = []
    row_format_override = None if is_header else _row_format_override(row)

    # Alternating row background for body rows; headers use a fixed bg.
    if is_header:
        row_bg = _HEADER_BG
    else:
        row_bg = _ROW_ODD_BG if body_row_index % 2 == 0 else _ROW_EVEN_BG

    for col_idx, raw_value in enumerate(row):
        if (row_idx, col_idx) in covered:
            continue

        merge = anchors.get((row_idx, col_idx))
        colspan = merge.colspan if merge else 1
        rowspan = merge.rowspan if merge else 1
        display_value = _format_grid_value(
            raw_value,
            col_idx,
            grid.format_rules_by_col,
            format_fn,
            row_format_override=row_format_override,
            is_header=is_header,
        )

        bg = row_bg

        weight, style_extras = _cell_style_props(
            grid,
            row_idx,
            col_idx,
            raw_value=raw_value,
            is_header=is_header,
        )

        attrs = [
            "text-align:center",
            "vertical-align:middle",
            "padding:10px",
            f"border:1px solid {_BORDER_COLOR}",
            f"font-family:{CHA_FONT_FAMILY}",
            f"background-color:{bg}",
            # Quarto wraps tables with Bootstrap table-striped; kill its grey overlay.
            "box-shadow:none !important",
            "--bs-table-bg-type:transparent",
            "--bs-table-accent-bg:transparent",
            # Keep labels on one line unless Excel Alt+Enter added an explicit break.
            "white-space:nowrap",
            f"font-weight:{weight}",
            *style_extras,
        ]
        style = ";".join(attrs).replace('"', "'")
        attr_parts = [f'style="{style}"']
        if colspan > 1:
            attr_parts.append(f'colspan="{colspan}"')
        if rowspan > 1:
            attr_parts.append(f'rowspan="{rowspan}"')

        cells_html.append(f"<{tag} {' '.join(attr_parts)}>{display_value}</{tag}>")

    return f"<tr>{''.join(cells_html)}</tr>"


def render_merged_table(
    grid: MergedTableGrid,
    *,
    format_fn: Callable[[Any, str], Any] | None = None,
) -> HTML:
    """
    Render a merged workbook table as HTML with colspan/rowspan preserved.

    Parameters
    ----------
    grid : MergedTableGrid
        Cell values and merge metadata extracted from the workbook.
    format_fn : callable, optional
        ``(value, format_name) -> formatted_value`` used for numeric columns.
    """
    if not grid.cells:
        return HTML("<table></table>")

    n_rows = len(grid.cells)
    n_cols = len(grid.cells[0]) if grid.cells else 0
    anchors, covered = _build_merge_lookup(grid.merges, n_rows, n_cols)
    format_fn = format_fn or (lambda value, _fmt: value)

    header_rows_html = [
        _render_row(
            row_idx,
            row,
            grid=grid,
            anchors=anchors,
            covered=covered,
            format_fn=format_fn,
        )
        for row_idx, row in enumerate(grid.cells[: grid.header_rows])
    ]
    body_rows_html = [
        _render_row(
            row_idx,
            row,
            grid=grid,
            anchors=anchors,
            covered=covered,
            format_fn=format_fn,
            body_row_index=body_idx,
        )
        for body_idx, (row_idx, row) in enumerate(
            zip(
                range(grid.header_rows, n_rows),
                grid.cells[grid.header_rows :],
            )
        )
    ]

    table_style = (
        f"border-collapse:collapse;width:100%;margin:20px 0;"
        f"font-family:{CHA_FONT_FAMILY};font-size:14px;"
    ).replace('"', "'")
    parts = [f'<table style="{table_style}">']
    if header_rows_html:
        parts.append(f"<thead>{''.join(header_rows_html)}</thead>")
    if body_rows_html:
        parts.append(f"<tbody>{''.join(body_rows_html)}</tbody>")
    parts.append("</table>")
    return HTML("".join(parts))
