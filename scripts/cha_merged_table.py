"""
Render workbook tables that use Excel merge-and-center within the data block.
"""

from __future__ import annotations

import html
import re
from collections.abc import Callable
from typing import Any

from IPython.display import HTML

from scripts.cha_table_styling import CHA_FONT_FAMILY
from scripts.workbook_loader import CellMerge, MergedTableGrid

_LABEL_BG = "#EAF5DB"
_DATA_BG = "#FFFFFF"
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


def _format_grid_value(
    value: Any,
    col_idx: int,
    format_rules_by_col: tuple[str, ...],
    format_fn: Callable[[Any, str], Any],
    row_format_override: str | None = None,
) -> str:
    if value is None:
        return ""
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
) -> str:
    is_header = row_idx < grid.header_rows
    tag = "th" if is_header else "td"
    cells_html: list[str] = []
    row_format_override = None if is_header else _row_format_override(row)

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
        )
        label_cell = _is_label_cell(raw_value, is_header=is_header)
        bg = _LABEL_BG if label_cell else _DATA_BG
        weight = "bold" if label_cell else "normal"

        attrs = [
            "text-align:center",
            "vertical-align:middle",
            "padding:10px",
            f"border:1px solid {_BORDER_COLOR}",
            f"font-family:{CHA_FONT_FAMILY}",
            f"background-color:{bg}",
            f"font-weight:{weight}",
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
        )
        for row_idx, row in enumerate(grid.cells[grid.header_rows :], start=grid.header_rows)
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
