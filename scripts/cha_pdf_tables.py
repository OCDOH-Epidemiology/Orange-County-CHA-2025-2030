"""
PDF-specific table rendering helpers for Quarto/Jupyter output.
"""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, Literal

import pandas as pd

if TYPE_CHECKING:
    from scripts.workbook_loader import CellStyle

TablePolicy = Literal["fit", "split"]
ResolvedTablePolicy = Literal["fit", "split"]

DEFAULT_TABLE_POLICY: Literal["auto", "fit", "split"] = "auto"

# Per-table behavior overrides. Keep this map small and explicit.
TABLE_POLICY_OVERRIDES: dict[str, TablePolicy] = {
    "tbl-population-demographics": "split",
    "tbl-age": "split",
    "tbl-race": "split",
    "tbl-income": "split",
}


def is_pdf_render() -> bool:
    """True when Quarto is currently executing for PDF output."""
    candidates = (
        os.environ.get("QUARTO_FORMAT"),
        os.environ.get("QUARTO_EXECUTE_INFO"),
    )
    for raw in candidates:
        if raw and "pdf" in raw.lower():
            return True
    return False


def _looks_like_overflow(df: pd.DataFrame) -> bool:
    """Width-first overflow check for right-edge page overflow."""
    return _estimate_table_width_units(df) > 92


def resolve_table_policy(table_id: str, df: pd.DataFrame) -> ResolvedTablePolicy:
    """Resolve policy from explicit overrides, then overflow detection."""
    if table_id in TABLE_POLICY_OVERRIDES:
        return TABLE_POLICY_OVERRIDES[table_id]
    if DEFAULT_TABLE_POLICY in {"fit", "split"}:
        return DEFAULT_TABLE_POLICY
    return "split" if _looks_like_overflow(df) else "fit"


_LATEX_ESCAPES: tuple[tuple[str, str], ...] = (
    ("\\", r"\textbackslash{}"),
    ("&", r"\&"),
    ("%", r"\%"),
    ("$", r"\$"),
    ("#", r"\#"),
    ("_", r"\_"),
    ("{", r"\{"),
    ("}", r"\}"),
    ("~", r"\textasciitilde{}"),
    ("^", r"\textasciicircum{}"),
)


def _latex_escape_text(text: str) -> str:
    for char, replacement in _LATEX_ESCAPES:
        text = text.replace(char, replacement)
    return text


def _latex_styled_cell(text: str, style: CellStyle) -> str:
    escaped = _latex_escape_text(text)
    if style.indent > 0:
        escaped = (r"\hspace{2em}" * style.indent) + escaped
    if style.bold:
        escaped = r"\textbf{" + escaped + "}"
    return escaped


def _apply_cell_styles_for_pdf(
    df: pd.DataFrame,
    cell_styles: tuple[tuple[CellStyle, ...], ...] | None,
) -> pd.DataFrame:
    if not cell_styles:
        return df

    out = df.copy().astype(object)
    for row_idx in range(len(df)):
        for col_idx in range(len(df.columns)):
            raw_value = df.iat[row_idx, col_idx]
            if raw_value is None or (isinstance(raw_value, float) and pd.isna(raw_value)):
                text = ""
            else:
                text = str(raw_value).strip()

            style = None
            if row_idx < len(cell_styles):
                row_styles = cell_styles[row_idx]
                if col_idx < len(row_styles):
                    style = row_styles[col_idx]

            if style and (style.bold or style.indent > 0) and text:
                out.iat[row_idx, col_idx] = _latex_styled_cell(text, style)
            elif text:
                out.iat[row_idx, col_idx] = _latex_escape_text(text)
            else:
                out.iat[row_idx, col_idx] = ""
    return out


def _flatten_columns_for_pdf(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if isinstance(out.columns, pd.MultiIndex):
        flattened: list[str] = []
        for parts in out.columns.to_flat_index():
            labels = [str(part).strip() for part in parts if str(part).strip() and str(part).strip() != "."]
            flattened.append(" / ".join(labels) if labels else "")
        out.columns = flattened
    else:
        out.columns = [str(col) for col in out.columns]
    return out


def _df_to_longtable_latex(df: pd.DataFrame, *, escape: bool = True) -> str:
    col_format = _column_format(df)
    body = df.to_latex(
        index=False,
        na_rep="",
        longtable=True,
        escape=escape,
        column_format=col_format,
    )
    return "\n".join(
        [
            r"\begingroup",
            r"\small",
            r"\setlength{\tabcolsep}{3pt}",
            r"\setlength{\LTleft}{0pt}",
            r"\setlength{\LTright}{0pt}",
            body,
            r"\endgroup",
        ]
    )


def _df_to_fit_latex(df: pd.DataFrame, *, escape: bool = True) -> str:
    col_format = _column_format(df)
    tabular = df.to_latex(
        index=False,
        na_rep="",
        longtable=False,
        escape=escape,
        column_format=col_format,
    )
    return "\n".join(
        [
            r"\begingroup",
            r"\setlength{\tabcolsep}{4pt}",
            r"\renewcommand{\arraystretch}{1.1}",
            r"\small",
            r"\resizebox{\linewidth}{!}{%",
            tabular,
            r"}",
            r"\endgroup",
        ]
    )


def _column_width_units(series: pd.Series, header: str) -> int:
    sample = series.head(20)
    max_cell_len = 0
    if not sample.empty:
        lengths = sample.map(lambda value: len(str(value).strip()))
        max_cell_len = int(lengths.max() if not lengths.empty else 0)
    header_len = len(str(header).strip())
    effective = max(header_len, max_cell_len)
    # Cap contribution so one extreme value does not dominate.
    effective = min(effective, 44)
    # Include padding/border pressure.
    return effective + 4


def _estimate_table_width_units(df: pd.DataFrame) -> int:
    if df.empty and len(df.columns) == 0:
        return 0
    width = 0
    for col in df.columns:
        width += _column_width_units(df[col], str(col))
    return width


def _split_wide_columns(df: pd.DataFrame, max_columns_per_part: int = 4) -> list[pd.DataFrame]:
    if len(df.columns) <= max_columns_per_part and not _looks_like_overflow(df):
        return [df]

    first_col = df.columns[0]
    trailing = list(df.columns[1:])
    first_col_units = _column_width_units(df[first_col], str(first_col))
    part_budget_units = 76
    parts: list[pd.DataFrame] = []

    current_cols: list[str] = [first_col]
    current_units = first_col_units
    for col in trailing:
        col_units = _column_width_units(df[col], str(col))
        would_overflow = (current_units + col_units) > part_budget_units
        reached_col_cap = len(current_cols) >= max_columns_per_part
        if len(current_cols) > 1 and (would_overflow or reached_col_cap):
            parts.append(df[current_cols].copy())
            current_cols = [first_col, col]
            current_units = first_col_units + col_units
            continue
        current_cols.append(col)
        current_units += col_units

    if len(current_cols) > 1:
        parts.append(df[current_cols].copy())

    return parts or [df]


def _column_format(df: pd.DataFrame) -> str:
    """Use wrapping paragraph columns so long headers/cells break lines."""
    col_count = len(df.columns)
    if col_count <= 1:
        return "p{0.95\\linewidth}"
    first_col_width = 0.28
    other_col_width = (0.95 - first_col_width) / max(1, col_count - 1)
    other_col_width = max(0.12, other_col_width)
    return "p{%.2f\\linewidth}%s" % (
        first_col_width,
        "".join([f"p{{{other_col_width:.2f}\\linewidth}}" for _ in range(col_count - 1)]),
    )


def render_pdf_table_latex(
    table_id: str,
    df: pd.DataFrame,
    *,
    cell_styles: tuple[tuple[CellStyle, ...], ...] | None = None,
) -> str:
    table_df = _flatten_columns_for_pdf(df)
    has_cell_styles = bool(cell_styles)
    styled_df = _apply_cell_styles_for_pdf(table_df, cell_styles if has_cell_styles else None)
    latex_escape = not has_cell_styles
    policy = resolve_table_policy(table_id, table_df)

    if policy == "split":
        parts = _split_wide_columns(styled_df)
        total = len(parts)
        rendered_parts = []
        for idx, part in enumerate(parts, start=1):
            part_latex = _df_to_longtable_latex(part, escape=latex_escape)
            if total > 1:
                rendered_parts.append(
                    "\n".join(
                        [
                            rf"\textit{{Table continuation ({idx}/{total})}}",
                            r"\vspace{0.3em}",
                            part_latex,
                            r"\vspace{0.8em}",
                        ]
                    )
                )
            else:
                rendered_parts.append(part_latex)
        return "\n".join(rendered_parts)

    return _df_to_fit_latex(styled_df, escape=latex_escape)
