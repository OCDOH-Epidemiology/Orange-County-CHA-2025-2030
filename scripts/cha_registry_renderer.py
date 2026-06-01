"""
Registry-driven rendering for CHA tables and figures.
"""

from __future__ import annotations

from pathlib import Path
from datetime import date, datetime
import re
from typing import Any

import numpy as np
import pandas as pd
import plotly.graph_objects as go

from scripts.cha_figure_builder import (
    build_clustered_bar_figure,
    build_horizontal_bar_figure,
    build_line_figure,
    build_simple_bar_figure,
    build_stacked_bar_figure,
)
from scripts.cha_table_styling import (
    CHA_REGION_ALIASES,
    CHA_REGION_ORDER,
    create_source_callout,
    style_cha_table,
)
from scripts.cha_pdf_tables import is_pdf_render, render_pdf_table_latex
from scripts.workbook_loader import WorkbookModel, load_cha_workbook, _VALID_FORMAT_CODES, _as_text


def _resolve_default_workbook_path() -> Path:
    # Convention: keep the project workbook at data/raw/workbook.xlsx.
    return Path("data/raw/workbook.xlsx")


DEFAULT_WORKBOOK_PATH = _resolve_default_workbook_path()


def _is_time_like_value(value: Any) -> bool:
    """True if *value* looks like a calendar year or year range.

    Excel often reads years as floats (2021.0); those must count as time-like
    so we do not incorrectly pivot region×time wide tables.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    if isinstance(value, (datetime, date)):
        return True
    # pandas.Timestamp
    if hasattr(value, "year") and hasattr(value, "to_pydatetime"):
        return True
    if isinstance(value, (int, np.integer)):
        return 1900 <= int(value) <= 2100
    if isinstance(value, (float, np.floating)):
        if pd.isna(value):
            return False
        if value != value:  # NaN
            return False
        if float(value).is_integer():
            iv = int(value)
            return 1900 <= iv <= 2100
        return False
    text = str(value).strip()
    if not text:
        return False
    if re.fullmatch(r"\d{4}\.0+", text):
        return True
    return bool(
        re.fullmatch(r"\d{4}", text)
        or re.fullmatch(r"\d{4}\s*[-–]\s*\d{4}", text)
        or re.fullmatch(r"\d{4}\s*[-–]\s*\d{2}", text)
    )


def _is_time_like_column(series: pd.Series | pd.DataFrame) -> bool:
    if isinstance(series, pd.DataFrame):
        if series.shape[1] == 0:
            return False
        series = series.iloc[:, 0]
    non_null = [value for value in series.tolist() if not pd.isna(value)]
    if not non_null:
        return False
    matches = sum(1 for value in non_null if _is_time_like_value(value))
    return matches / len(non_null) >= 0.8


def _normalize_region_label_for_axis(label: Any) -> str:
    text = str(label).strip()
    # Some workbook columns prefix county names with a grouping label
    # (e.g., "Three-Year Average|Dutchess"). If the right side is a known
    # region, normalize to that region name so county pivoting can work.
    if "|" in text:
        _prefix, suffix = text.rsplit("|", 1)
        suffix = suffix.strip()
        normalized_suffix = CHA_REGION_ALIASES.get(suffix, suffix)
        if normalized_suffix in CHA_REGION_ORDER:
            return normalized_suffix
    return CHA_REGION_ALIASES.get(text, text)


def _detect_region_columns(df: pd.DataFrame) -> list[str]:
    region_cols: list[str] = []
    columns = list(df.columns)
    if len(columns) <= 1:
        return region_cols
    for col in columns[1:]:
        normalized = _normalize_region_label_for_axis(col)
        if normalized in CHA_REGION_ORDER:
            region_cols.append(col)
    return region_cols


def _resolve_show_data_labels(spec: Any) -> bool:
    if getattr(spec, "show_data_labels", None) is not None:
        return bool(spec.show_data_labels)
    return False


def _load_model(workbook_path: str | Path | None) -> WorkbookModel:
    path = Path(workbook_path) if workbook_path else DEFAULT_WORKBOOK_PATH
    return load_cha_workbook(path)


def _format_value(value: Any, format_name: str) -> Any:
    if pd.isna(value):
        return ""
    if isinstance(value, str) and value.strip().lower() in {"", "nan"}:
        return ""

    def _to_float(v: Any) -> float:
        if isinstance(v, str):
            v = v.strip().lstrip("$").replace(",", "").rstrip("%").strip()
        return float(v)

    def _format_short_date(v: Any) -> Any:
        if pd.isna(v):
            return ""
        if isinstance(v, (pd.Timestamp, datetime, date)):
            dt = pd.Timestamp(v)
        else:
            text = str(v).strip()
            if text == "":
                return ""
            dt = pd.to_datetime(text, errors="coerce")
            if pd.isna(dt):
                return v
        return pd.Timestamp(dt).strftime("%m/%d/%Y")

    def _format_year_like_whole(v: Any) -> str:
        return str(int(round(_to_float(v))))

    if format_name == "number":
        try:
            if _is_time_like_value(value):
                return _format_year_like_whole(value)
            return f"{_to_float(value):,.0f}"
        except (ValueError, TypeError):
            return value
    if format_name == "integer":
        try:
            if _is_time_like_value(value):
                return _format_year_like_whole(value)
            return f"{int(round(_to_float(value))):,}"
        except (ValueError, TypeError):
            return value
    if format_name == "percent1":
        try:
            return f"{_to_float(value):.1f}"
        except (ValueError, TypeError):
            return value
    if format_name == "percent2":
        try:
            return f"{_to_float(value):.2f}"
        except (ValueError, TypeError):
            return value
    if format_name == "currency":
        try:
            return f"${_to_float(value):,.0f}"
        except (ValueError, TypeError):
            return value
    if format_name == "currency2":
        try:
            return f"${_to_float(value):,.2f}"
        except (ValueError, TypeError):
            return value
    if format_name == "date":
        return _format_short_date(value)
    if format_name == "ratio":
        return value
    return value


def _coerce_ratio_numeric(value: Any) -> float:
    try:
        if pd.isna(value):
            return float("nan")
        if isinstance(value, str):
            text = value.strip()
            if ":" in text:
                text = text.split(":", 1)[0].strip()
            text = text.replace(",", "")
            return float(text)
        return float(value)
    except (ValueError, TypeError):
        return float("nan")


def _figure_format_rules(model: WorkbookModel, figure_id: str) -> dict[str, str]:
    if figure_id in model.table_specs:
        return model.table_specs[figure_id].format_rules
    if figure_id.startswith("fig-"):
        paired = f"tbl-{figure_id[4:]}"
        if paired in model.table_specs:
            return model.table_specs[paired].format_rules
    return {}


def _prepare_table_df(df: pd.DataFrame, format_rules: dict[str, str]) -> pd.DataFrame:
    formatted = df.copy().where(~df.isna(), "")
    for col_name, format_name in format_rules.items():
        if col_name not in formatted.columns:
            continue
        col_data = formatted[col_name]
        if isinstance(col_data, pd.DataFrame):
            col_data = col_data.iloc[:, 0]
        formatted[col_name] = col_data.apply(lambda value: _format_value(value, format_name))
    return formatted


def _strip_format_tokens_from_label(label: Any) -> str:
    text = str(label).strip()
    if not text:
        return text
    if "|" in text:
        left, right = text.split("|", 1)
        if _as_text(left) in _VALID_FORMAT_CODES:
            return _as_text(right)
    return text


def _ensure_unique_column_labels(columns: list[Any]) -> list[str]:
    """Return deterministic, unique string labels for DataFrame columns."""
    counts: dict[str, int] = {}
    out: list[str] = []
    for raw in columns:
        base = str(raw).strip() or "Series"
        counts[base] = counts.get(base, 0) + 1
        if counts[base] == 1:
            out.append(base)
        else:
            out.append(f"{base} ({counts[base]})")
    return out


def _rebuild_multiindex(df: pd.DataFrame) -> pd.DataFrame:
    tuples = []
    for col in df.columns:
        if "|" in str(col):
            top, sub = str(col).split("|", 1)
            tuples.append((top.strip(), sub.strip()))
        else:
            tuples.append((" ", str(col)))
    out = df.copy()
    out.columns = pd.MultiIndex.from_tuples(tuples)
    return out


def render_table_object(
    object_id: str,
    workbook_path: str | Path | None = None,
) -> Any:
    from IPython.display import Latex

    model = _load_model(workbook_path)
    if object_id not in model.registry:
        placeholder = pd.DataFrame({"": [f"Table '{object_id}' not found in workbook."]})
        if is_pdf_render():
            return Latex(render_pdf_table_latex(object_id, placeholder))
        return style_cha_table(placeholder, has_multilevel_headers=False)

    record = model.registry[object_id]
    if record.object_type != "table":
        raise ValueError(f"Object '{object_id}' is not a table.")

    table_spec = model.table_specs[object_id]
    source_df = model.data_frames[record.data_sheet].copy()
    source_df = _prepare_table_df(source_df, table_spec.format_rules)
    source_df.columns = [_strip_format_tokens_from_label(col) for col in source_df.columns]
    # Some figure-oriented sheets use "County" as configured X metadata even when
    # the first column values are calendar years. Keep table headers semantic.
    if not source_df.empty:
        first_col = source_df.columns[0]
        first_col_text = _as_text(first_col).lower()
        if first_col_text == "county" and _is_time_like_column(source_df[first_col]):
            source_df = source_df.rename(columns={first_col: "Year"})
    if table_spec.has_multilevel_headers:
        source_df = _rebuild_multiindex(source_df)
    if is_pdf_render():
        return Latex(render_pdf_table_latex(object_id, source_df))
    return style_cha_table(source_df, has_multilevel_headers=table_spec.has_multilevel_headers)


def render_figure_object(
    figure_id: str,
    workbook_path: str | Path | None = None,
):
    model = _load_model(workbook_path)
    if figure_id not in model.registry:
        fig = go.Figure()
        fig.update_layout(
            title_text=f"Figure '{figure_id}' not found in workbook",
            xaxis=dict(visible=False),
            yaxis=dict(visible=False),
            annotations=[
                dict(
                    text=f"Figure '{figure_id}' not found in workbook.",
                    x=0.5,
                    y=0.5,
                    xref="paper",
                    yref="paper",
                    showarrow=False,
                )
            ],
        )
        return fig

    record = model.registry[figure_id]
    if record.object_type != "figure":
        raise ValueError(f"Object '{figure_id}' is not a figure.")

    spec = model.figure_specs[figure_id]
    df = model.data_frames[record.data_sheet].copy()
    df.columns = [_strip_format_tokens_from_label(col) for col in df.columns]
    df.columns = _ensure_unique_column_labels(list(df.columns))
    first_col = df.columns[0]
    first_col_is_time_like = _is_time_like_column(df[first_col])

    used_region_pivot = False
    # For bar charts with region columns, pivot so counties appear on the X-axis
    # when the row dimension is *not* time (e.g. one row per category) or there
    # is only one data row (single-period snapshot).  Multi-row year×county
    # wide tables stay as-is: years on the X axis, counties as grouped series.
    # (Excel often stores years as floats; first_col_is_time_like must catch that
    # or we would wrongly pivot — see _is_time_like_value.)
    is_bar_type = spec.figure_type in {"clustered_bar", "stacked_bar", "simple_bar", "horizontal_bar"}
    region_cols = _detect_region_columns(df) if is_bar_type else []
    allow_region_pivot = False
    if is_bar_type and region_cols:
        force_region_pivot = str(spec.x_col).strip().lower() == "county"
        if force_region_pivot:
            # Honor explicit metadata intent: grouped bars by year on county X-axis.
            allow_region_pivot = True
        elif len(df) <= 1:
            allow_region_pivot = True
        elif first_col_is_time_like:
            allow_region_pivot = False
        else:
            allow_region_pivot = True
    if allow_region_pivot:
        if region_cols:
            row_label_col = first_col
            # Preserve workbook category order (matches table row order) when
            # pivoting to county x-axis grouped bars.
            row_order = [
                value for value in dict.fromkeys(df[row_label_col].tolist())
                if not pd.isna(value)
            ]
            # Use a temp variable name to avoid collision when the first
            # column is already called "County".
            _region_var = "__region__"
            long_df = df[[row_label_col] + region_cols].melt(
                id_vars=[row_label_col], value_vars=region_cols,
                var_name=_region_var, value_name="__value",
            )
            long_df[_region_var] = long_df[_region_var].map(_normalize_region_label_for_axis)
            wide = long_df.pivot(index=_region_var, columns=row_label_col, values="__value")
            if row_order:
                ordered_categories = [label for label in row_order if label in wide.columns]
                if ordered_categories:
                    wide = wide.reindex(columns=ordered_categories)
            ordered = [region for region in CHA_REGION_ORDER if region in wide.index]
            if ordered:
                wide = wide.reindex(index=ordered)
            df = wide.reset_index().rename(columns={_region_var: "County"})
            df.columns = [str(c) for c in df.columns]
            x_col = "County"
            y_cols = [col for col in df.columns if col != x_col]
            x_axis_title = "County"
            used_region_pivot = True

    if not used_region_pivot:
        if spec.x_col in df.columns:
            x_col = spec.x_col
            needs_pivot = False
        elif spec.pivot_for_chart:
            if spec.figure_type == "line" and first_col_is_time_like:
                x_col = first_col
                needs_pivot = False
            else:
                x_col = spec.x_col or "Category"
                needs_pivot = True
        else:
            # Auto-detect wide single-row category data that needs pivoting:
            # all columns are non-region category names and x_col is absent.
            region_cols_check = _detect_region_columns(df)
            is_wide_category = (
                spec.x_col
                and spec.x_col not in df.columns
                and not region_cols_check
                and len(df) <= 2
                and spec.figure_type in {"simple_bar", "clustered_bar", "stacked_bar", "horizontal_bar"}
            )
            if is_wide_category:
                x_col = spec.x_col or "Category"
                needs_pivot = True
            else:
                x_col = first_col
                needs_pivot = False

        if needs_pivot:
            pivot_key = df.columns[0]
            if len(df) <= 1:
                # Single-row wide data: column names become X values.
                all_cols = list(df.columns)
                values = df.iloc[0].tolist() if len(df) == 1 else [None] * len(all_cols)
                df = pd.DataFrame({x_col: all_cols, "Value": values})
            else:
                df = df.set_index(pivot_key).T.reset_index().rename(columns={"index": x_col})

        y_cols = [col for col in spec.y_cols if col in df.columns and col != x_col]
        if not y_cols:
            y_cols = [col for col in df.columns if col != x_col]
        if spec.figure_type in {"simple_bar", "horizontal_bar"}:
            y_cols = y_cols[:1]
        x_axis_title = spec.x_axis_title or x_col
        if x_col == first_col and not str(spec.x_axis_title).strip() and first_col_is_time_like:
            x_axis_title = "Year"

    figure_rules = _figure_format_rules(model, figure_id)
    if figure_id == "fig-alice-threshold":
        desired = ["ALICE", "Above ALICE Threshold", "Poverty"]
        y_cols = [col for col in desired if col in y_cols] + [col for col in y_cols if col not in desired]

    for col in y_cols:
        if col not in df.columns:
            continue
        # Some workbook sheets can produce duplicate column labels; in that case
        # df[col] is a DataFrame. Use the first matching column deterministically.
        match_positions = np.where(df.columns == col)[0]
        if len(match_positions) == 0:
            continue
        col_idx = int(match_positions[0])
        series = df.iloc[:, col_idx]
        fmt = figure_rules.get(col, "")
        if fmt == "ratio":
            series = series.apply(_coerce_ratio_numeric)
        elif fmt == "date":
            continue
        else:
            # Strip annotation markers (*, **, †, etc.) before coercing.
            if series.dtype == object:
                series = series.astype(str).str.replace(
                    r"[*†‡§#]+$", "", regex=True
                )
            series = pd.to_numeric(series, errors="coerce")
        df.iloc[:, col_idx] = series

    y_axis_title = spec.y_axis_title or ("Percent" if str(spec.hover_suffix).strip() == "%" else "Value")

    if spec.figure_type == "line":
        return build_line_figure(
            df=df,
            x_col=x_col,
            y_cols=y_cols,
            x_axis_title=x_axis_title,
            y_axis_title=y_axis_title,
            start_at_zero=spec.start_at_zero,
            hover_value_format=".1f",
            hover_suffix=spec.hover_suffix,
        )
    if spec.figure_type == "clustered_bar":
        return build_clustered_bar_figure(
            df=df,
            x_col=x_col,
            y_cols=y_cols,
            x_axis_title=x_axis_title,
            y_axis_title=y_axis_title,
            start_at_zero=True,
            hover_value_format=".1f",
            hover_suffix=spec.hover_suffix,
            show_data_labels=_resolve_show_data_labels(spec),
        )
    if spec.figure_type == "stacked_bar":
        return build_stacked_bar_figure(
            df=df,
            x_col=x_col,
            y_cols=y_cols,
            x_axis_title=x_axis_title,
            y_axis_title=y_axis_title,
            start_at_zero=spec.start_at_zero,
            hover_value_format=".1f",
            hover_suffix=spec.hover_suffix,
            show_data_labels=_resolve_show_data_labels(spec),
        )
    if spec.figure_type == "simple_bar":
        return build_simple_bar_figure(
            df=df,
            x_col=x_col,
            y_cols=y_cols,
            x_axis_title=x_axis_title,
            y_axis_title=y_axis_title,
            start_at_zero=spec.start_at_zero,
            hover_value_format=".1f",
            hover_suffix=spec.hover_suffix,
            show_data_labels=_resolve_show_data_labels(spec),
        )
    if spec.figure_type == "horizontal_bar":
        return build_horizontal_bar_figure(
            df=df,
            x_col=x_col,
            y_cols=y_cols,
            x_axis_title=x_axis_title,
            y_axis_title=y_axis_title,
            start_at_zero=spec.start_at_zero,
            hover_value_format=".1f",
            hover_suffix=spec.hover_suffix,
        )

    raise ValueError(f"Unsupported figure_type '{spec.figure_type}' for '{figure_id}'.")


def render_source_callout_for_object(
    object_id: str,
    workbook_path: str | Path | None = None,
) -> str:
    model = _load_model(workbook_path)
    source_spec = model.source_specs.get(object_id)
    if not source_spec:
        return ""
    return create_source_callout(
        table_id=source_spec.table_id or None,
        url=source_spec.url or None,
        data_year=source_spec.data_year,
        estimate_type=source_spec.estimate_type,
        citation_month=source_spec.citation_month,
        citation_year=source_spec.citation_year,
        custom_text=source_spec.custom_text or None,
    )
