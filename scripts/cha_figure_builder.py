"""
CHA Figure Builder

Utilities for creating CHA figures (lines, clustered bars, stacked bars, simple bars,
horizontal bars, horizontal stacked/clustered bars, pie charts, and dot-whisker / CI
plots) with consistent styling and ordering.
Includes a helper to display a figure above its table output in Quarto/Jupyter.
"""

from __future__ import annotations

import warnings
import textwrap
import math

import pandas as pd
import plotly.graph_objects as go

from scripts.cha_table_styling import (
    CHA_REGION_ALIASES,
    CHA_REGION_ORDER,
    style_cha_table,
)


CHA_COLOR_PALETTE = [
    "#9ACD4B",
    "#FAA83B",
    "#D35840",
    "#9941B1",
    "#63A0CC",
    "#82FFFF",
    "#5E8425",
    "#EBE603",
    "#D91AC4",
    "#18AC93",
]

# Blue / orange pairing for 2022 vs 2025 (and similar) comparison dots.
DOT_WHISKER_PALETTE = [
    "#63A0CC",
    "#FAA83B",
    "#9ACD4B",
    "#D35840",
    "#9941B1",
]

DEFAULT_DASHED_SERIES = {"NYS": "dash", "US": "dash"}
BAR_PATTERN_SEQUENCE = ["", "/", "\\", "x", ".", "+"]
LINE_SYMBOL_SEQUENCE = ["circle", "square", "diamond", "triangle-up", "triangle-down", "cross"]
CHA_FONT_FAMILY = '"Tw Cen MT", "Tw Cen MT Std", "Century Gothic", "Trebuchet MS", "Segoe UI", sans-serif'


def _normalize_label(label: str) -> str:
    normalized = str(label).strip()
    return CHA_REGION_ALIASES.get(normalized, normalized)


def _ordered_series(series: list[str]) -> list[str]:
    normalized = {name: _normalize_label(name) for name in series}
    ordered = [name for name in series if normalized[name] in CHA_REGION_ORDER]
    ordered.sort(key=lambda name: CHA_REGION_ORDER.index(normalized[name]))
    remaining = [name for name in series if name not in ordered]
    return ordered + remaining


def _series_look_like_regions(series: list[str]) -> bool:
    """True when every series label maps to a known CHA region/county."""
    if not series:
        return False
    region_set = set(CHA_REGION_ORDER)
    return all(_normalize_label(name) in region_set for name in series)


def _series_colors(series: list[str], palette: list[str] | None = None) -> dict[str, str]:
    palette = palette or CHA_COLOR_PALETTE
    return {name: palette[idx % len(palette)] for idx, name in enumerate(series)}


def _series_dashes(series: list[str], overrides: dict[str, str] | None = None) -> dict[str, str]:
    dashes = {name: "solid" for name in series}
    for name, style in (overrides or DEFAULT_DASHED_SERIES).items():
        if name in dashes:
            dashes[name] = style
    return dashes


def _series_patterns(series: list[str]) -> dict[str, str]:
    return {name: BAR_PATTERN_SEQUENCE[idx % len(BAR_PATTERN_SEQUENCE)] for idx, name in enumerate(series)}


def _series_symbols(series: list[str]) -> dict[str, str]:
    return {name: LINE_SYMBOL_SEQUENCE[idx % len(LINE_SYMBOL_SEQUENCE)] for idx, name in enumerate(series)}


def _round_up_to_nice_number(value: float) -> float:
    """
    Round up to the next nice round number.
    Examples: 92 -> 100, 150 -> 200, 250 -> 300, 850 -> 900, 950 -> 1000
    """
    if value <= 0:
        return 100.0
    if value <= 100:
        return 100.0
    
    # For values > 100, round up to the next 100
    return ((int(value) // 100) + 1) * 100


def _y_range(values: pd.Series, start_at_zero: bool, padding: float, is_bar_graph: bool = False) -> list[float]:
    y_min = float(values.min())
    y_max = float(values.max())
    
    # Use original logic for all graph types (bar graphs and line graphs)
    if start_at_zero:
        y_min = 0.0
    y_span = y_max - y_min
    if y_span == 0:
        y_span = 1.0
    lower = y_min - y_span * padding
    upper = y_max + y_span * padding

    # Clustered/stacked bars should not dip below zero when zero-baselined.
    if is_bar_graph and start_at_zero:
        lower = 0.0

    return [lower, upper]


def _y_axis_tick_settings(
    values: pd.Series,
    axis_range: list[float] | tuple[float, float] | None = None,
) -> dict[str, object]:
    """
    Prefer whole-number y-axis ticks unless decimals are needed to show
    tight-range variation.
    """
    numeric = pd.to_numeric(values, errors="coerce").dropna()
    if numeric.empty:
        return {}

    def _nice_step(span: float, target_ticks: int = 12) -> float:
        if span <= 0:
            return 1.0
        raw = span / max(target_ticks, 1)
        magnitude = 10 ** math.floor(math.log10(raw))
        normalized = raw / magnitude
        if normalized <= 1:
            nice = 1
        elif normalized <= 2:
            nice = 2
        elif normalized <= 5:
            nice = 5
        else:
            nice = 10
        return nice * magnitude

    integer_like = (numeric.sub(numeric.round()).abs() < 1e-9).all()
    data_span = float(numeric.max() - numeric.min())
    span = data_span
    if axis_range and len(axis_range) == 2:
        axis_span = float(axis_range[1]) - float(axis_range[0])
        if axis_span > 0:
            span = axis_span

    if integer_like:
        step = max(1.0, _nice_step(span))
        if step == 10 and span <= 80:
            step = 5.0
        return {"dtick": step, "tickformat": ",.0f"}

    half_step_like = (numeric.mul(2).sub(numeric.mul(2).round()).abs() < 1e-9).all()
    # If values are all .0/.5, keep that granularity only for tighter ranges.
    if half_step_like and data_span <= 8:
        return {"dtick": 0.5, "tickformat": ",.1f"}

    step = _nice_step(span)
    if step == 10 and span <= 80:
        step = 5.0
    if step >= 1:
        # Broader decimal ranges still read best with whole-number axis labels.
        return {"dtick": step, "tickformat": ",.0f"}
    return {"dtick": step, "tickformat": ",.1f"}


def _coerce_year_axis(values: pd.Series | pd.DataFrame) -> tuple[pd.Series, dict[str, object]]:
    """
    Detect year-like X values and return numeric years + integer tick settings.
    """
    # Duplicate column labels can make df[x_col] return a DataFrame.
    # In that case, use the first matching column as the X-axis values.
    if isinstance(values, pd.DataFrame):
        if values.shape[1] == 0:
            return pd.Series(dtype="object"), {}
        if values.shape[1] > 1:
            warnings.warn(
                "Duplicate x-axis column labels detected; using the first matching column.",
                stacklevel=2,
            )
        values = values.iloc[:, 0]

    numeric = pd.to_numeric(values, errors="coerce")
    if numeric.isna().any():
        return values, {}

    # Treat as a year axis only when all values are whole-number years.
    is_integer_like = (numeric.sub(numeric.round()).abs() < 1e-9).all()
    is_year_range = ((numeric >= 1000) & (numeric <= 3000)).all()
    if not (is_integer_like and is_year_range):
        return values, {}

    coerced = numeric.round().astype(int)
    # Show only years present in the source data (e.g., 2016, 2018, 2021)
    # rather than every intermediate year on the axis.
    axis_options = dict(
        type="linear",
        tickmode="array",
        tickvals=coerced.drop_duplicates().tolist(),
        ticktext=coerced.drop_duplicates().astype(str).tolist(),
        tickformat="d",
    )
    return coerced, axis_options


def _wrap_tick_label(label: str, width: int) -> str:
    text = str(label).strip()
    if not text:
        return text
    if len(text) <= width:
        return text
    lines = textwrap.wrap(
        text,
        width=width,
        break_long_words=False,
        break_on_hyphens=False,
    )
    return "<br>".join(lines) if lines else text


def _format_categorical_tick_labels(labels: list[str], figure_width: int) -> tuple[list[str], int, int]:
    if not labels:
        return labels, 11, 120

    category_count = max(len(labels), 1)
    max_label_len = max(len(str(label)) for label in labels)

    # Approximate readable chars/line from figure width and category count.
    # Clamp to avoid overly narrow or overly wide wrapping behavior.
    chars_per_line = max(12, min(28, int(figure_width / (category_count * 7))))
    wrapped = [_wrap_tick_label(label, chars_per_line) for label in labels]
    max_lines = max(str(label).count("<br>") + 1 for label in wrapped)

    if category_count >= 12 or max_label_len >= 40:
        tick_font_size = 9
    elif category_count >= 9 or max_label_len >= 28:
        tick_font_size = 10
    else:
        tick_font_size = 11

    bottom_margin = 80 + ((max_lines - 1) * 16) + ((11 - tick_font_size) * 8)
    bottom_margin = max(100, min(220, bottom_margin))
    return wrapped, tick_font_size, bottom_margin


def _prepare_categorical_x_axis(
    x_values: pd.Series,
    figure_width: int,
) -> tuple[list[str], list[str], int, int]:
    raw_labels = x_values.astype(str).tolist()
    wrapped_labels, tick_font_size, bottom_margin = _format_categorical_tick_labels(
        raw_labels,
        figure_width,
    )
    return raw_labels, wrapped_labels, tick_font_size, bottom_margin


def _bar_data_label_settings(
    show_data_labels: bool,
    value_format: str,
    value_suffix: str,
    *,
    orientation: str = "v",
) -> dict[str, object]:
    if not show_data_labels:
        return {}
    value_token = "x" if orientation == "h" else "y"
    return {
        "texttemplate": f"%{{{value_token}:{value_format}}}{value_suffix}",
        "textposition": "outside",
        "cliponaxis": False,
    }


def _horizontal_category_height(n_categories: int, n_series: int = 1, *, grouped: bool = False) -> int:
    """Choose a readable figure height for horizontal multi-series bars."""
    per_category = 28 if not grouped else max(28, 18 * max(n_series, 1))
    return max(420, per_category * max(n_categories, 1) + 140)


def _bottom_horizontal_legend(*, y: float = -0.24, title: str | None = None) -> dict:
    """
    Shared bottom legend styling.

    ``y`` is kept low enough that the x-axis title (or an optional legend title)
    does not collide with the legend labels.
    """
    legend: dict = dict(
        orientation="h",
        x=0.5,
        y=y,
        xanchor="center",
        yanchor="top",
        font=dict(size=11),
        bgcolor="rgba(0,0,0,0)",
        borderwidth=0,
        # Extra gap between a legend title row and the swatch/label row.
        tracegroupgap=12,
    )
    if title is not None:
        # Extra blank line under the title — Plotly has no title-padding property.
        legend["title"] = dict(
            text=f"{title}<br><br>\u00a0" if str(title).strip() else "",
            font=dict(size=12),
            side="top",
        )
    return legend


def _apply_layout(
    fig: go.Figure,
    x_axis_title: str,
    y_axis_title: str,
    y_range: list[float] | None,
    width: int,
    height: int,
    font_family: str,
    is_bar_graph: bool = False,
) -> None:
    yaxis_dict = dict(
        title=dict(text=y_axis_title, font=dict(size=14, family=font_family)),
        showgrid=True,
        gridcolor="rgba(128, 128, 128, 0.2)",
        gridwidth=1,
    )
    
    if y_range:
        yaxis_dict["range"] = y_range
        if is_bar_graph:
            # For bar graphs, disable autorange to ensure the range stays fixed
            yaxis_dict["autorange"] = False
    
    fig.update_layout(
        font=dict(family=font_family),
        xaxis=dict(
            title=dict(
                text=x_axis_title,
                font=dict(size=14, family=font_family),
                standoff=22,
            ),
            showgrid=True,
            gridcolor="rgba(128, 128, 128, 0.2)",
            gridwidth=1,
        ),
        yaxis=yaxis_dict,
        hovermode="x unified",
        legend=dict(
            x=1.02,
            y=0.5,
            xanchor="left",
            yanchor="middle",
            font=dict(size=11),
            bgcolor="rgba(255, 255, 255, 0.9)",
            bordercolor="rgba(0, 0, 0, 0.2)",
            borderwidth=1,
        ),
        plot_bgcolor="white",
        paper_bgcolor="white",
        width=width,
        height=height,
        margin=dict(l=80, r=200, t=40, b=60),
    )


def build_line_figure(
    df: pd.DataFrame,
    x_col: str,
    y_cols: list[str] | None = None,
    *,
    x_axis_title: str | None = None,
    y_axis_title: str,
    start_at_zero: bool = False,
    y_padding: float = 0.1,
    palette: list[str] | None = None,
    dash_overrides: dict[str, str] | None = None,
    width: int = 1000,
    height: int = 600,
    font_family: str = CHA_FONT_FAMILY,
    hover_value_format: str = ".1f",
    hover_suffix: str = "",
) -> go.Figure:
    """Build a static time series line figure with years on the X-axis and counties as grouped lines."""
    return build_interactive_line_figure(
        df,
        x_col,
        y_cols=y_cols,
        x_axis_title=x_axis_title,
        y_axis_title=y_axis_title,
        start_at_zero=start_at_zero,
        y_padding=y_padding,
        palette=palette,
        dash_overrides=dash_overrides,
        width=width,
        height=height,
        font_family=font_family,
        hover_value_format=hover_value_format,
        hover_suffix=hover_suffix,
    )


def build_clustered_bar_figure(
    df: pd.DataFrame,
    x_col: str,
    y_cols: list[str] | None = None,
    *,
    x_axis_title: str | None = None,
    y_axis_title: str,
    start_at_zero: bool = True,
    y_padding: float = 0.1,
    palette: list[str] | None = None,
    width: int = 1000,
    height: int = 600,
    font_family: str = CHA_FONT_FAMILY,
    hover_value_format: str = ".1f",
    hover_suffix: str = "",
    show_data_labels: bool = False,
) -> go.Figure:
    series = y_cols or [col for col in df.columns if col != x_col]
    ordered = _ordered_series(series)
    colors = _series_colors(ordered, palette)
    patterns = _series_patterns(ordered)
    x_axis_title = x_axis_title or x_col
    value_label = y_axis_title or "Value"
    # When series are regions (Orange, NYS, …), the workbook often stores
    # "Region/County" as the X-axis title. Put it on the legend instead so it
    # labels the swatches and does not collide with them.
    legend_title: str | None = None
    layout_x_axis_title = x_axis_title
    if _series_look_like_regions(ordered) and str(x_axis_title).strip():
        legend_title = str(x_axis_title).strip()
        layout_x_axis_title = ""
    x_values_raw, x_values_wrapped, tick_font_size, bottom_margin = _prepare_categorical_x_axis(
        df[x_col],
        width,
    )

    fig = go.Figure()
    data_label_settings = _bar_data_label_settings(show_data_labels, hover_value_format, hover_suffix)
    for col in ordered:
        fig.add_trace(
            go.Bar(
                x=x_values_wrapped,
                y=df[col],
                name=col,
                marker=dict(
                    color=colors[col],
                    pattern=dict(shape=patterns[col], solidity=0.22),
                ),
                customdata=x_values_raw,
                hovertemplate=(
                    f"<b>{col}</b><br>{x_axis_title}: %{{customdata}}<br>"
                    f"{value_label}: %{{y:{hover_value_format}}}{hover_suffix}"
                    "<extra></extra>"
                ),
                **data_label_settings,
            )
        )

    y_values = df[ordered].to_numpy().flatten()
    y_series = pd.Series(y_values)
    y_range = _y_range(y_series, start_at_zero, y_padding, is_bar_graph=True)

    _apply_layout(
        fig=fig,
        x_axis_title=layout_x_axis_title,
        y_axis_title=y_axis_title,
        y_range=y_range,
        width=width,
        height=height,
        font_family=font_family,
        is_bar_graph=True,
    )
    y_tick_settings = _y_axis_tick_settings(y_series, y_range)
    fig.update_layout(
        barmode="group",
        bargap=0.30,
        bargroupgap=0.10,
        paper_bgcolor="white",
        plot_bgcolor="white",
        legend=_bottom_horizontal_legend(y=-0.22, title=legend_title),
        # Room for legend title + labels under the category ticks.
        margin=dict(l=80, r=40, t=40, b=min(300, bottom_margin + (110 if legend_title else 72))),
    )
    fig.update_yaxes(
        gridcolor="rgba(0, 0, 0, 0.15)",
        zerolinecolor="rgba(0, 0, 0, 0.2)",
        **y_tick_settings,
    )
    fig.update_xaxes(
        tickangle=0,
        automargin=True,
        tickfont=dict(size=tick_font_size),
        gridcolor="rgba(0, 0, 0, 0)",
        # Keep the axis title nearer the ticks so it does not crowd the legend.
        title_standoff=6,
    )
    return fig


def build_stacked_bar_figure(
    df: pd.DataFrame,
    x_col: str,
    y_cols: list[str] | None = None,
    *,
    x_axis_title: str | None = None,
    y_axis_title: str,
    start_at_zero: bool = True,
    y_padding: float = 0.1,
    palette: list[str] | None = None,
    width: int = 1000,
    height: int = 600,
    font_family: str = CHA_FONT_FAMILY,
    hover_value_format: str = ".1f",
    hover_suffix: str = "",
    show_data_labels: bool = False,
) -> go.Figure:
    series = y_cols or [col for col in df.columns if col != x_col]
    ordered = _ordered_series(series)
    colors = _series_colors(ordered, palette)
    patterns = _series_patterns(ordered)
    x_axis_title = x_axis_title or x_col
    value_label = y_axis_title or "Value"
    x_values_raw, x_values_wrapped, tick_font_size, bottom_margin = _prepare_categorical_x_axis(
        df[x_col],
        width,
    )

    fig = go.Figure()
    data_label_settings = _bar_data_label_settings(show_data_labels, hover_value_format, hover_suffix)
    for col in ordered:
        fig.add_trace(
            go.Bar(
                x=x_values_wrapped,
                y=df[col],
                name=col,
                marker=dict(
                    color=colors[col],
                    pattern=dict(shape=patterns[col], solidity=0.22),
                ),
                customdata=x_values_raw,
                hovertemplate=(
                    f"<b>{col}</b><br>{x_axis_title}: %{{customdata}}<br>"
                    f"{value_label}: %{{y:{hover_value_format}}}{hover_suffix}"
                    "<extra></extra>"
                ),
                **data_label_settings,
            )
        )

    totals = df[ordered].sum(axis=1)
    y_series = pd.Series(totals)
    y_range = _y_range(y_series, start_at_zero, y_padding, is_bar_graph=True)
    _apply_layout(
        fig=fig,
        x_axis_title=x_axis_title,
        y_axis_title=y_axis_title,
        y_range=y_range,
        width=width,
        height=height,
        font_family=font_family,
        is_bar_graph=True,
    )
    fig.update_layout(
        barmode="stack",
        margin=dict(l=80, r=200, t=40, b=min(320, bottom_margin + 20)),
    )
    fig.update_yaxes(**_y_axis_tick_settings(y_series, y_range))
    fig.update_xaxes(
        tickangle=0,
        automargin=True,
        tickfont=dict(size=tick_font_size),
    )
    return fig


def build_simple_bar_figure(
    df: pd.DataFrame,
    x_col: str,
    y_cols: list[str] | None = None,
    *,
    x_axis_title: str | None = None,
    y_axis_title: str,
    start_at_zero: bool = True,
    y_padding: float = 0.1,
    palette: list[str] | None = None,
    width: int = 1000,
    height: int = 600,
    font_family: str = CHA_FONT_FAMILY,
    hover_value_format: str = ".1f",
    hover_suffix: str = "",
    show_data_labels: bool = False,
) -> go.Figure:
    """
    Build a single-series bar chart.
    """
    series = y_cols or [col for col in df.columns if col != x_col]
    if not series:
        raise ValueError("No y-axis series found for simple bar chart.")
    y_col = series[0]
    x_axis_title = x_axis_title or x_col
    value_label = y_axis_title or "Value"
    color = (palette or CHA_COLOR_PALETTE)[0]

    x_values_raw, x_values_wrapped, tick_font_size, bottom_margin = _prepare_categorical_x_axis(
        df[x_col],
        width,
    )

    fig = go.Figure()
    data_label_settings = _bar_data_label_settings(show_data_labels, hover_value_format, hover_suffix)
    fig.add_trace(
        go.Bar(
            x=x_values_wrapped,
            y=df[y_col],
            name=y_col,
            marker=dict(color=color),
            customdata=x_values_raw,
            hovertemplate=(
                f"<b>{y_col}</b><br>{x_axis_title}: %{{customdata}}<br>"
                f"{value_label}: %{{y:{hover_value_format}}}{hover_suffix}"
                "<extra></extra>"
            ),
            **data_label_settings,
        )
    )

    y_values = pd.to_numeric(df[y_col], errors="coerce")
    y_series = pd.Series(y_values)
    y_range = _y_range(y_series, start_at_zero, y_padding, is_bar_graph=True)
    _apply_layout(
        fig=fig,
        x_axis_title=x_axis_title,
        y_axis_title=y_axis_title,
        y_range=y_range,
        width=width,
        height=height,
        font_family=font_family,
        is_bar_graph=True,
    )
    fig.update_layout(
        barmode="group",
        margin=dict(l=80, r=40, t=40, b=min(300, bottom_margin + 20)),
    )
    fig.update_yaxes(**_y_axis_tick_settings(y_series, y_range))
    fig.update_xaxes(
        tickangle=0,
        automargin=True,
        tickfont=dict(size=tick_font_size),
        gridcolor="rgba(0, 0, 0, 0)",
    )
    return fig


def build_horizontal_bar_figure(
    df: pd.DataFrame,
    x_col: str,
    y_cols: list[str] | None = None,
    *,
    x_axis_title: str | None = None,
    y_axis_title: str,
    start_at_zero: bool = True,
    y_padding: float = 0.1,
    palette: list[str] | None = None,
    width: int = 1000,
    height: int = 600,
    font_family: str = CHA_FONT_FAMILY,
    hover_value_format: str = ".1f",
    hover_suffix: str = "",
) -> go.Figure:
    """
    Build a single-series horizontal bar chart.
    """
    series = y_cols or [col for col in df.columns if col != x_col]
    if not series:
        raise ValueError("No y-axis series found for horizontal bar chart.")
    value_col = series[0]
    category_axis_title = x_axis_title or x_col
    value_axis_title = y_axis_title or "Value"
    color = (palette or CHA_COLOR_PALETTE)[0]

    plot_df = df[[x_col, value_col]].copy()
    plot_df[value_col] = pd.to_numeric(plot_df[value_col], errors="coerce")
    plot_df = plot_df.dropna(subset=[value_col, x_col])
    plot_df = plot_df.sort_values(value_col, ascending=True)

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=plot_df[value_col],
            y=plot_df[x_col].astype(str),
            orientation="h",
            marker=dict(color=color),
            customdata=plot_df[x_col].astype(str),
            hovertemplate=(
                f"{category_axis_title}: %{{customdata}}<br>"
                f"{value_axis_title}: %{{x:{hover_value_format}}}{hover_suffix}"
                "<extra></extra>"
            ),
        )
    )

    x_range = _y_range(plot_df[value_col], start_at_zero, y_padding, is_bar_graph=True)
    fig.update_layout(
        showlegend=False,
        paper_bgcolor="white",
        plot_bgcolor="white",
        width=width,
        height=height,
        margin=dict(l=250, r=40, t=40, b=70),
        font=dict(family=font_family),
    )
    fig.update_xaxes(
        title=dict(text=value_axis_title, font=dict(size=14, family=font_family)),
        range=x_range,
        gridcolor="rgba(0, 0, 0, 0.15)",
        zerolinecolor="rgba(0, 0, 0, 0.2)",
        **_y_axis_tick_settings(plot_df[value_col], x_range),
    )
    fig.update_yaxes(
        title=dict(text=category_axis_title, font=dict(size=14, family=font_family)),
        showgrid=False,
        ticklabelstandoff=10,
        automargin=True,
    )
    return fig


def build_horizontal_stacked_bar_figure(
    df: pd.DataFrame,
    x_col: str,
    y_cols: list[str] | None = None,
    *,
    x_axis_title: str | None = None,
    y_axis_title: str,
    start_at_zero: bool = True,
    y_padding: float = 0.1,
    palette: list[str] | None = None,
    width: int = 1000,
    height: int | None = None,
    font_family: str = CHA_FONT_FAMILY,
    hover_value_format: str = ".1f",
    hover_suffix: str = "",
    show_data_labels: bool = False,
) -> go.Figure:
    """
    Build a horizontal stacked bar chart (categories on Y, stacked values on X).
    """
    series = y_cols or [col for col in df.columns if col != x_col]
    ordered = _ordered_series(series)
    if not ordered:
        raise ValueError("No value series found for horizontal stacked bar chart.")
    colors = _series_colors(ordered, palette)
    patterns = _series_patterns(ordered)
    category_axis_title = x_axis_title or x_col
    value_axis_title = y_axis_title or "Value"

    plot_df = df.copy()
    categories = plot_df[x_col].astype(str).tolist()
    # First workbook row at top (matches survey reading order).
    categories_plot = list(reversed(categories))
    plot_df = plot_df.iloc[::-1].reset_index(drop=True)

    fig = go.Figure()
    data_label_settings = _bar_data_label_settings(
        show_data_labels,
        hover_value_format,
        hover_suffix,
        orientation="h",
    )
    for col in ordered:
        fig.add_trace(
            go.Bar(
                x=plot_df[col],
                y=categories_plot,
                orientation="h",
                name=col,
                marker=dict(
                    color=colors[col],
                    pattern=dict(shape=patterns[col], solidity=0.22),
                ),
                customdata=categories_plot,
                hovertemplate=(
                    f"<b>{col}</b><br>{category_axis_title}: %{{customdata}}<br>"
                    f"{value_axis_title}: %{{x:{hover_value_format}}}{hover_suffix}"
                    "<extra></extra>"
                ),
                **data_label_settings,
            )
        )

    totals = plot_df[ordered].sum(axis=1)
    x_series = pd.Series(totals)
    x_range = _y_range(x_series, start_at_zero, y_padding, is_bar_graph=True)
    fig_height = height if height is not None else _horizontal_category_height(len(categories))

    max_label_len = max((len(c) for c in categories), default=10)
    left_margin = 280 if max_label_len > 40 else 220
    fig.update_layout(
        barmode="stack",
        paper_bgcolor="white",
        plot_bgcolor="white",
        width=width,
        height=fig_height,
        margin=dict(l=left_margin, r=40, t=40, b=110),
        font=dict(family=font_family),
        legend=_bottom_horizontal_legend(y=-0.18),
        hovermode="y unified",
    )
    fig.update_xaxes(
        title=dict(text=value_axis_title, font=dict(size=14, family=font_family), standoff=8),
        range=x_range,
        gridcolor="rgba(0, 0, 0, 0.15)",
        zerolinecolor="rgba(0, 0, 0, 0.2)",
        **_y_axis_tick_settings(x_series, x_range),
    )
    fig.update_yaxes(
        title=dict(text=category_axis_title, font=dict(size=14, family=font_family)),
        showgrid=False,
        ticklabelstandoff=10,
        automargin=True,
    )
    return fig


def build_horizontal_clustered_bar_figure(
    df: pd.DataFrame,
    x_col: str,
    y_cols: list[str] | None = None,
    *,
    x_axis_title: str | None = None,
    y_axis_title: str,
    start_at_zero: bool = True,
    y_padding: float = 0.1,
    palette: list[str] | None = None,
    width: int = 1000,
    height: int | None = None,
    font_family: str = CHA_FONT_FAMILY,
    hover_value_format: str = ".1f",
    hover_suffix: str = "",
    show_data_labels: bool = False,
) -> go.Figure:
    """
    Build a horizontal clustered (grouped) bar chart (categories on Y, grouped values on X).
    """
    series = y_cols or [col for col in df.columns if col != x_col]
    ordered = _ordered_series(series)
    if not ordered:
        raise ValueError("No value series found for horizontal clustered bar chart.")
    colors = _series_colors(ordered, palette)
    patterns = _series_patterns(ordered)
    category_axis_title = x_axis_title or x_col
    value_axis_title = y_axis_title or "Value"

    plot_df = df.copy()
    categories = plot_df[x_col].astype(str).tolist()
    categories_plot = list(reversed(categories))
    plot_df = plot_df.iloc[::-1].reset_index(drop=True)

    fig = go.Figure()
    data_label_settings = _bar_data_label_settings(
        show_data_labels,
        hover_value_format,
        hover_suffix,
        orientation="h",
    )
    for col in ordered:
        fig.add_trace(
            go.Bar(
                x=plot_df[col],
                y=categories_plot,
                orientation="h",
                name=col,
                marker=dict(
                    color=colors[col],
                    pattern=dict(shape=patterns[col], solidity=0.22),
                ),
                customdata=categories_plot,
                hovertemplate=(
                    f"<b>{col}</b><br>{category_axis_title}: %{{customdata}}<br>"
                    f"{value_axis_title}: %{{x:{hover_value_format}}}{hover_suffix}"
                    "<extra></extra>"
                ),
                **data_label_settings,
            )
        )

    y_values = plot_df[ordered].to_numpy().flatten()
    x_series = pd.Series(y_values)
    x_range = _y_range(x_series, start_at_zero, y_padding, is_bar_graph=True)
    fig_height = height if height is not None else _horizontal_category_height(
        len(categories),
        len(ordered),
        grouped=True,
    )

    max_label_len = max((len(c) for c in categories), default=10)
    left_margin = 280 if max_label_len > 40 else 220
    fig.update_layout(
        barmode="group",
        bargap=0.25,
        bargroupgap=0.08,
        paper_bgcolor="white",
        plot_bgcolor="white",
        width=width,
        height=fig_height,
        margin=dict(l=left_margin, r=40, t=40, b=110),
        font=dict(family=font_family),
        legend=_bottom_horizontal_legend(y=-0.18),
        hovermode="y unified",
    )
    fig.update_xaxes(
        title=dict(text=value_axis_title, font=dict(size=14, family=font_family), standoff=8),
        range=x_range,
        gridcolor="rgba(0, 0, 0, 0.15)",
        zerolinecolor="rgba(0, 0, 0, 0.2)",
        **_y_axis_tick_settings(x_series, x_range),
    )
    fig.update_yaxes(
        title=dict(text=category_axis_title, font=dict(size=14, family=font_family)),
        showgrid=False,
        ticklabelstandoff=10,
        automargin=True,
    )
    return fig


def build_horizontal_color_bar_figure(
    df: pd.DataFrame,
    x_col: str,
    y_cols: list[str] | None = None,
    *,
    color_by: str,
    x_axis_title: str | None = None,
    y_axis_title: str,
    start_at_zero: bool = True,
    y_padding: float = 0.1,
    palette: list[str] | None = None,
    width: int = 1000,
    height: int | None = None,
    font_family: str = CHA_FONT_FAMILY,
    hover_value_format: str = ".0f",
    hover_suffix: str = "",
    show_data_labels: bool = True,
) -> go.Figure:
    """
    Build a horizontal bar chart colored by a grouping column.

    Long-format input: one row per category (x_col), a numeric value column, and
    a color_by group. Row order is preserved (first workbook row at the top), with
    small spacer gaps between color groups — matching Excel color-clustered bars.
    """
    series = y_cols or [col for col in df.columns if col not in {x_col, color_by}]
    if not series:
        raise ValueError("No value series found for horizontal color bar chart.")
    if color_by not in df.columns:
        raise ValueError(f"Color-by column '{color_by}' not found in data.")
    value_col = series[0]
    category_axis_title = x_axis_title if x_axis_title is not None else ""
    value_axis_title = y_axis_title or "Value"

    plot_df = df[[x_col, color_by, value_col]].copy()
    plot_df[value_col] = pd.to_numeric(plot_df[value_col], errors="coerce")
    plot_df[x_col] = plot_df[x_col].astype(str)
    plot_df[color_by] = plot_df[color_by].astype(str)
    plot_df = plot_df.dropna(subset=[value_col, x_col, color_by])
    if plot_df.empty:
        raise ValueError("No plottable rows for horizontal color bar chart.")

    # Preserve workbook order; insert blank spacer categories between groups.
    ordered_rows: list[dict[str, object]] = []
    category_order_top_first: list[str] = []
    prev_group: str | None = None
    spacer_idx = 0
    for row in plot_df.itertuples(index=False):
        category, group, value = row[0], row[1], row[2]
        if prev_group is not None and group != prev_group:
            spacer = f"\u200b{spacer_idx}"
            spacer_idx += 1
            category_order_top_first.append(spacer)
            ordered_rows.append(
                {x_col: spacer, color_by: prev_group, value_col: None, "__spacer__": True}
            )
        category_order_top_first.append(str(category))
        ordered_rows.append(
            {x_col: str(category), color_by: str(group), value_col: value, "__spacer__": False}
        )
        prev_group = str(group)

    long_df = pd.DataFrame(ordered_rows)
    # Plotly draws the first categoryarray entry at the bottom.
    category_array = list(reversed(category_order_top_first))

    group_order = list(dict.fromkeys(plot_df[color_by].tolist()))
    colors = _series_colors(group_order, palette)
    # Match Excel CHA export: "Other" is black.
    for group in group_order:
        if group.strip().lower() == "other":
            colors[group] = "#1A1A1A"

    fig = go.Figure()
    data_label_settings = _bar_data_label_settings(
        show_data_labels,
        hover_value_format,
        hover_suffix,
        orientation="h",
    )
    # Prefer end-of-bar labels like the Excel reference.
    if show_data_labels:
        data_label_settings = {
            **data_label_settings,
            "textposition": "inside",
            "insidetextanchor": "end",
            "textfont": dict(color="white", size=11, family=font_family),
            "cliponaxis": False,
        }

    for group in group_order:
        subset = long_df[long_df[color_by] == group].copy()
        # Keep spacers belonging to the previous group out of value traces.
        subset = subset[subset["__spacer__"] == False]  # noqa: E712
        fig.add_trace(
            go.Bar(
                x=subset[value_col],
                y=subset[x_col],
                orientation="h",
                name=group,
                marker=dict(color=colors[group]),
                customdata=subset[x_col],
                hovertemplate=(
                    f"<b>%{{customdata}}</b><br>"
                    f"{group}<br>"
                    f"{value_axis_title}: %{{x:{hover_value_format}}}{hover_suffix}"
                    "<extra></extra>"
                ),
                **data_label_settings,
            )
        )

    x_series = pd.to_numeric(plot_df[value_col], errors="coerce").dropna()
    x_range = _y_range(x_series, start_at_zero, y_padding, is_bar_graph=True)
    n_categories = len(category_order_top_first)
    fig_height = height if height is not None else _horizontal_category_height(n_categories)

    max_label_len = max((len(c) for c in plot_df[x_col].astype(str)), default=10)
    left_margin = 300 if max_label_len > 40 else 240
    fig.update_layout(
        barmode="overlay",
        bargap=0.12,
        paper_bgcolor="white",
        plot_bgcolor="white",
        width=width,
        height=fig_height,
        margin=dict(l=left_margin, r=40, t=40, b=130),
        font=dict(family=font_family),
        legend=_bottom_horizontal_legend(y=-0.18, title=""),
        hovermode="closest",
    )
    fig.update_xaxes(
        title=dict(text=value_axis_title, font=dict(size=14, family=font_family), standoff=8),
        range=x_range,
        gridcolor="rgba(0, 0, 0, 0.15)",
        zerolinecolor="rgba(0, 0, 0, 0.2)",
        **_y_axis_tick_settings(x_series, x_range),
    )
    # Hide spacer tick labels while preserving gap categories.
    ticktext = ["" if c.startswith("\u200b") else c for c in category_array]
    fig.update_yaxes(
        title=dict(text=category_axis_title, font=dict(size=14, family=font_family)),
        showgrid=False,
        ticklabelstandoff=10,
        automargin=True,
        categoryorder="array",
        categoryarray=category_array,
        tickmode="array",
        tickvals=category_array,
        ticktext=ticktext,
    )
    return fig


def build_pie_figure(
    df: pd.DataFrame,
    x_col: str,
    y_cols: list[str] | None = None,
    *,
    x_axis_title: str | None = None,
    y_axis_title: str | None = None,
    start_at_zero: bool = True,
    y_padding: float = 0.1,
    palette: list[str] | None = None,
    width: int = 900,
    height: int = 600,
    font_family: str = CHA_FONT_FAMILY,
    hover_value_format: str = ".1f",
    hover_suffix: str = "",
    show_data_labels: bool = True,
) -> go.Figure:
    """
    Build a pie chart from a category column and a single value series.
    """
    del start_at_zero, y_padding  # Unused; kept for a shared renderer call signature.
    series = y_cols or [col for col in df.columns if col != x_col]
    if not series:
        raise ValueError("No value series found for pie chart.")
    value_col = series[0]
    category_title = x_axis_title or x_col
    value_title = y_axis_title or value_col or "Value"

    plot_df = df[[x_col, value_col]].copy()
    plot_df[value_col] = pd.to_numeric(plot_df[value_col], errors="coerce")
    plot_df = plot_df.dropna(subset=[value_col, x_col])
    labels = plot_df[x_col].astype(str).tolist()
    values = plot_df[value_col].tolist()
    colors = (palette or CHA_COLOR_PALETTE)[: max(len(labels), 1)]
    # Cycle palette when there are more slices than colors.
    while len(colors) < len(labels):
        colors.extend(palette or CHA_COLOR_PALETTE)
    colors = colors[: len(labels)]

    textinfo = "percent" if show_data_labels else "none"
    fig = go.Figure(
        data=[
            go.Pie(
                labels=labels,
                values=values,
                sort=False,
                direction="clockwise",
                hole=0.0,
                marker=dict(
                    colors=colors,
                    line=dict(color="white", width=1.5),
                ),
                textinfo=textinfo,
                textposition="outside",
                textfont=dict(size=12, family=font_family),
                hovertemplate=(
                    f"<b>%{{label}}</b><br>"
                    f"{value_title}: %{{value:{hover_value_format}}}{hover_suffix}<br>"
                    f"Share: %{{percent}}"
                    "<extra></extra>"
                ),
            )
        ]
    )
    fig.update_layout(
        showlegend=True,
        legend=dict(
            orientation="v",
            x=1.02,
            y=0.5,
            xanchor="left",
            yanchor="middle",
            font=dict(size=11),
            bgcolor="rgba(255, 255, 255, 0.9)",
            bordercolor="rgba(0, 0, 0, 0.2)",
            borderwidth=1,
            # Extra blank line under the title — Plotly has no title-padding property.
            title=dict(
                text=f"{category_title}<br><br>\u00a0" if str(category_title).strip() else "",
                font=dict(size=12),
                side="top",
            ),
            tracegroupgap=12,
        ),
        paper_bgcolor="white",
        plot_bgcolor="white",
        width=width,
        height=height,
        margin=dict(l=40, r=200, t=40, b=40),
        font=dict(family=font_family),
    )
    return fig


def _resolve_ci_column(columns: list[str], series_name: str) -> str | None:
    """Find a confidence-interval column paired with a value series."""
    candidates = [
        f"{series_name} CI",
        f"{series_name} (±)",
        f"{series_name} (+/-)",
        f"{series_name}_ci",
        f"{series_name} ci",
    ]
    lower_map = {str(col).strip().lower(): col for col in columns}
    for candidate in candidates:
        match = lower_map.get(candidate.lower())
        if match is not None:
            return match
    return None


def _is_ci_column_label(label: str) -> bool:
    text = str(label).strip().lower()
    return (
        text.endswith(" ci")
        or text.endswith("_ci")
        or "(±)" in str(label)
        or "(+/-)" in str(label)
    )


def build_dot_whisker_figure(
    df: pd.DataFrame,
    x_col: str,
    y_cols: list[str] | None = None,
    *,
    x_axis_title: str | None = None,
    y_axis_title: str,
    start_at_zero: bool = True,
    y_padding: float = 0.1,
    palette: list[str] | None = None,
    width: int = 1000,
    height: int | None = None,
    font_family: str = CHA_FONT_FAMILY,
    hover_value_format: str = ".1f",
    hover_suffix: str = "",
) -> go.Figure:
    """
    Build a horizontal dot plot with whiskers (error bars / CIs).

    Expected columns: category column plus value columns (e.g. 2022, 2025),
    optionally paired with CI half-widths named ``{series} CI``.
    Categories keep workbook row order (first row at bottom, matching matplotlib).
    """
    if y_cols:
        value_cols = [col for col in y_cols if col in df.columns and col != x_col]
    else:
        value_cols = [col for col in df.columns if col != x_col]
    value_cols = [col for col in value_cols if not _is_ci_column_label(col)]
    if not value_cols:
        raise ValueError("No value series found for dot whisker chart.")

    category_axis_title = x_axis_title or x_col
    value_axis_title = y_axis_title or "Percent"
    colors = palette or DOT_WHISKER_PALETTE
    symbols = _series_symbols(value_cols)

    plot_df = df.copy()
    plot_df[x_col] = plot_df[x_col].astype(str)
    categories = plot_df[x_col].tolist()
    category_index = {cat: i for i, cat in enumerate(categories)}
    n_cats = max(len(categories), 1)
    fig_height = height if height is not None else max(480, 28 * n_cats + 120)

    # Slight vertical offsets so overlapping year markers remain readable.
    offsets = {
        name: (idx - (len(value_cols) - 1) / 2) * 0.18
        for idx, name in enumerate(value_cols)
    }

    fig = go.Figure()
    all_values: list[float] = []

    for idx, series_name in enumerate(value_cols):
        values = pd.to_numeric(plot_df[series_name], errors="coerce")
        ci_col = _resolve_ci_column(list(plot_df.columns), series_name)
        if ci_col is not None:
            ci = pd.to_numeric(plot_df[ci_col], errors="coerce").fillna(0.0)
        else:
            ci = pd.Series(0.0, index=plot_df.index)

        valid = values.notna()
        x_vals = values[valid]
        y_positions = [
            category_index[cat] + offsets[series_name]
            for cat in plot_df.loc[valid, x_col]
        ]
        ci_vals = ci[valid]
        all_values.extend((x_vals - ci_vals).tolist())
        all_values.extend((x_vals + ci_vals).tolist())

        lower = (x_vals - ci_vals).round(1)
        upper = (x_vals + ci_vals).round(1)
        custom = list(
            zip(
                plot_df.loc[valid, x_col].astype(str),
                lower.tolist(),
                upper.tolist(),
            )
        )

        fig.add_trace(
            go.Scatter(
                x=x_vals,
                y=y_positions,
                mode="markers",
                name=str(series_name),
                marker=dict(
                    color=colors[idx % len(colors)],
                    size=10,
                    symbol=symbols[series_name],
                    line=dict(width=0.5, color="rgba(0,0,0,0.35)"),
                ),
                error_x=dict(
                    type="data",
                    array=ci_vals,
                    visible=True,
                    color=colors[idx % len(colors)],
                    thickness=1.5,
                    width=4,
                ),
                customdata=custom,
                hovertemplate=(
                    f"<b>{series_name}</b><br>"
                    f"{category_axis_title}: %{{customdata[0]}}<br>"
                    f"{value_axis_title}: %{{x:{hover_value_format}}}{hover_suffix}<br>"
                    f"95% CI: %{{customdata[1]:{hover_value_format}}}–"
                    f"%{{customdata[2]:{hover_value_format}}}{hover_suffix}"
                    "<extra></extra>"
                ),
            )
        )

    value_series = pd.Series(all_values, dtype=float).dropna()
    if value_series.empty:
        value_series = pd.Series([0.0, 1.0])
    x_range = _y_range(value_series, start_at_zero, y_padding, is_bar_graph=True)

    left_margin = 280 if max((len(c) for c in categories), default=10) > 40 else 220
    fig.update_layout(
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        paper_bgcolor="white",
        plot_bgcolor="white",
        width=width,
        height=fig_height,
        margin=dict(l=left_margin, r=40, t=60, b=70),
        font=dict(family=font_family),
    )
    fig.update_xaxes(
        title=dict(text=value_axis_title, font=dict(size=14, family=font_family)),
        range=x_range,
        gridcolor="rgba(0, 0, 0, 0.15)",
        zerolinecolor="rgba(0, 0, 0, 0.2)",
        **_y_axis_tick_settings(value_series, x_range),
    )
    fig.update_yaxes(
        title=dict(text="", font=dict(size=14, family=font_family)),
        tickmode="array",
        tickvals=list(range(n_cats)),
        ticktext=categories,
        range=[-0.6, n_cats - 0.4],
        showgrid=True,
        gridcolor="rgba(0, 0, 0, 0.08)",
        zeroline=False,
        automargin=True,
    )
    return fig


def build_interactive_line_figure(
    df: pd.DataFrame,
    x_col: str,
    y_cols: list[str] | None = None,
    *,
    x_axis_title: str | None = None,
    y_axis_title: str,
    start_at_zero: bool = False,
    y_padding: float = 0.1,
    palette: list[str] | None = None,
    dash_overrides: dict[str, str] | None = None,
    width: int = 1000,
    height: int = 600,
    font_family: str = CHA_FONT_FAMILY,
    hover_value_format: str = ".1f",
    hover_suffix: str = "",
) -> go.Figure:
    """
    Build a static time series line figure.

    Years appear on the X-axis; each county/region is a separate colored line.
    """
    series = y_cols or [col for col in df.columns if col != x_col]
    ordered = _ordered_series(series)
    colors = _series_colors(ordered, palette)
    dashes = _series_dashes(ordered, dash_overrides)
    symbols = _series_symbols(ordered)
    x_axis_title = x_axis_title or x_col
    value_label = y_axis_title or "Value"
    x_values, x_axis_options = _coerce_year_axis(df[x_col])
    is_categorical_x = not bool(x_axis_options)

    x_values_raw: list[str] | None = None
    tick_font_size = 11
    bottom_margin = 60
    plotted_x_values = x_values
    hover_x_template = "%{x}"
    if is_categorical_x:
        x_values_raw, x_values_wrapped, tick_font_size, bottom_margin = _prepare_categorical_x_axis(
            pd.Series(x_values),
            width,
        )
        plotted_x_values = x_values_wrapped
        hover_x_template = "%{customdata}"

    fig = go.Figure()

    for col in ordered:
        customdata = x_values_raw if is_categorical_x else None
        fig.add_trace(
            go.Scatter(
                x=plotted_x_values,
                y=df[col],
                mode="lines+markers",
                name=col,
                customdata=customdata,
                line=dict(
                    color=colors[col],
                    width=3 if col in ["NYS", "US"] else 2.5,
                    dash=dashes[col],
                ),
                marker=dict(
                    size=8,
                    symbol=symbols[col],
                    color=colors[col],
                    line=dict(width=1.5, color="white"),
                ),
                hovertemplate=(
                    f"<b>{col}</b><br>{x_axis_title}: {hover_x_template}<br>"
                    f"{value_label}: %{{y:{hover_value_format}}}{hover_suffix}"
                    "<extra></extra>"
                ),
            )
        )

    y_values = df[ordered].to_numpy().flatten()
    y_series = pd.Series(y_values)
    y_range = _y_range(y_series, start_at_zero, y_padding)

    _apply_layout(
        fig=fig,
        x_axis_title=x_axis_title,
        y_axis_title=y_axis_title,
        y_range=y_range,
        width=width,
        height=height,
        font_family=font_family,
    )
    fig.update_yaxes(**_y_axis_tick_settings(y_series, y_range))
    if x_axis_options:
        fig.update_xaxes(**x_axis_options)
    elif is_categorical_x:
        fig.update_layout(margin=dict(l=80, r=200, t=40, b=bottom_margin))
        fig.update_xaxes(
            tickangle=0,
            automargin=True,
            tickfont=dict(size=tick_font_size),
        )

    return fig


def render_figure_and_table(
    fig: go.Figure,
    df: pd.DataFrame,
    *,
    has_multilevel_headers: bool = False,
    data_type: str | None = None,
    row_label_col: str | None = None,
) -> None:
    """
    Display a figure above its table in Quarto/Jupyter output.
    Use chunk options for figure and table captions/labels.

    Parameters
    ----------
    fig : go.Figure
        The Plotly figure to display.
    df : pd.DataFrame
        The table data.
    has_multilevel_headers : bool, optional
        Passed through to ``style_cha_table``.
    data_type : str, optional
        Plain-English data type label (e.g. ``"Percent"``, ``"Rate per 100,000"``).
        Controls number formatting in the table.  See ``style_cha_table`` for
        accepted values.
    row_label_col : str, optional
        Name of the row-label column (first column).  Not formatted as a number.
    """
    from IPython.display import display

    fig.show()
    styled_table = style_cha_table(
        df,
        has_multilevel_headers=has_multilevel_headers,
        data_type=data_type,
        row_label_col=row_label_col,
    )
    display(styled_table)
