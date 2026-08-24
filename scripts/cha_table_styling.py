"""
CHA Table Styling Module

This module provides standardized table styling functions for the Community Health Assessment.
Use this module to ensure consistent table formatting across all chapters.

Usage:
    from scripts.cha_table_styling import style_cha_table
    
    df = pd.DataFrame(your_data)
    styled_table = style_cha_table(df)
    styled_table  # Display in Quarto
"""

import html

import pandas as pd

CHA_FONT_FAMILY = '"Tw Cen MT", "Tw Cen MT Std", "Century Gothic", "Trebuchet MS", "Segoe UI", sans-serif'
# Approximate one Excel indent level (~3 character widths at 14px).
EXCEL_INDENT_PX_PER_LEVEL = 27
# Base cell padding; hierarchy indent is added on top via padding-left.
CELL_PADDING_PX = 10

_EXCEL_H_ALIGN_TO_CSS = {
    "left": "left",
    "center": "center",
    "right": "right",
    "justify": "justify",
    "centercontinuous": "center",
    "distributed": "justify",
    "fill": "left",
}
_EXCEL_V_ALIGN_TO_CSS = {
    "top": "top",
    "center": "middle",
    "bottom": "bottom",
    "justify": "middle",
    "distributed": "middle",
}


def css_text_align_from_excel(horizontal: str, *, default: str) -> str:
    """Map Excel horizontal alignment to CSS; ``default`` when unset."""
    token = (horizontal or "").strip().lower()
    return _EXCEL_H_ALIGN_TO_CSS.get(token, default)


def css_vertical_align_from_excel(vertical: str, *, default: str = "middle") -> str:
    """Map Excel vertical alignment to CSS; ``default`` when unset."""
    token = (vertical or "").strip().lower()
    return _EXCEL_V_ALIGN_TO_CSS.get(token, default)


def _htmlize_cell_value(value) -> str:
    """Escape cell text; convert Excel Alt+Enter newlines to <br>."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value)
    if text.strip().lower() in {"", "nan"}:
        return ""
    return (
        html.escape(text)
        .replace("\r\n", "<br>")
        .replace("\n", "<br>")
        .replace("\r", "<br>")
    )


def _htmlize_dataframe_cells(df: pd.DataFrame) -> pd.DataFrame:
    """Prepare DataFrame values for HTML display (escape + explicit line breaks)."""
    # pandas 2.1+ has DataFrame.map; fall back to applymap for older versions.
    mapper = getattr(df, "map", None) or df.applymap
    return mapper(_htmlize_cell_value)

CHA_REGION_ORDER = [
    "Dutchess",
    "Orange",
    "Putnam",
    "Rockland",
    "Sullivan",
    "Ulster",
    "Westchester",
    "Mid-Hudson",
    "NYS excl NYC",
    "NYS",
    "US",
]

CHA_REGION_ALIASES = {
    "Mid Hudson": "Mid-Hudson",
    "Mid-Hudson Region": "Mid-Hudson",
    "NYS excl. NYC": "NYS excl NYC",
    "NYS exc NYC": "NYS excl NYC",
    "NYS excluding NYC": "NYS excl NYC",
    "NYS excel NYC": "NYS excl NYC",
}


def _normalize_region_label(label):
    if label is None or pd.isna(label):
        return ""
    normalized = str(label).strip()
    return CHA_REGION_ALIASES.get(normalized, normalized)


def _reorder_columns_by_region(df):
    if isinstance(df.columns, pd.MultiIndex):
        return df

    columns = list(df.columns)
    normalized = {col: _normalize_region_label(col) for col in columns}
    region_cols = [col for col in columns if normalized[col] in CHA_REGION_ORDER]
    if len(region_cols) < 2:
        return df

    ordered_region_cols = sorted(
        region_cols,
        key=lambda col: CHA_REGION_ORDER.index(normalized[col]),
    )
    region_iter = iter(ordered_region_cols)
    new_columns = []
    for col in columns:
        if col in region_cols:
            new_columns.append(next(region_iter))
        else:
            new_columns.append(col)
    return df[new_columns]


def _reorder_rows_by_region(df):
    first_col = df.columns[0]
    row_labels = df[first_col]
    # If duplicate first-column headers exist, use the first actual column.
    if isinstance(row_labels, pd.DataFrame):
        row_labels = row_labels.iloc[:, 0]
    normalized_values = row_labels.map(_normalize_region_label)
    if normalized_values.empty:
        return df
    if (normalized_values == "").any():
        return df
    if not normalized_values.isin(CHA_REGION_ORDER).all():
        return df
    if normalized_values.nunique() < 2:
        return df

    order_index = row_labels.map(
        lambda value: CHA_REGION_ORDER.index(_normalize_region_label(value))
    )
    orig_columns = df.columns
    sorted_df = (
        df.set_axis(range(len(df.columns)), axis=1)
        .assign(_cha_region_order=order_index.values)
        .sort_values("_cha_region_order", kind="stable")
        .drop(columns=["_cha_region_order"])
        .set_axis(orig_columns, axis=1)
    )
    return sorted_df


def apply_cha_region_order(df):
    df = df.copy()
    df = _reorder_columns_by_region(df)
    df = _reorder_rows_by_region(df)
    return df


# ---------------------------------------------------------------------------
# Number formatting by data type
# ---------------------------------------------------------------------------

# Maps the plain-English "Data Type" label (from the Template.xlsx Test sheet
# and the Dropdowns sheet Y-labels column) to a pandas Styler format string.
# The format string is applied to every data column (all columns except the
# row-label column, which is always the first column).
DATA_TYPE_FORMATS: dict[str, str] = {
    # Percentages — one decimal place + % sign
    "percent":              "{:.1f}%",
    # Same label as used in the Dropdowns sheet
    "Percent":              "{:.1f}%",

    # Rates — one decimal place, no suffix
    "rate per 1,000":       "{:.1f}",
    "Rate per 1,000":       "{:.1f}",
    "rate per 10,000":      "{:.1f}",
    "Rate per 10,000":      "{:.1f}",
    "rate per 100,000":     "{:.1f}",
    "Rate per 100,000":     "{:.1f}",
    "case rate":            "{:.1f}",
    "Case Rate":            "{:.1f}",

    # Counts — whole number with thousands separator
    "number of cases":      "{:,.0f}",
    "Number of Cases":      "{:,.0f}",
    "count":                "{:,.0f}",
    "Count":                "{:,.0f}",

    # Ratio — whole number (e.g. 1:1,250 residents per provider)
    "ratio":                "{:,.0f}",
    "Ratio":                "{:,.0f}",

    # Index scores — two decimal places
    "index":                "{:.2f}",
    "Index":                "{:.2f}",

    # Currency — dollar sign + thousands separator, no decimals
    "currency":             "${:,.0f}",
    "Currency":             "${:,.0f}",
}


def get_format_string(data_type: str | None) -> str | None:
    """
    Return the pandas Styler format string for a plain-English data type label.

    Parameters
    ----------
    data_type : str or None
        Plain-English label such as ``"Percent"``, ``"Rate per 100,000"``, etc.
        Case-insensitive lookup is attempted if an exact match is not found.

    Returns
    -------
    str or None
        A Python format string (e.g. ``"{:.1f}%"``), or ``None`` if the data
        type is unrecognised (in which case numbers are left as-is).
    """
    if not data_type:
        return None
    # Exact match first
    if data_type in DATA_TYPE_FORMATS:
        return DATA_TYPE_FORMATS[data_type]
    # Case-insensitive fallback
    lower = data_type.strip().lower()
    for key, fmt in DATA_TYPE_FORMATS.items():
        if key.lower() == lower:
            return fmt
    return None


def _blank_header_label(value) -> bool:
    text = "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value).strip()
    return text == ""


def _header_label_text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    return "" if text == "" else text


def _multiindex_header_grid(columns: pd.MultiIndex):
    """Build header rows + colspan/rowspan merges for a MultiIndex."""
    from scripts.workbook_loader import CellMerge

    n_cols = len(columns)
    n_levels = columns.nlevels
    header_rows: list[list[object]] = [
        [_header_label_text(columns[c][level]) for c in range(n_cols)]
        for level in range(n_levels)
    ]
    merges: list[CellMerge] = []

    # First column: blank upper levels → rowspan leaf label across header rows.
    if n_cols > 0 and n_levels > 1:
        first_parts = [_header_label_text(columns[0][level]) for level in range(n_levels)]
        if all(_blank_header_label(part) for part in first_parts[:-1]) and not _blank_header_label(
            first_parts[-1]
        ):
            header_rows[0][0] = first_parts[-1]
            for level in range(1, n_levels):
                header_rows[level][0] = ""
            merges.append(CellMerge(row=0, col=0, rowspan=n_levels, colspan=1))

    # Merge consecutive identical labels on each header level (skip blanks).
    covered = {(m.row + dr, m.col + dc) for m in merges for dr in range(m.rowspan) for dc in range(m.colspan)}
    for level in range(n_levels):
        col = 0
        while col < n_cols:
            if (level, col) in covered:
                col += 1
                continue
            label = header_rows[level][col]
            if _blank_header_label(label):
                col += 1
                continue
            end = col + 1
            while (
                end < n_cols
                and (level, end) not in covered
                and header_rows[level][end] == label
            ):
                end += 1
            span = end - col
            if span > 1:
                merges.append(CellMerge(row=level, col=col, rowspan=1, colspan=span))
                for extra in range(col + 1, end):
                    header_rows[level][extra] = ""
                    covered.add((level, extra))
            col = end

    return header_rows, merges


def _dataframe_to_merged_grid(df: pd.DataFrame, cell_styles=None):
    """Convert a (possibly MultiIndex) DataFrame into a MergedTableGrid."""
    from scripts.workbook_loader import CellStyle, MergedTableGrid

    n_cols = len(df.columns)
    if isinstance(df.columns, pd.MultiIndex):
        header_rows, merges = _multiindex_header_grid(df.columns)
    else:
        header_rows = [[_header_label_text(col) for col in df.columns]]
        merges = []

    header_count = len(header_rows)
    body_rows = [
        tuple("" if (isinstance(v, float) and pd.isna(v)) or v is None else v for v in df.iloc[i].tolist())
        for i in range(len(df))
    ]
    cells = tuple(tuple(row) for row in header_rows) + tuple(body_rows)

    # MergedTableGrid.styles includes header rows; cell_styles are body-only.
    style_rows: list[tuple] = []
    for _ in range(header_count):
        style_rows.append(tuple(CellStyle(bold=True) for _ in range(n_cols)))
    if cell_styles:
        for row_idx in range(len(df)):
            if row_idx < len(cell_styles):
                row = list(cell_styles[row_idx])
                while len(row) < n_cols:
                    row.append(CellStyle())
                style_rows.append(tuple(row[:n_cols]))
            else:
                style_rows.append(tuple(CellStyle() for _ in range(n_cols)))
    elif n_cols > 0:
        # Match prior Styler behavior: bold the first column when no Excel styles.
        for _ in range(len(df)):
            style_rows.append(
                tuple(
                    CellStyle(bold=(col_idx == 0), horizontal="left" if col_idx == 0 else "")
                    for col_idx in range(n_cols)
                )
            )

    return MergedTableGrid(
        cells=cells,
        merges=tuple(merges),
        header_rows=header_count,
        format_rules_by_col=tuple("" for _ in range(n_cols)),
        styles=tuple(style_rows) if style_rows else None,
    )


def style_cha_table(df, has_multilevel_headers=False, data_type=None, row_label_col=None, cell_styles=None):
    """
    Apply consistent CHA table styling to a pandas DataFrame.
    
    Styling specifications:
    - Header: Light green (#EAF5DB) background, bold, centered
      (both levels stay green when multilevel headers are applied)
    - Row 1: White
    - Row 2: #EAF5DB (light green)
    - Row 3: White
    - Alternates: white, #EAF5DB, white, #EAF5DB...
    - First column: Bold (or workbook-driven), left-aligned when hierarchy styles apply
    - Other columns: Centered
    - Dark green separator line after "Westchester" row to separate county data from grouped areas

    Styles are emitted as inline CSS so Quarto HTML post-processing cannot strip
    them (pandas Styler ``<style>`` blocks are removed from published pages).
    
    Parameters
    ----------
    df : pandas.DataFrame
        The DataFrame to style
    has_multilevel_headers : bool, optional
        If True, applies special styling for MultiIndex column headers
        to create merged cell appearance (default: False)
    data_type : str, optional
        Plain-English data type label that controls how numbers are formatted
        in the table cells.  Accepted values match the Y-labels dropdown in
        Template.xlsx:

        * ``"Percent"``          → ``63.3%``
        * ``"Rate per 1,000"``   → ``12.4``
        * ``"Rate per 10,000"``  → ``45.2``
        * ``"Rate per 100,000"`` → ``234.5``
        * ``"Case Rate"``        → ``234.5``
        * ``"Number of Cases"``  → ``1,234``
        * ``"Ratio"``            → ``1,250``
        * ``"Index"``            → ``0.45``
        * ``"Currency"``         → ``$2,000``

        If ``None`` (default), numbers are left as-is.
    row_label_col : str, optional
        Name of the column that contains row labels (e.g. years).  This column
        is never formatted as a number.  Defaults to the first column.
    cell_styles : tuple of tuple of CellStyle, optional
        Per-cell Excel formatting aligned with ``df`` rows and columns.  When
        provided, bold and indent come from the workbook instead of structural
        first-column bolding.

    Returns
    -------
    IPython.display.HTML
        Inline-styled HTML table ready for display in Quarto

    Example
    -------
    >>> import pandas as pd
    >>> from scripts.cha_table_styling import style_cha_table
    >>>
    >>> data = {'Region': ['A', 'B'], 'Value': [100, 200]}
    >>> df = pd.DataFrame(data)
    >>> styled = style_cha_table(df, data_type="Percent")
    >>> styled  # Display in Quarto
    """
    # Lazy import avoids a circular dependency with cha_merged_table.
    from scripts.cha_merged_table import render_merged_table

    df = apply_cha_region_order(df)

    # ── Number formatting ────────────────────────────────────────────────────
    # Determine the row-label column (first column) – never formatted as a number
    _row_label_col = row_label_col if row_label_col is not None else (
        df.columns[0] if len(df.columns) > 0 else None
    )
    _fmt_str = get_format_string(data_type)
    _format_dict: dict = {}
    if _fmt_str:
        for col in df.columns:
            if col != _row_label_col:
                _format_dict[col] = _fmt_str
    if _format_dict:
        # Format numerics before render so format strings see raw values.
        # HTML escaping happens in render_merged_table.
        df = df.copy()
        for col, fmt in _format_dict.items():
            if col not in df.columns:
                continue
            col_data = df[col]
            if isinstance(col_data, pd.DataFrame):
                continue
            df[col] = col_data.map(
                lambda v, f=fmt: (
                    ""
                    if v is None or (isinstance(v, float) and pd.isna(v))
                    else (f.format(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else v)
                )
            )
    # ─────────────────────────────────────────────────────────────────────────

    # has_multilevel_headers is reserved for callers that already built a
    # MultiIndex; detection is based on the column type either way.
    _ = has_multilevel_headers
    grid = _dataframe_to_merged_grid(df, cell_styles=cell_styles)
    return render_merged_table(grid)


def format_source_citation(table_id, url, data_year=2023, estimate_type="5-Year Estimates", citation_month="April", citation_year=2025, custom_text=None):
    """
    Create a standardized source citation with hyperlink.
    
    This is the STANDARD format for all CHA table sources.
    Flexible to handle different table IDs, years, and estimate types.
    
    Parameters
    ----------
    table_id : str
        The Census Bureau table ID (e.g., "S0101", "B03002", "S1601")
    url : str
        The full URL to the Census Bureau data table
    data_year : int, optional
        The data year (default: 2023)
    estimate_type : str, optional
        The estimate type - "5-Year Estimates" or "5-year estimates" (default: "5-Year Estimates")
    citation_month : str, optional
        The citation month (default: "April")
    citation_year : int, optional
        The citation year (default: 2025)
    custom_text : str, optional
        Custom citation text to use instead of standard format. If provided, 
        other parameters are ignored except url which should be embedded in custom_text.
        
    Returns
    -------
    str
        Formatted source citation in standard CHA format
        
    Examples
    --------
    >>> # Standard Census Bureau table
    >>> citation = format_source_citation(
    ...     "S0101",
    ...     "https://data.census.gov/table/ACSST5Y2023.S0101?..."
    ... )
    
    >>> # Different year
    >>> citation = format_source_citation(
    ...     "S1810",
    ...     "https://data.census.gov/table/...",
    ...     data_year=2020
    ... )
    
    >>> # Custom citation (for non-Census sources)
    >>> citation = format_source_citation(
    ...     "",
    ...     "",
    ...     custom_text="New York State Department of Health, [Data Source](https://...), 2025"
    ... )
    """
    if custom_text:
        return custom_text
    
    return f'US Census Bureau; American Community Survey, {data_year} American Community Survey {estimate_type}, [Table {table_id}]({url}), {citation_month} {citation_year}'


def create_source_callout(table_id=None, url=None, data_year=2023, estimate_type="5-Year Estimates", citation_month="April", citation_year=2025, custom_text=None):
    """
    Create a complete collapsible source callout box in Quarto format.
    
    This is the STANDARD format for all CHA table sources.
    Flexible to handle different sources from the CHA document.
    
    Parameters
    ----------
    table_id : str, optional
        The Census Bureau table ID (e.g., "S0101", "B03002", "S1601")
        Required if custom_text is not provided
    url : str, optional
        The full URL to the data source
        Required if custom_text is not provided
    data_year : int, optional
        The data year (default: 2023)
    estimate_type : str, optional
        The estimate type - "5-Year Estimates" or "5-year estimates" (default: "5-Year Estimates")
    citation_month : str, optional
        The citation month (default: "April")
    citation_year : int, optional
        The citation year (default: 2025)
    custom_text : str, optional
        Custom citation text for non-Census sources or special cases.
        Should include hyperlinks in markdown format: [text](url)
        If provided, other parameters are ignored.
        
    Returns
    -------
    str
        Complete Quarto callout block for source citation
        
    Examples
    --------
    >>> # Standard Census Bureau table
    >>> callout = create_source_callout(
    ...     "S0101",
    ...     "https://data.census.gov/table/ACSST5Y2023.S0101?..."
    ... )
    
    >>> # Different year (2020 data)
    >>> callout = create_source_callout(
    ...     "S1810",
    ...     "https://data.census.gov/table/...",
    ...     data_year=2020
    ... )
    
    >>> # Custom source (e.g., NYS Department of Health)
    >>> callout = create_source_callout(
    ...     custom_text="New York State Department of Health, [Vital Statistics](https://...), 2025"
    ... )
    """
    citation = format_source_citation(
        table_id or "", 
        url or "", 
        data_year, 
        estimate_type, 
        citation_month, 
        citation_year,
        custom_text
    )
    return f'''::: {{.callout-note collapse="true"}}
## Source

{citation}
:::'''


# Standard source citation format
STANDARD_SOURCE_FORMAT = """::: {.callout-note collapse="true"}
## Source

US Census Bureau; American Community Survey, {year} American Community Survey 5-Year Estimates, [Table {table_id}]({url}), {month} {citation_year}
:::"""

# Template for Quarto table code block
QUARTO_TABLE_TEMPLATE = '''```{{python}}
#| echo: false
#| warning: false
#| message: false
#| label: tbl-{table_label}
#| tbl-cap: "{table_caption}"
import pandas as pd
import numpy as np
from scripts.cha_table_styling import style_cha_table

# Create the data
data = {{
    # Your data dictionary here
}}

df = pd.DataFrame(data)

# Format the data (adjust as needed)
# df["Column Name"] = df["Column Name"].apply(lambda x: f"{{x:,}}")
# df["Percent Column"] = df["Percent Column"].apply(lambda x: f"{{x:.1f}}" if pd.notna(x) else "N/A")

# Apply standard CHA styling
styled_table = style_cha_table(df)
styled_table
```

{source_callout}
'''
