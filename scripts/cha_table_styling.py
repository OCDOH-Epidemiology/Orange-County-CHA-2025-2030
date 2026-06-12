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

import pandas as pd

CHA_FONT_FAMILY = '"Tw Cen MT", "Tw Cen MT Std", "Century Gothic", "Trebuchet MS", "Segoe UI", sans-serif'
# Approximate one Excel indent level (~3 character widths at 14px).
EXCEL_INDENT_PX_PER_LEVEL = 27


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


def style_cha_table(df, has_multilevel_headers=False, data_type=None, row_label_col=None, cell_styles=None):
    """
    Apply consistent CHA table styling to a pandas DataFrame.
    
    Styling specifications:
    - Header: White background, bold, centered
    - Row 1: #EAF5DB (light green)
    - Row 2: White
    - Row 3: #EAF5DB (light green)
    - Alternates: white, #EAF5DB, white, #EAF5DB...
    - First column: Bold, center-aligned
    - Other columns: Center-aligned
    - Dark green separator line after "Westchester" row to separate county data from grouped areas
    
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
    pandas.io.formats.style.Styler
        A styled DataFrame ready for display in Quarto

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
    df = apply_cha_region_order(df)

    # ── Number formatting ────────────────────────────────────────────────────
    # Determine the row-label column (first column) – never formatted as a number
    _row_label_col = row_label_col if row_label_col is not None else (
        df.columns[0] if len(df.columns) > 0 else None
    )
    # Build a per-column format dict for pandas Styler.format()
    _fmt_str = get_format_string(data_type)
    _format_dict: dict = {}
    if _fmt_str:
        for col in df.columns:
            if col != _row_label_col:
                _format_dict[col] = _fmt_str
    # ─────────────────────────────────────────────────────────────────────────

    styles = [
        # Header styling - white background
        {'selector': 'th', 'props': [
            ('font-weight', 'bold'), 
            ('text-align', 'center'), 
            ('background-color', '#FFFFFF'), 
            ('font-family', CHA_FONT_FAMILY),
            ('padding', '10px'),
            ('border', '1px solid #ddd')
        ]},
        # First column - bold, center-aligned (skipped when per-cell styles are provided)
        *([] if cell_styles else [{'selector': 'td:first-child', 'props': [
            ('font-weight', 'bold'),
            ('text-align', 'center'),
            ('font-family', CHA_FONT_FAMILY),
            ('padding', '10px'),
            ('border', '1px solid #ddd')
        ]}]),
        # When Excel cell styles drive formatting, avoid text-align/padding on
        # td:first-child — those shorthand rules override per-cell indent styles.
        *([{'selector': 'td:first-child', 'props': [
            ('font-family', CHA_FONT_FAMILY),
            ('padding', '10px'),
            ('border', '1px solid #ddd'),
        ]}] if cell_styles else []),
        # Other columns - center-aligned
        {'selector': 'td:not(:first-child)', 'props': [
            ('text-align', 'center'), 
            ('font-family', CHA_FONT_FAMILY),
            ('padding', '10px'),
            ('border', '1px solid #ddd')
        ]},
        # Ensure row striping starts with row 1 (green), then white, then alternate.
        # Target td cells directly with !important to win against framework defaults.
        {'selector': 'tbody tr:nth-child(odd) td', 'props': [
            ('background-color', '#EAF5DB !important')
        ]},
        {'selector': 'tbody tr:nth-child(even) td', 'props': [
            ('background-color', '#FFFFFF !important')
        ]},
        # Table container
        {'selector': 'table', 'props': [
            ('border-collapse', 'collapse'), 
            ('width', '100%'), 
            ('margin', '20px 0'),
            ('font-family', CHA_FONT_FAMILY),
            ('font-size', '14px')
        ]}
    ]
    
    # Add MultiIndex header styling if needed
    if has_multilevel_headers and isinstance(df.columns, pd.MultiIndex):
        # Style for top-level headers (merged appearance)
        # Target all top-level header cells
        styles.append({
            'selector': 'thead tr:first-child th', 
            'props': [
                ('border-bottom', '2px solid #333'),
                ('font-weight', 'bold'),
                ('background-color', '#FFFFFF'),
                ('text-align', 'center'),
                ('vertical-align', 'middle'),
                ('padding', '10px')
            ]
        })
        # Style for second-level headers
        styles.append({
            'selector': 'thead tr:last-child th', 
            'props': [
                ('font-weight', 'normal'),
                ('font-size', '0.9em'),
                ('text-align', 'center'),
                ('background-color', '#FFFFFF'),
                ('padding', '10px')
            ]
        })
        # For MultiIndex, we need to handle the first column header separately
        styles.append({
            'selector': 'thead tr:first-child th:first-child', 
            'props': [
                ('text-align', 'left'),
            ]
        })
    
    # Find the row index where "Westchester" appears in the first column
    # and add a dark green border-bottom separator
    first_col = df.columns[0]
    
    # Create the styled table
    styled = df.style.set_table_styles(styles).hide(axis="index")
    # Apply number formatting if a data_type was provided
    if _format_dict:
        styled = styled.format(_format_dict, na_rep="")
    
    # Add dark green border-bottom to the Westchester row if found
    # Dark green color: using a dark green shade
    dark_green = '#2d5016'  # Dark green color
    
    # Create a function that enforces zebra striping and (optionally)
    # applies a dark green border below the Westchester row.
    row_position_lookup = {idx: pos for pos, idx in enumerate(df.index)}

    def style_row(row):
        # Row 1 should be green, then white, then alternate.
        row_pos = row_position_lookup.get(row.name, 0)
        base_bg = '#EAF5DB' if row_pos % 2 == 0 else '#FFFFFF'
        base_css = f'background-color: {base_bg} !important'

        # Check if this row contains "Westchester" in the first column.
        first_val = row[first_col] if first_col in row.index else None
        if isinstance(first_val, (pd.Series, pd.DataFrame)):
            if isinstance(first_val, pd.DataFrame):
                first_val = first_val.iloc[0, 0] if not first_val.empty else None
            else:
                first_val = first_val.iloc[0] if not first_val.empty else None

        westchester_border = (
            first_val is not None
            and pd.notna(first_val)
            and str(first_val).strip().lower() == 'westchester'
        )

        cell_css_list: list[str] = []
        for col_idx, _col in enumerate(row.index):
            parts = [base_css]
            if westchester_border:
                parts.append(f'border-bottom: 3px solid {dark_green}')
            if cell_styles and row_pos < len(cell_styles):
                row_styles = cell_styles[row_pos]
                if col_idx < len(row_styles):
                    style = row_styles[col_idx]
                    parts.append('font-weight: bold' if style.bold else 'font-weight: normal')
                    if col_idx == 0:
                        parts.append('text-align: left !important')
                        if style.indent > 0:
                            indent_px = style.indent * EXCEL_INDENT_PX_PER_LEVEL
                            parts.append(f'text-indent: {indent_px}px !important')
            cell_css_list.append('; '.join(parts))
        return cell_css_list
    
    # Apply row-level styling only when index/columns are unique;
    # pandas Styler does not support .apply with non-unique labels.
    if df.index.is_unique and df.columns.is_unique:
        styled = styled.apply(style_row, axis=1)
    
    return styled


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
