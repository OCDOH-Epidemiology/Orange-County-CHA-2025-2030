"""
Workbook loader for metadata-driven CHA rendering.

Supports both:
- Normalized workbook metadata sheets (`_registry`, `_figure_specs`, `_table_specs`)
- Flat per-indicator sheets used by the Mid-Hudson CHA workbook template
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any
import json
import re

import pandas as pd


VALID_OBJECT_TYPES = {"table", "figure"}
VALID_FIGURE_TYPES = {"line", "clustered_bar", "stacked_bar", "simple_bar", "horizontal_bar"}
_VALID_FORMAT_CODES: frozenset[str] = frozenset(
    {"integer", "number", "percent1", "percent2", "currency", "currency2", "ratio", "date"}
)

_FIGURE_TYPE_ALIASES: dict[str, str] = {
    "1": "line",
    "line": "line",
    "2": "clustered_bar",
    "clustered_bar": "clustered_bar",
    "clustered bar": "clustered_bar",
    "cluster bar": "clustered_bar",
    "3": "stacked_bar",
    "stacked_bar": "stacked_bar",
    "stacked bar": "stacked_bar",
    "stack bar": "stacked_bar",
    "horizontal stacked bar": "stacked_bar",
    "4": "simple_bar",
    "simple_bar": "simple_bar",
    "simple bar": "simple_bar",
    "bar": "simple_bar",
    "5": "horizontal_bar",
    "horizontal_bar": "horizontal_bar",
    "horizontal bar": "horizontal_bar",
}


@dataclass(frozen=True)
class RegistryRecord:
    object_id: str
    object_type: str
    label: str
    caption: str
    data_sheet: str
    enabled: bool
    section_tag: str
    order_index: int


@dataclass(frozen=True)
class FigureSpec:
    object_id: str
    figure_type: str
    x_col: str
    y_cols: list[str]
    x_axis_title: str
    y_axis_title: str
    start_at_zero: bool
    hover_suffix: str
    pivot_for_chart: bool = False
    show_data_labels: bool | None = None


@dataclass(frozen=True)
class TableSpec:
    object_id: str
    has_multilevel_headers: bool
    format_rules: dict[str, str]
    row_label_col: str


@dataclass(frozen=True)
class SourceSpec:
    object_id: str
    table_id: str
    url: str
    data_year: int
    estimate_type: str
    citation_month: str
    citation_year: int
    custom_text: str


@dataclass(frozen=True)
class WorkbookModel:
    workbook_path: Path
    registry: dict[str, RegistryRecord]
    figure_specs: dict[str, FigureSpec]
    table_specs: dict[str, TableSpec]
    source_specs: dict[str, SourceSpec]
    data_frames: dict[str, pd.DataFrame]


def _as_bool(value: Any, default: bool = False) -> bool:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return default
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text == "":
        return default
    return text in {"1", "true", "yes", "y", "enabled"}


def _as_text(value: Any, default: str = "") -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return default
    return str(value).strip()


def _as_int(value: Any, default: int = 0) -> int:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return default
    text = str(value).strip()
    if text == "":
        return default
    try:
        return int(float(text.replace(",", "")))
    except ValueError:
        # Accept year ranges like "2021-2023" / "2021–2023" by taking the first year.
        match = re.search(r"\d{4}", text)
        if match:
            return int(match.group(0))
        return default


def _parse_string_list(value: Any) -> list[str]:
    text = _as_text(value)
    if not text:
        return []
    return [item.strip() for item in text.split(",") if item.strip()]


def _normalize_figure_type(value: Any, default: str | None = "line") -> str:
    text = _as_text(value).strip().lower()
    if not text:
        return default or ""
    normalized = _FIGURE_TYPE_ALIASES.get(text)
    if normalized:
        return normalized
    text = text.replace("-", "_")
    if text in VALID_FIGURE_TYPES:
        return text
    return default or ""


def _group_by_to_pivot_for_chart(group_by_value: Any) -> bool | None:
    """
    Map optional grouping intent metadata to pivot_for_chart semantics.

    - group_by = x_col  -> group bars by x_col values (pivot required)
    - group_by = series -> group bars by existing series columns (no pivot)
    """
    text = _as_text(group_by_value).strip().lower()
    if not text:
        return None
    normalized = text.replace(" ", "_").replace("-", "_")
    if normalized == "x_col":
        return True
    if normalized == "series":
        return False
    return None


def _flat_slug(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")


def _normalize_config_key(text: Any) -> str:
    raw = _as_text(text).lower()
    return re.sub(r"[^a-z0-9]+", " ", raw).strip()


def _config_value(config: dict[str, Any], key: str, default: Any = None) -> Any:
    if key in config:
        return config[key]
    normalized_key = _normalize_config_key(key)
    for existing_key, value in config.items():
        if _normalize_config_key(existing_key) == normalized_key:
            return value
    return default


def _read_excel_raw(path: Path) -> dict[str, pd.DataFrame]:
    if not path.exists():
        raise FileNotFoundError(f"Workbook not found: {path}")
    return pd.read_excel(path, sheet_name=None, header=None, keep_default_na=False, na_values=[""])


def _read_excel_with_headers(path: Path) -> dict[str, pd.DataFrame]:
    if not path.exists():
        raise FileNotFoundError(f"Workbook not found: {path}")
    return pd.read_excel(path, sheet_name=None, keep_default_na=False, na_values=[""])


def _is_flat_indicator_sheet(df: pd.DataFrame) -> bool:
    return bool(df is not None and not df.empty and df.shape[1] >= 2 and _as_text(df.iloc[0, 0]).lower() == "name")


def _parse_flat_indicator_sheet(
    sheet_name: str, df: pd.DataFrame
) -> tuple[dict[str, Any], pd.DataFrame, dict[str, str], bool]:
    config: dict[str, Any] = {}
    for i in range(len(df)):
        key = _as_text(df.iloc[i, 0] if df.shape[1] > 0 else None)
        if key:
            config[key] = df.iloc[i, 1] if df.shape[1] > 1 else None

    data_header_idx: int | None = None
    for i in range(len(df)):
        if df.shape[1] > 4 and _as_text(df.iloc[i, 4]) == "Enter Data":
            data_header_idx = i
            break

    # In flat template sheets, all columns to the right of "Enter Data"
    # are indicator data columns (first data column starts at index 5 / col F).
    data_col_start_idx = 5
    if data_header_idx is None or df.shape[1] <= data_col_start_idx:
        return config, pd.DataFrame(), {}, False

    # Determine right-most used data column from a small header/data window.
    # This avoids trailing blank worksheet columns becoming ghost headers.
    header_scan_rows = range(data_header_idx, min(len(df), data_header_idx + 6))
    last_used_col = data_col_start_idx - 1
    for j in range(data_col_start_idx, df.shape[1]):
        if any(_as_text(df.iloc[i, j]) != "" for i in header_scan_rows):
            last_used_col = j

    if last_used_col < data_col_start_idx:
        return config, pd.DataFrame(), {}, False

    data_col_end_idx = last_used_col + 1
    header_row_idx = data_header_idx
    format_row_idx = data_header_idx - 1
    raw_headers = [df.iloc[header_row_idx, j] for j in range(data_col_start_idx, data_col_end_idx)]
    headers = [_as_text(h) for h in raw_headers]

    # Some sheets place format codes (integer/percent1/etc.) on the same row as
    # "Enter Data" and put real headers on the next row.
    nonempty_headers = [h for h in headers if h]
    looks_like_format_header_row = bool(nonempty_headers) and all(
        _as_text(h).lower() in _VALID_FORMAT_CODES for h in nonempty_headers
    )
    if looks_like_format_header_row and header_row_idx + 1 < len(df):
        next_headers = [
            _as_text(df.iloc[header_row_idx + 1, j])
            for j in range(data_col_start_idx, data_col_end_idx)
        ]
        next_nonempty = [h for h in next_headers if h]
        next_is_real_headers = bool(next_nonempty) and not all(
            h.lower() in _VALID_FORMAT_CODES for h in next_nonempty
        )
        if next_is_real_headers:
            format_row_idx = header_row_idx
            header_row_idx = header_row_idx + 1
            headers = next_headers

    # Detect two-row merged header pattern (multilevel table).
    data_start_idx = header_row_idx + 1
    auto_multilevel = False
    sub_header_row_idx = data_header_idx + 1
    explicit_multilevel = (
        _as_bool(_config_value(config, "Multilevel Headers", False))
        or _as_bool(_config_value(config, "Multiheader Levels", False))
    )
    if sub_header_row_idx < len(df):
        row_label_blank = _as_text(df.iloc[sub_header_row_idx, data_col_start_idx]) == ""
        has_data_headers = any(
            _as_text(df.iloc[sub_header_row_idx, j])
            for j in range(data_col_start_idx + 1, data_col_end_idx)
        )

        def _looks_numeric_cell(value: Any) -> bool:
            if value is None or (isinstance(value, float) and pd.isna(value)):
                return False
            if isinstance(value, (int, float)):
                return True
            text = _as_text(value)
            if not text:
                return False
            text = text.replace(",", "").replace("$", "").rstrip("%").strip()
            text = re.sub(r"[*†‡§#]+$", "", text)
            try:
                float(text)
                return True
            except ValueError:
                return False

        # Primary detection: traditional template shape (blank first subheader cell).
        use_subheader_row = row_label_blank and has_data_headers

        # Secondary detection (enabled for explicit multilevel metadata):
        # some sheets repeat "County" in the first subheader cell.
        # Treat that row as header only when the following row looks numeric.
        if not use_subheader_row and explicit_multilevel and has_data_headers:
            candidate_cells = [
                df.iloc[sub_header_row_idx, j]
                for j in range(data_col_start_idx + 1, data_col_end_idx)
            ]
            candidate_nonempty = [cell for cell in candidate_cells if _as_text(cell)]
            candidate_numeric_count = sum(1 for cell in candidate_cells if _looks_numeric_cell(cell))

            next_row_idx = sub_header_row_idx + 1
            next_row_numeric_count = 0
            if next_row_idx < len(df):
                next_row_cells = [
                    df.iloc[next_row_idx, j]
                    for j in range(data_col_start_idx + 1, data_col_end_idx)
                ]
                next_row_numeric_count = sum(1 for cell in next_row_cells if _looks_numeric_cell(cell))

            has_label_like_subheaders = len(candidate_nonempty) >= 2 and candidate_numeric_count <= 1
            has_numeric_data_after = next_row_numeric_count >= 2
            use_subheader_row = has_label_like_subheaders and has_numeric_data_after

        if use_subheader_row:
            sub_headers = [
                _as_text(df.iloc[sub_header_row_idx, j])
                for j in range(data_col_start_idx, data_col_end_idx)
            ]
            merged: list[str] = []
            last_top = ""
            for col_idx, (h_top, h_sub) in enumerate(zip(headers, sub_headers)):
                effective_top = h_top if h_top else last_top
                if h_top:
                    last_top = h_top
                if col_idx == 0:
                    merged.append(h_sub or h_top)
                elif effective_top and h_sub:
                    merged.append(f"{effective_top}|{h_sub}")
                elif h_sub:
                    merged.append(h_sub)
                else:
                    merged.append(effective_top)
            headers = merged
            data_start_idx = sub_header_row_idx + 1
            # Only treat as true multilevel when the top header is semantic
            # (e.g., "Three-Year Average"), not a format token (e.g., "percent1").
            auto_multilevel = any(
                "|" in col and _as_text(col.split("|", 1)[0]).lower() not in _VALID_FORMAT_CODES
                for col in headers
            )

    # Read per-column format rules from row above Enter Data.
    format_rules: dict[str, str] = {}
    if format_row_idx >= 0:
        fmt_row_idx = format_row_idx
        for j in range(data_col_start_idx, data_col_end_idx):
            header_pos = j - data_col_start_idx
            if header_pos < len(headers):
                fmt_val = _as_text(df.iloc[fmt_row_idx, j]).lower()
                if fmt_val in _VALID_FORMAT_CODES:
                    format_rules[headers[header_pos]] = fmt_val

    # Also infer format rules from tokenized headers: "percent1|Dutchess"
    for header in headers:
        text = _as_text(header)
        if "|" in text:
            left, _right = text.split("|", 1)
            fmt_name = _as_text(left).lower()
            if fmt_name in _VALID_FORMAT_CODES:
                format_rules[text] = fmt_name

    def _clean_header_label(label: Any) -> str:
        text = _as_text(label)
        if "|" in text:
            left, right = text.split("|", 1)
            if _as_text(left).lower() in _VALID_FORMAT_CODES:
                return _as_text(right)
        return text

    data_rows: list[list[Any]] = []
    for i in range(data_start_idx, len(df)):
        row_label = df.iloc[i, data_col_start_idx]
        if row_label is None or (isinstance(row_label, float) and pd.isna(row_label)) or _as_text(row_label) == "":
            break
        row_values = [
            df.iloc[i, j] if j < df.shape[1] else None
            for j in range(data_col_start_idx, data_col_start_idx + len(headers))
        ]
        data_rows.append(row_values)

    data_df = pd.DataFrame(data_rows, columns=headers) if data_rows else pd.DataFrame(columns=headers)
    cleaned_headers = [_clean_header_label(col) for col in data_df.columns]
    data_df = data_df.copy()
    data_df.columns = cleaned_headers

    remapped_rules: dict[str, str] = {}
    for col_name, fmt in format_rules.items():
        remapped_rules[_clean_header_label(col_name)] = fmt

    return config, data_df, remapped_rules, auto_multilevel


def _load_flat_workbook(source_path: Path) -> WorkbookModel:
    raw_sheets = _read_excel_raw(source_path)
    skip_sheets = {"Master", "Dropdowns", "Template", "_Template", "_Template (2)"}

    registry: dict[str, RegistryRecord] = {}
    figure_specs: dict[str, FigureSpec] = {}
    table_specs: dict[str, TableSpec] = {}
    source_specs: dict[str, SourceSpec] = {}
    data_frames: dict[str, pd.DataFrame] = {}
    order_counter = 0

    for sheet_name, sheet_df in raw_sheets.items():
        if sheet_name in skip_sheets:
            continue
        if sheet_name.startswith("_"):
            continue
        if not _is_flat_indicator_sheet(sheet_df):
            continue

        config, data_df, format_rules, auto_multilevel = _parse_flat_indicator_sheet(sheet_name, sheet_df)

        # When the user sets "X Column" but it doesn't match any data column,
        # reconcile the mismatch so the spec, y_cols, and renderer all agree.
        configured_x = _as_text(_config_value(config, "X Column", ""))
        # A comma in the configured value means it is a list of category
        # values, not a column name — skip the rename in that case.
        is_column_name = configured_x and "," not in configured_x
        if is_column_name and not data_df.empty and configured_x not in data_df.columns:
            first_col_name = data_df.columns[0]
            if first_col_name == "":
                # Blank first column → rename to match the configured X Column.
                new_cols = list(data_df.columns)
                new_cols[0] = configured_x
                data_df.columns = new_cols
            elif first_col_name.lower() in {
                "year", "period", "race/ethnicity", "group", "category",
                "county", "region", "integer",
            }:
                # Common synonym mismatch (e.g., X Column says "Period" but
                # header says "Year"). Rename to the configured value so the
                # spec stays consistent.
                new_cols = list(data_df.columns)
                new_cols[0] = configured_x
                data_df.columns = new_cols

        raw_type = _as_text(_config_value(config, "Table/Figure/Both", "both")).lower()
        sheet_slug = _flat_slug(sheet_name)

        object_id_override = _as_text(_config_value(config, "Object ID", ""))
        figure_id_override = _as_text(_config_value(config, "Figure ID", "")) or object_id_override

        def _make_base_id(override: str) -> str:
            slug = _flat_slug(override) if override else sheet_slug
            for prefix in ("tbl-", "fig-"):
                if slug.startswith(prefix):
                    slug = slug[len(prefix):]
            return slug

        tbl_base_id = _make_base_id(object_id_override)
        fig_base_id = _make_base_id(figure_id_override)

        caption = _as_text(_config_value(config, "Name", sheet_name))
        section_tag = tbl_base_id
        data_sheet_name = f"data_{tbl_base_id}"
        data_frames[data_sheet_name] = data_df.copy()

        to_create: list[str] = []
        if raw_type in {"figure", "both"}:
            to_create.append("figure")
        if raw_type in {"table", "both"}:
            to_create.append("table")

        # Some flat-workbook sheets keep "Figure" in Table/Figure/Both while
        # still providing explicit tbl-/fig- IDs. Respect explicit IDs so both
        # include objects are generated when requested in metadata.
        object_id_lower = object_id_override.lower()
        figure_id_lower = figure_id_override.lower()
        if object_id_lower.startswith("tbl-") and "table" not in to_create:
            to_create.append("table")
        if figure_id_lower.startswith("fig-") and "figure" not in to_create:
            to_create.append("figure")

        for obj_type in to_create:
            base = fig_base_id if obj_type == "figure" else tbl_base_id
            prefix = "fig" if obj_type == "figure" else "tbl"
            object_id = f"{prefix}-{base}"
            registry[object_id] = RegistryRecord(
                object_id=object_id,
                object_type=obj_type,
                label=object_id,
                caption=caption,
                data_sheet=data_sheet_name,
                enabled=True,
                section_tag=section_tag,
                order_index=order_counter,
            )

            if obj_type == "figure":
                figure_type = _normalize_figure_type(_config_value(config, "Figure Type", "line"), default="line")
                x_col = _as_text(_config_value(config, "X Column", "")) or (data_df.columns[0] if not data_df.empty else "")
                y_cols_cfg = _parse_string_list(_config_value(config, "Y Column", ""))
                if y_cols_cfg:
                    y_cols = [col for col in y_cols_cfg if col in data_df.columns and col != x_col]
                else:
                    y_cols = [col for col in data_df.columns if col != x_col]
                group_by_override = _group_by_to_pivot_for_chart(_config_value(config, "Group By", ""))
                pivot_for_chart = (
                    group_by_override
                    if group_by_override is not None
                    else _as_bool(_config_value(config, "Pivot For Chart", False))
                )
                show_data_labels_raw = _config_value(config, "Show Data Labels", "")
                figure_specs[object_id] = FigureSpec(
                    object_id=object_id,
                    figure_type=figure_type,
                    x_col=x_col,
                    y_cols=y_cols,
                    x_axis_title=_as_text(_config_value(config, "X Axis Title", x_col)),
                    y_axis_title=_as_text(_config_value(config, "Y Axis Title", "")),
                    start_at_zero=_as_bool(_config_value(config, "Start at Zero", False)),
                    hover_suffix=_as_text(_config_value(config, "Hover Suffix", "%")),
                    pivot_for_chart=pivot_for_chart,
                    show_data_labels=(
                        _as_bool(show_data_labels_raw, default=False)
                        if _as_text(show_data_labels_raw) != ""
                        else None
                    ),
                )
            else:
                has_multilevel = auto_multilevel or _as_bool(_config_value(config, "Multilevel Headers", False))
                # fallback by generic Data Type if no explicit format rules detected
                if not format_rules and not data_df.empty:
                    data_type = _as_text(_config_value(config, "Data Type", "")).lower()
                    data_cols = list(data_df.columns[1:])
                    if data_type in {"percent", "percentage"}:
                        format_rules = {col: "percent1" for col in data_cols}
                    elif data_type == "number":
                        format_rules = {col: "number" for col in data_cols}
                    elif data_type == "currency":
                        format_rules = {col: "currency" for col in data_cols}
                table_specs[object_id] = TableSpec(
                    object_id=object_id,
                    has_multilevel_headers=has_multilevel,
                    format_rules=format_rules,
                    row_label_col=(data_df.columns[0] if not data_df.empty else ""),
                )

            source_specs[object_id] = SourceSpec(
                object_id=object_id,
                table_id=_as_text(_config_value(config, "Table ID", "")),
                url=_as_text(_config_value(config, "URL", "")),
                data_year=_as_int(_config_value(config, "Data Year", 2023), default=2023),
                estimate_type=_as_text(_config_value(config, "Estimate Type", "5-Year Estimates")),
                citation_month=_as_text(_config_value(config, "Citation Month", "April")),
                citation_year=_as_int(_config_value(config, "Citation Year", 2025), default=2025),
                custom_text=_as_text(_config_value(config, "Custom Text", "")),
            )

        order_counter += 1

    return WorkbookModel(
        workbook_path=source_path,
        registry=registry,
        figure_specs=figure_specs,
        table_specs=table_specs,
        source_specs=source_specs,
        data_frames=data_frames,
    )


def _load_normalized_workbook(source_path: Path) -> WorkbookModel:
    sheets = _read_excel_with_headers(source_path)
    required = {"_registry", "_figure_specs", "_table_specs"}
    if not required.issubset(sheets.keys()):
        missing = sorted(required - set(sheets.keys()))
        raise ValueError("Missing required metadata sheets: " + ", ".join(missing))

    registry_df = sheets["_registry"].copy()
    figure_df = sheets["_figure_specs"].copy()
    table_df = sheets["_table_specs"].copy()
    source_df = sheets.get("_source_specs", pd.DataFrame())

    registry: dict[str, RegistryRecord] = {}
    for row in registry_df.to_dict("records"):
        if not _as_bool(row.get("enabled", True), default=True):
            continue
        object_id = _as_text(row.get("object_id"))
        if not object_id:
            continue
        object_type = _as_text(row.get("object_type")).lower()
        if object_type not in VALID_OBJECT_TYPES:
            continue
        registry[object_id] = RegistryRecord(
            object_id=object_id,
            object_type=object_type,
            label=_as_text(row.get("label", object_id)),
            caption=_as_text(row.get("caption", "")),
            data_sheet=_as_text(row.get("data_sheet", "")),
            enabled=True,
            section_tag=_as_text(row.get("section_tag", "")),
            order_index=_as_int(row.get("order_index", 0), default=0),
        )

    figure_specs: dict[str, FigureSpec] = {}
    for row in figure_df.to_dict("records"):
        object_id = _as_text(row.get("object_id", ""))
        if object_id not in registry:
            continue
        group_by_override = _group_by_to_pivot_for_chart(row.get("group_by", ""))
        pivot_for_chart = (
            group_by_override
            if group_by_override is not None
            else _as_bool(row.get("pivot_for_chart", False))
        )
        figure_specs[object_id] = FigureSpec(
            object_id=object_id,
            figure_type=_normalize_figure_type(row.get("figure_type", "line"), default="line"),
            x_col=_as_text(row.get("x_col", "")),
            y_cols=_parse_string_list(row.get("y_cols", "")),
            x_axis_title=_as_text(row.get("x_axis_title", "")),
            y_axis_title=_as_text(row.get("y_axis_title", "")),
            start_at_zero=_as_bool(row.get("start_at_zero", False)),
            hover_suffix=_as_text(row.get("hover_suffix", "")),
            pivot_for_chart=pivot_for_chart,
            show_data_labels=(
                _as_bool(row.get("show_data_labels", False))
                if str(row.get("show_data_labels", "")).strip() != ""
                else None
            ),
        )

    table_specs: dict[str, TableSpec] = {}
    for row in table_df.to_dict("records"):
        object_id = _as_text(row.get("object_id", ""))
        if object_id not in registry:
            continue
        rules_raw = _as_text(row.get("format_rules_json", "{}"), default="{}")
        try:
            parsed_rules = json.loads(rules_raw) if rules_raw else {}
            if not isinstance(parsed_rules, dict):
                parsed_rules = {}
        except json.JSONDecodeError:
            parsed_rules = {}
        table_specs[object_id] = TableSpec(
            object_id=object_id,
            has_multilevel_headers=_as_bool(row.get("has_multilevel_headers", False)),
            format_rules={_as_text(k): _as_text(v) for k, v in parsed_rules.items()},
            row_label_col=_as_text(row.get("row_label_col", "")),
        )

    source_specs: dict[str, SourceSpec] = {}
    if not source_df.empty:
        for row in source_df.to_dict("records"):
            object_id = _as_text(row.get("object_id", ""))
            if object_id not in registry:
                continue
            source_specs[object_id] = SourceSpec(
                object_id=object_id,
                table_id=_as_text(row.get("table_id", "")),
                url=_as_text(row.get("url", "")),
                data_year=_as_int(row.get("data_year", 2023), default=2023),
                estimate_type=_as_text(row.get("estimate_type", "5-Year Estimates")),
                citation_month=_as_text(row.get("citation_month", "April")),
                citation_year=_as_int(row.get("citation_year", 2025), default=2025),
                custom_text=_as_text(row.get("custom_text", "")),
            )

    data_frames: dict[str, pd.DataFrame] = {}
    for record in registry.values():
        if record.data_sheet in sheets:
            data_frames[record.data_sheet] = sheets[record.data_sheet].copy()

    return WorkbookModel(
        workbook_path=source_path,
        registry=registry,
        figure_specs=figure_specs,
        table_specs=table_specs,
        source_specs=source_specs,
        data_frames=data_frames,
    )


def load_cha_workbook(workbook_path: str | Path) -> WorkbookModel:
    source_path = Path(workbook_path)
    sheets = _read_excel_raw(source_path)
    required = {"_registry", "_figure_specs", "_table_specs"}
    if required.issubset(sheets.keys()):
        return _load_normalized_workbook(source_path)
    return _load_flat_workbook(source_path)
