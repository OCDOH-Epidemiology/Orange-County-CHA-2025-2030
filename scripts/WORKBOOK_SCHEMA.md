# CHA Workbook Schema

Use this schema to drive CHA table and figure rendering from an Excel workbook.

## File format options

- **Supported**: single Excel workbook (`.xlsx`) with metadata + data sheets.
- **Note**: CSV-directory loading is not currently implemented in `scripts/workbook_loader.py`.

## Naming rules

- `object_id`: lowercase, hyphenated, unique across workbook.
  - Example: `fig-labor-force`, `tbl-unemployment`
- `object_type`: `table` or `figure`
- `label`: Quarto label, usually matches `object_id`
- `caption`: text used in `#| fig-cap` or `#| tbl-cap`
- `data_sheet`: exact sheet name containing source data for that object

## Required metadata sheets

### 1) `_registry` (required)

One row per renderable object.

Expected columns:

- `object_id`
- `object_type` (`table`/`figure`)
- `label`
- `caption`
- `data_sheet`
- `enabled` (`TRUE`/`FALSE`)
- `section_tag` (free text, e.g., `economic_stability`)
- `order_index` (integer sort order)

Column explanations:

- `object_id`: unique key for this object, used by `render_figure_object()` or `render_table_object()`.
- `object_type`: controls whether the row is routed to figure rendering or table rendering.
- `label`: Quarto cross-reference label (for example, `@fig-unemployment` or `@tbl-unemployment`).
- `caption`: display title used in figure/table captions.
- `data_sheet`: exact worksheet name containing source data for this object.
- `enabled`: set to `FALSE` to keep an object in the workbook without rendering it.
- `section_tag`: free-form grouping tag for organization (not required for rendering logic).
- `order_index`: numeric ordering key when generating/iterating content.

### 2) `_figure_specs` (required)

One row per figure object.

Expected columns:

- `object_id`
- `figure_type` (`line`/`clustered_bar`/`stacked_bar`/`simple_bar`/`horizontal_bar`/`horizontal_stacked_bar`/`horizontal_clustered_bar`/`pie`/`dot_whisker`)
- `x_col`
- `y_cols` (comma-separated column list)
- `x_axis_title`
- `y_axis_title`
- `start_at_zero` (`TRUE`/`FALSE`)
- `hover_suffix` (usually `%` or empty string)

Optional columns:

- `pivot_for_chart` (`TRUE`/`FALSE`, default `FALSE`)
- `group_by` (`x_col`/`series`, optional; overrides `pivot_for_chart` when provided)
- `show_data_labels` (`TRUE`/`FALSE`, optional; bar charts default to `FALSE` when blank)
- `color_by` (optional column name; colors horizontal bars by this grouping field)

Column explanations:

- `object_id`: links this row to `_registry.object_id` (must match exactly).
- `figure_type`: selects chart geometry.
  - `line` for trend lines
  - `clustered_bar` for grouped bars
  - `stacked_bar` for stacked totals
  - `simple_bar` for one-series bars
  - `horizontal_bar` for one-series horizontal bars
  - `horizontal_stacked_bar` for multi-series horizontal stacked bars
  - `horizontal_clustered_bar` for multi-series horizontal grouped bars
  - `pie` for pie charts
  - `dot_whisker` for horizontal dots with confidence-interval whiskers
- `x_col`: column used for x-axis after any pivot/reshape logic.
- `y_cols`: explicit series order; if blank, all non-`x_col` columns are used.
- `x_axis_title`: x-axis label text shown on the chart.
- `y_axis_title`: y-axis label text shown on the chart.
- `start_at_zero`: forces y-axis baseline to 0 when `TRUE`.
- `hover_suffix`: text appended in hover labels (for example, `%`).
- `pivot_for_chart`: when `TRUE`, reshapes wide data before plotting so chart orientation can change without editing source data layout.
- `group_by`: explicit grouping intent for bar charts.
- `show_data_labels`: controls on-bar value labels for `clustered_bar`, `stacked_bar`, and `simple_bar`.
  - when omitted/blank, bar charts hide labels by default
  - set `TRUE` to show labels for selected charts
  - `x_col`: group bars by values in `x_col` (counties become x-axis after pivot)
  - `series`: group bars by existing series columns (no pivot; `x_col` stays on axis)
- `color_by`: for long-format survey charts (category | group | value), colors each
  horizontal bar by the group column and builds a legend. Use with
  `horizontal_bar` or `horizontal_clustered_bar`. Keep `pivot_for_chart = FALSE`.
  Flat workbook key: `Color By`.

#### `x_col`, `group_by`, and `pivot_for_chart` behavior

- `x_col` chooses the x-axis field used by the chart renderer.
- `pivot_for_chart = TRUE` reshapes wide data before plotting.
- `group_by` is an easier orientation control for bar charts:
  - `group_by = x_col` is equivalent to `pivot_for_chart = TRUE`
  - `group_by = series` is equivalent to `pivot_for_chart = FALSE`
  - if both are present, `group_by` takes precedence

Common use case (counties on x-axis, grouped by period):

- Source data sheet columns:
  - first column: `Period`
  - remaining columns: `Dutchess`, `Orange`, `Putnam`, etc.
- Figure spec values:
  - `figure_type = clustered_bar`
  - `x_col = Period` (or `Year`)
  - `group_by = x_col` (or `pivot_for_chart = TRUE`)
- Result:
  - x-axis shows counties
  - each period becomes a separate series (grouped bars)

Flat per-indicator workbook key names map to the same settings:

- `X Column` -> `x_col`
- `Group By` -> `group_by`
- `Pivot For Chart` -> `pivot_for_chart`
- `Show Data Labels` -> `show_data_labels`
- `Color By` -> `color_by`

### 3) `_table_specs` (required)

One row per table object.

Expected columns:

- `object_id`
- `has_multilevel_headers` (`TRUE`/`FALSE`)
- `format_rules_json` (JSON object with column->format mapping)
- `row_label_col` (first/label column; informational)

Column explanations:

- `object_id`: links this row to `_registry.object_id` (must match exactly).
- `has_multilevel_headers`: set `TRUE` when table should render grouped/multi-row headers.
- `format_rules_json`: per-column display formatting rules (JSON object).
- `row_label_col`: identifies the left label column that should not be numerically formatted.

Supported formats in `format_rules_json`:

- `number` - comma separated, no decimals (e.g., `1,234`)
- `integer` - rounded integer, comma separated (e.g., `1,234`)
- `percent1` - one decimal place (e.g., `12.3`)
- `percent2` - two decimals (e.g., `12.34`)
- `currency` - dollar sign prefix + comma-separated thousands, no decimals (e.g., `$2,000`)
- `currency2` - dollar sign prefix + comma-separated thousands, **two** decimals (e.g., `$2,000.50`)
- `ratio` - table cells preserve the original ratio text (e.g., `1400:1`), while charts plot only the numeric segment before `:` (e.g., `1400`)
- `date` - renders as `MM/DD/YYYY` with 4-digit year (e.g., `03/10/2026`)
- `text` - leaves cell values unchanged (no numeric formatting)

All values in `format_rules_json` must match names understood by the table formatter in `scripts/cha_registry_renderer.py` (`_format_value`).

Examples:

- Simple percent formatting:
  - `{"Dutchess":"percent1","Orange":"percent1","US":"percent1"}`
- Mixed percent and currency formatting:
  - `{"Dutchess":"percent1","Orange":"currency","US":"currency"}`

### 4) `_source_specs` (optional but recommended)

One row per object that needs a source callout.

Expected columns:

- `object_id`
- `table_id`
- `url`
- `data_year`
- `estimate_type`
- `citation_month`
- `citation_year`
- `custom_text`

Column explanations:

- `object_id`: links source metadata to a figure/table object.
- `table_id`: source table identifier (for example, Census table code like `S2301`).
- `url`: source URL used in generated citation callouts.
- `data_year`: data vintage year shown in source text.
- `estimate_type`: estimate descriptor (for example, `5-Year Estimates`).
- `citation_month`: month name used in citation text.
- `citation_year`: year used in citation text.
- `custom_text`: full override citation text; when present, standard generated source text is skipped.

Use `custom_text` for non-Census sources. If `custom_text` is filled, it overrides standard generated citation text.

## Data sheets

- Add one sheet per dataset referenced by `_registry.data_sheet`.
- Column names must exactly match `x_col`, `y_cols`, and table format-rule keys.
- Keep region names aligned with CHA standard naming where possible.

For flat, per-sheet workbooks that follow `build_content_registry.py`, you can also set a sheet-level `value_format` in the CONFIG block:

- `value_format = "percent"` -> all data columns render as `percent1`
- `value_format = "number"` -> all data columns render as whole numbers with commas
- `value_format = "currency"` -> all data columns render as `$x,xxx`
- `value_format = "ratio"` -> tables keep full ratio text (e.g., `1400:1`), charts use only the numerator for plotting (e.g., `1400`)
- `value_format = "date"` -> all data columns render as `MM/DD/YYYY` (4-digit year)
- `value_format = "text"` or blank -> no automatic numeric formatting

Per-column overrides from `format_rules_json` (normalized workbooks) or format rows in flat sheets always take precedence over `value_format`.

### Flat per-indicator sheets: where format codes go

In the flat per-sheet workbook format (one indicator per sheet):

- The row where column **E** contains `Enter Data` is the header row for the data block.
- The row directly above `Enter Data` is reserved for per-column format codes:
  - Starting in column **G** and to the right, cells may contain:
    - `integer`
    - `number`
    - `percent1`
    - `percent2`
    - `currency`
    - `currency2`
    - `ratio`
    - `date`
    - `text`
  - These codes are parsed as formatting metadata only and never appear in figure axes or table headers.
- Data rows start on the row immediately below the `Enter Data` header row.

Important guardrails:

- Do not put `percent1`, `integer`, `currency`, etc. in the actual header row or data cells.
- If merged headers are used (e.g., `Male` / `Female` groups with a second header row), keep format codes in the dedicated format row above `Enter Data`; the parser will:
  - read that row into per-column format rules
  - exclude it from the data block entirely
  - strip accidental leading `percent1|` prefixes from labels before rendering

## Example object flow

1. Add row to `_registry`
2. Add row to `_figure_specs` or `_table_specs` with same `object_id`
3. Add data sheet named in `data_sheet`
4. (Optional) Add source row in `_source_specs`
5. In chapter code, call `render_figure_object(object_id, workbook_path)` or `render_table_object(object_id, workbook_path)`

## Validation behavior

Current fail-fast validation in `scripts/workbook_loader.py` is limited to:

- missing normalized metadata sheets (`_registry`, `_figure_specs`, `_table_specs`) when loading normalized mode

Other behaviors are permissive (non-fatal):

- rows with missing/invalid values are often skipped or defaulted
- many fields use defaults when blank (for example, `figure_type`, axis titles, booleans)
- unsupported `figure_type` values typically normalize to defaults during load, and only raise later if rendering still receives an unsupported type

Figure type values accepted by the normalizer:

- `1` = `line`
- `2` = `clustered_bar`
- `3` = `stacked_bar`
- `4` = `simple_bar` (single-series bar chart)
- `5` = `horizontal_bar` (single-series horizontal bar chart)
- `6` = `dot_whisker` (horizontal dot plot with CI whiskers)
- `7` = `horizontal_stacked_bar`
- `8` = `horizontal_clustered_bar`
- `9` = `pie`
