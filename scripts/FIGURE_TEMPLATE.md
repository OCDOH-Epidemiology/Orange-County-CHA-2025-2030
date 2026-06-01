# CHA Figure Template Guide

This guide shows how to build CHA figures with a consistent palette and
ordering, and (optionally) display a figure above its table.

## Quick Start (Line Figure)

```python
```{python}
#| echo: false
#| warning: false
#| message: false
#| label: fig-labor-force
#| fig-cap: "Percentage of the Labor Force, Population 16 Years and Older, 2021-2023"
import pandas as pd
from scripts.cha_figure_builder import build_line_figure

data = {
    "Year": [2021, 2022, 2023],
    "Dutchess": [63.0, 62.8, 63.3],
    "Orange": [63.9, 63.4, 63.5],
    "Putnam": [65.2, 64.7, 64.7],
    "Rockland": [63.9, 63.2, 63.2],
    "Sullivan": [56.1, 58.1, 59.1],
    "Ulster": [60.1, 58.9, 58.7],
    "Westchester": [65.5, 65.2, 65.4],
    "NYS": [63.1, 62.9, 63.0],
    "US": [63.6, 63.5, 63.5],
}

df = pd.DataFrame(data)
fig = build_line_figure(
    df,
    x_col="Year",
    y_axis_title="Percentage of Labor Force (%)",
    start_at_zero=False,
    hover_suffix="%",
)
fig.show()
```
```

## Figure Above Table (Single Chunk)

If you do not need separate labels for the figure and table, you can
render both in one chunk:

```python
```{python}
#| echo: false
#| warning: false
#| message: false
#| label: fig-and-table-example
#| fig-cap: "Example Figure"
import pandas as pd
from scripts.cha_figure_builder import build_line_figure, render_figure_and_table

data = {
    "Year": [2021, 2022, 2023],
    "Dutchess": [63.0, 62.8, 63.3],
    "Orange": [63.9, 63.4, 63.5],
    "Putnam": [65.2, 64.7, 64.7],
}
df = pd.DataFrame(data)
fig = build_line_figure(
    df,
    x_col="Year",
    y_axis_title="Percentage of Labor Force (%)",
    start_at_zero=False,
    hover_suffix="%",
)
render_figure_and_table(fig, df)
```
```

## Clustered Bar Figure

```python
```{python}
#| echo: false
#| warning: false
#| message: false
#| label: fig-clustered-example
#| fig-cap: "Clustered Bar Example"
import pandas as pd
from scripts.cha_figure_builder import build_clustered_bar_figure

data = {
    "Year": [2021, 2022, 2023],
    "Dutchess": [5.2, 5.0, 4.8],
    "Orange": [5.4, 5.3, 5.4],
    "Putnam": [4.8, 4.4, 4.1],
}
df = pd.DataFrame(data)
fig = build_clustered_bar_figure(
    df,
    x_col="Year",
    y_axis_title="Unemployment (%)",
    hover_suffix="%",
)
fig.show()
```
```

## Stacked Bar Figure

```python
```{python}
#| echo: false
#| warning: false
#| message: false
#| label: fig-stacked-example
#| fig-cap: "Stacked Bar Example"
import pandas as pd
from scripts.cha_figure_builder import build_stacked_bar_figure

data = {
    "Year": [2021, 2022, 2023],
    "Dutchess": [10, 12, 11],
    "Orange": [9, 10, 9],
    "Putnam": [7, 8, 7],
}
df = pd.DataFrame(data)
fig = build_stacked_bar_figure(
    df,
    x_col="Year",
    y_axis_title="Total (per 100,000)",
)
fig.show()
```
```

## Simple Bar Figure (Single Series)

```python
```{python}
#| echo: false
#| warning: false
#| message: false
#| label: fig-simple-bar-example
#| fig-cap: "Simple Bar Example"
import pandas as pd
from scripts.cha_figure_builder import build_simple_bar_figure

data = {
    "County": ["Dutchess", "Orange", "Putnam"],
    "Rate": [8.3, 13.0, 6.5],
}
df = pd.DataFrame(data)
fig = build_simple_bar_figure(
    df,
    x_col="County",
    y_cols=["Rate"],
    y_axis_title="Rate (%)",
    hover_suffix="%",
)
fig.show()
```
```

## Horizontal Bar Figure (Single Series)

```python
```{python}
#| echo: false
#| warning: false
#| message: false
#| label: fig-horizontal-bar-example
#| fig-cap: "Horizontal Bar Example"
import pandas as pd
from scripts.cha_figure_builder import build_horizontal_bar_figure

data = {
    "Health Issue": ["Mental Health", "Access to Healthcare Providers", "Substance Misuse"],
    "Percent": [68, 39, 31],
}
df = pd.DataFrame(data)
fig = build_horizontal_bar_figure(
    df,
    x_col="Health Issue",
    y_cols=["Percent"],
    y_axis_title="Percent (%)",
    hover_suffix="%",
)
fig.show()
```
```

## Notes

- If you need separate `fig-` and `tbl-` references, keep the figure and
  table in separate chunks (figure chunk above the table chunk).
- The builder orders county series by the standard CHA region order and
  uses the official color palette.
- For indicators that represent **currency**, keep the underlying data
  numeric but adjust display settings:
  - Use a y-axis title that mentions dollars (for example, `"Dollars ($)"`).
  - Set `hover_value_format` to `",.0f"` for whole-dollar values or `",.1f"`
    if you need one decimal place.
  - Optionally set `hover_suffix` to `" $"` or leave it blank if the axis
    title already makes the unit clear.
