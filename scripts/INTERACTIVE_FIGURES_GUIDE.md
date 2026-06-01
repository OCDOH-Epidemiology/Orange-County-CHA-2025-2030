# Interactive Figures Guide

This guide explains how to use interactive figures with county comparison capabilities in your CHA documents.

## Two Approaches

### 1. Enhanced Plotly Figures (Recommended for Quarto)

These work directly in Quarto documents without needing a separate server. They use Plotly's built-in dropdown menus and buttons.

**Usage in Quarto:**

```python
from scripts.cha_figure_builder import build_interactive_line_figure

# Your data
data = {
    "Year": [2021, 2022, 2023],
    "Dutchess": [63.0, 62.8, 63.3],
    "Orange": [63.9, 63.4, 63.5],
    # ... more counties
}

df = pd.DataFrame(data)

# Create interactive figure
fig = build_interactive_line_figure(
    df=df,
    x_col="Year",
    x_axis_title="Year",
    y_axis_title="Percentage of Labor Force (%)",
    start_at_zero=False,
    y_padding=0.1,
    hover_value_format=".1f",
    hover_suffix="%",
)

fig.show()
```

**Features:**
- **View Type Dropdown**: Switch between "Time Series (All Years)" and "Compare by Year (2021/2022/2023)"
- **County Filter Dropdown**: Filter to specific counties or preset groups:
  - All Counties
  - Individual counties
  - Mid-Hudson Counties (excludes NYS/US benchmarks)
  - Counties Only
  - With Benchmarks

**Advantages:**
- Works directly in Quarto HTML output
- No server required
- Interactive controls embedded in the figure
- Easy to use for readers

### 2. Standalone Dash App

A full-featured Dash application that runs as a separate web server with more advanced controls.

**Running the Dash App:**

```bash
# From project root
python scripts/cha_dash_app.py
```

Then open your browser to `http://localhost:8050`

**Features:**
- Radio buttons to switch between time series and single-year views
- Year dropdown selector (enabled/disabled based on view type)
- Multi-select checklist for counties
- More detailed controls and instructions

**Advantages:**
- More flexible UI controls
- Can be deployed as a standalone web application
- Better for complex filtering needs

**Customizing the Dash App:**

Edit `scripts/cha_dash_app.py` to:
- Change the data source
- Modify the layout
- Add additional filters or controls
- Customize styling

## Comparison

| Feature | Enhanced Plotly | Dash App |
|---------|----------------|----------|
| Works in Quarto | ✅ Yes | ❌ No (separate server) |
| Interactive Controls | ✅ Dropdowns | ✅ Radio buttons, dropdowns, checklists |
| Filter Counties | ✅ Yes | ✅ Yes |
| Compare by Year | ✅ Yes | ✅ Yes |
| Time Series View | ✅ Yes | ✅ Yes |
| Deployment | ✅ Static HTML | ⚠️ Requires server |
| Ease of Use | ✅ Simple | ⚠️ More complex |

## Recommendations

- **For Quarto documents**: Use `build_interactive_line_figure()` - it works seamlessly in your rendered HTML
- **For standalone web apps**: Use the Dash app - deploy it separately and link to it from your document
- **For maximum interactivity**: Consider embedding the Dash app in an iframe (requires hosting the Dash app)

## Example: Adding to Your Chapter

```python
```{python}
#| echo: false
#| warning: false
#| message: false
#| label: fig-my-interactive-figure
#| fig-cap: "Interactive Comparison: [Your Title]"
from scripts.cha_figure_builder import build_interactive_line_figure

# Your data here
df = pd.DataFrame({...})

fig = build_interactive_line_figure(
    df=df,
    x_col="Year",
    y_axis_title="Your Y-Axis Title",
    hover_suffix="%",  # or whatever unit you need
)

fig.show()
```
```

## Troubleshooting

**Dropdowns not appearing:**
- Make sure you're using `build_interactive_line_figure()` not `build_line_figure()`
- Check that your Quarto output format is HTML

**Bar chart view not working:**
- Ensure your x_col contains multiple unique values (years)
- Check that all county columns have data for the selected year

**Dash app won't start:**
- Make sure you have `dash` installed: `pip install dash`
- Check that you're running from the project root directory
- Verify port 8050 is not already in use
