# CHA Table Template Guide

This guide explains how to create consistently styled tables across all CHA chapters using the standardized template.

## Quick Start

### 1. Import the Styling Module

At the top of your Python code block, import the styling function:

```python
import pandas as pd
import numpy as np
from scripts.cha_table_styling import style_cha_table
```

### 2. Create Your Data

```python
# Create the data
data = {
    " ": ["Region A", "Region B", "Region C"],
    "Total Population": [100000, 200000, 300000],
    "Percent": [10.5, 20.3, 30.2]
}

df = pd.DataFrame(data)
```

### 3. Format Your Data (if needed)

```python
# Format numbers with commas
df["Total Population"] = df["Total Population"].apply(lambda x: f"{x:,}")

# Format percentages with one decimal
df["Percent"] = df["Percent"].apply(lambda x: f"{x:.1f}")

# Handle missing values
df["Percent"] = df["Percent"].apply(lambda x: f"{x:.1f}" if pd.notna(x) else "N/A")
```

### 4. Apply Styling and Display

```python
# Apply standard CHA styling
styled_table = style_cha_table(df)
styled_table
```

### 5. Add Table Label and Caption

Use Quarto chunk options:

```python
#| label: tbl-your-table-label
#| tbl-cap: "Table X: Your Table Title"
```

### 6. Add Source Citation (Collapsible) - STANDARD FORMAT

After your table code block, ALWAYS use this exact format:

````markdown
::: {.callout-note collapse="true"}
## Source

US Census Bureau; American Community Survey, 2023 American Community Survey 5-Year Estimates, [Table S0101](https://data.census.gov/table/...), April 2025
:::
````

**This is the STANDARD source format for all CHA tables.**

### Helper Function Examples

```python
from scripts.cha_table_styling import create_source_callout

# Standard Census Bureau table (2023 data)
source = create_source_callout(
    "S0101",  # Table ID
    "https://data.census.gov/table/ACSST5Y2023.S0101?q=s0101&g=..."  # Full URL
)

# Different year (2020 data)
source = create_source_callout(
    "S1810",
    "https://data.census.gov/table/...",
    data_year=2020
)

# Different estimate type (lowercase)
source = create_source_callout(
    "S1601",
    "https://data.census.gov/table/...",
    estimate_type="5-year estimates"
)

# Custom source (for non-Census sources from the CHA document)
source = create_source_callout(
    custom_text="New York State Department of Health, [Data Source](https://...), 2025"
)

print(source)  # Copy and paste into your document
```

**Note**: Always check the source document (2025 Regional CHA Document Orange County 12.17.2025.docx) for the correct table ID, year, and URL for each table.

## Complete Template

```python
```{python}
#| echo: false
#| warning: false
#| message: false
#| label: tbl-your-label
#| tbl-cap: "Table X: Your Table Title"
import pandas as pd
import numpy as np
from scripts.cha_table_styling import style_cha_table

# Create the data
data = {
    " ": ["Item 1", "Item 2", "Item 3"],
    "Column 1": [100, 200, 300],
    "Column 2": [10.5, 20.3, 30.2]
}

df = pd.DataFrame(data)

# Format the data (adjust as needed)
df["Column 1"] = df["Column 1"].apply(lambda x: f"{x:,}")
df["Column 2"] = df["Column 2"].apply(lambda x: f"{x:.1f}")

# Apply standard CHA styling
styled_table = style_cha_table(df)
styled_table
```

::: {.callout-note collapse="true"}
## Source

US Census Bureau; American Community Survey, 2023 American Community Survey 5-Year Estimates, [Table S0101](https://data.census.gov/table/...), April 2025
:::
```

## Styling Specifications

The `style_cha_table()` function applies the following consistent styling:

- **Header**: White background, bold text, centered
- **Row 1**: #EAF5DB (light green background)
- **Row 2**: White background
- **Row 3**: #EAF5DB (light green background)
- **Continues alternating**: white, #EAF5DB, white, #EAF5DB...
- **First column**: Bold text, left-aligned
- **Other columns**: Center-aligned
- **Borders**: 1px solid #ddd on all cells
- **Padding**: 10px on all cells

## Cross-Referencing Tables

To link to a table in your text, use:

```markdown
[see @tbl-your-label]
```

The label should match the `label:` in your chunk options (without the `tbl-` prefix in the chunk option, but with it in the reference).

## Source Citation Links

Always include clickable links to Census Bureau data tables:

```markdown
[Table S0101](https://data.census.gov/table/ACSST5Y2023.S0101?q=...)
```

## Using with Cursor AI

When asking Cursor to create a table, you can reference this template:

> "Create a table using the CHA table template from scripts/cha_table_styling.py with the following data..."

Or simply:

> "Create a CHA-styled table with [your data]"

Cursor will use the `style_cha_table()` function and follow the standard formatting.

## File Location

The styling module is located at:
`scripts/cha_table_styling.py`

Import it in any chapter with:
```python
from scripts.cha_table_styling import style_cha_table
```

