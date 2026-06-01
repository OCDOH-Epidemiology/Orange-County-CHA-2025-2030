# STANDARD SOURCE FORMAT FOR CHA TABLES

## ⚠️ THIS IS THE OFFICIAL STANDARD - USE FOR ALL TABLES

All CHA tables MUST use this exact source citation format:

```markdown
::: {.callout-note collapse="true"}
## Source

US Census Bureau; American Community Survey, 2023 American Community Survey 5-Year Estimates, [Table S0101](https://data.census.gov/table/...), April 2025
:::
```

## Format Components

1. **Collapsible callout box**: `::: {.callout-note collapse="true"}`
2. **Source header**: `## Source`
3. **Citation text**: 
   - "US Census Bureau; American Community Survey,"
   - Year (e.g., "2023")
   - "American Community Survey 5-Year Estimates,"
   - Table ID as clickable link: `[Table S0101](url)`
   - Citation date: "April 2025"

## Template

```
::: {.callout-note collapse="true"}
## Source

US Census Bureau; American Community Survey, [YEAR] American Community Survey 5-Year Estimates, [Table [TABLE_ID]]([URL]), [MONTH] [YEAR]
:::
```

## Examples

### Table S0101
```markdown
::: {.callout-note collapse="true"}
## Source

US Census Bureau; American Community Survey, 2023 American Community Survey 5-Year Estimates, [Table S0101](https://data.census.gov/table/ACSST5Y2023.S0101?q=s0101&g=...), April 2025
:::
```

### Table B03002
```markdown
::: {.callout-note collapse="true"}
## Source

US Census Bureau; American Community Survey, 2023 American Community Survey 5-Year Estimates, [Table B03002](https://data.census.gov/table/ACSDT5Y2023.B03002?q=b03002&g=...), April 2025
:::
```

## Using the Helper Function

You can generate the source callout programmatically:

```python
from scripts.cha_table_styling import create_source_callout

source = create_source_callout(
    "S0101",  # Table ID
    "https://data.census.gov/table/ACSST5Y2023.S0101?q=s0101&g=..."  # Full URL
)
print(source)
```

## Why This Format?

- **Consistent**: Same format across all tables
- **Collapsible**: Doesn't clutter the page
- **Clickable**: Direct links to Census Bureau data
- **Professional**: Standard academic citation format
- **Accessible**: Easy to find and expand when needed

## Variations Based on Source Document

The source information will vary based on:
1. **Table ID**: Different Census tables (S0101, B03002, S1601, S1501, S1901, S2101, S1810, etc.)
2. **Data Year**: May be 2023, 2020, or other years
3. **Estimate Type**: "5-Year Estimates" (capital) or "5-year estimates" (lowercase)
4. **URL**: Each table has a unique Census Bureau URL
5. **Other Sources**: Some tables may reference NYS Department of Health or other sources

**Always check the source document**: `2025 Regional CHA Document Orange County 12.17.2025.docx` for the correct citation information for each table.

## For Cursor AI

When asking Cursor to create tables, it will automatically use this standard format because it's documented in:
- `.cursorrules`
- `scripts/TABLE_TEMPLATE.md`
- `scripts/cha_table_styling.py`

Just say: "Create a CHA table with source citation" and Cursor will use this format. Make sure to provide the correct table ID, year, and URL from the source document.

