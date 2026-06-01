"""
transform_ch04_qmd.py

Replaces hardcoded Python figure/table code blocks in the ch04 QMD with
compact registry calls using render_figure_object() / render_table_object().

Only blocks whose labels match known migrated indicator IDs are replaced.
All narrative text, callout notes, and source callout markdown are preserved.

Run from the project root:
    python scripts/transform_ch04_qmd.py
"""

from __future__ import annotations
import re
from pathlib import Path

QMD_PATH = Path("chapters/04-Social and Physical Determinants of Health.qmd")

# Object IDs now in the workbook — used to decide which blocks to replace.
# Key = quarto label (e.g. "fig-food-insecurity"), value = object_id for renderer.
# The label and object_id are always the same in our system.
MIGRATED_IDS: set[str] = {
    "fig-labor-force",
    "tbl-labor-force",
    "fig-unemployed",
    "tbl-unemployed",
    "fig-food-insecurity",
    "tbl-food-insecurity",
    "fig-childhood-food-insecurity",
    "tbl-childhood-food-insecurity",
    "fig-cost-burdened-renters",
    "tbl-cost-burdened-renters",
    "fig-severely-cost-burdened",
    "tbl-severely-cost-burdened",
    "fig-hud-housing",
    "tbl-hud-housing",
    "tbl-poverty-threshold",
    "fig-poverty-rate",
    "tbl-poverty-rate",
    "fig-poverty-by-race",
    "tbl-poverty-by-race",
    "fig-economically-disadvantaged",
    "tbl-economically-disadvantaged",
    "tbl-alice-budget",
    "fig-alice-threshold",
    "tbl-alice-threshold",
    "fig-graduation-rate",
    "tbl-graduation-rate",
    "fig-graduation-by-race",
    "tbl-graduation-by-race",
    "fig-educational-attainment",
    "tbl-educational-attainment",
    "fig-language-proficiency",
    "tbl-language-proficiency",
    "fig-disconnected-youth",
    "tbl-disconnected-youth",
    "tbl-residential-segregation",
    "fig-food-environment-index",
    "tbl-food-environment-index",
    "fig-limited-access-healthy-foods",
    "tbl-limited-access-healthy-foods",
    "fig-violent-crime",
    "tbl-violent-crime",
    "fig-air-pollution",
    "tbl-air-pollution",
    "fig-lead-testing",
    "tbl-lead-testing",
    "fig-severe-housing-problems",
    "tbl-severe-housing-problems",
    "fig-no-vehicles",
    "tbl-no-vehicles",
    "fig-modes-transportation",
    "tbl-modes-transportation",
    "fig-commute-time",
    "tbl-commute-time",
    # Health care access section
    "fig-residential-segregation",
    "fig-medical-care-cost",
    "tbl-medical-care-cost",
    "fig-children-uninsured",
    "tbl-children-uninsured",
    "fig-adults-insured",
    "tbl-adults-insured",
    "tbl-medically-underserved",
    "fig-primary-care-providers",
    "tbl-primary-care-providers",
    "fig-dentists",
    "tbl-dentists",
    "fig-mental-health-providers",
    "tbl-mental-health-providers",
    "fig-access-to-primary-care",
    "tbl-access-to-primary-care",
}

# Captions for each object — used to keep the Quarto #| fig-cap / tbl-cap header.
CAPTIONS: dict[str, str] = {
    "fig-labor-force": "Percentage of the Labor Force, Population 16 Years and Older, 2021-2023",
    "tbl-labor-force": "Percentage of the Labor Force, Population 16 Years and Older, 2021-2023",
    "fig-unemployed": "Percentage of Population Unemployed, 16 Years and Older, 2021-2023",
    "tbl-unemployed": "Percentage of Population Unemployed, 16 Years and Older, 2021-2023",
    "fig-food-insecurity": "Percentage of Overall Food Insecurity, 2020-2023",
    "tbl-food-insecurity": "Percentage of Overall Food Insecurity, 2020-2023",
    "fig-childhood-food-insecurity": "Percentage of Food Insecurity, Children 18 Years and Younger, 2020-2023",
    "tbl-childhood-food-insecurity": "Percentage of Food Insecurity, Children 18 Years and Younger, 2020-2023",
    "fig-cost-burdened-renters": "Percentage of Cost Burdened Renter Occupied Units, 2021-2023",
    "tbl-cost-burdened-renters": "Cost Burdened Renter Occupied Units, 2021-2023",
    "fig-severely-cost-burdened": "Percentage of Severely Cost Burdened Households, 2016-2023",
    "tbl-severely-cost-burdened": "Severely Cost Burdened Households, 2016-2020 to 2019-2023",
    "fig-hud-housing": "Number of People Living in HUD-Subsidized Housing in the Past 12 Months, 2021\u20132024",
    "tbl-hud-housing": "Number of People Living in Housing and Urban Development (HUD)-Subsidized Housing in the Past 12 Months, 2021\u20132024",
    "tbl-poverty-threshold": "Poverty Threshold for 2024 by Size of Family and Number of Related Children 18 Years and Younger",
    "fig-poverty-rate": "Percentage of Population in Poverty, 2021-2023",
    "tbl-poverty-rate": "Percentage of Population in Poverty, 2021-2023",
    "fig-poverty-by-race": "Percentage of Families Below Poverty Level by Race and Ethnicity, 2023",
    "tbl-poverty-by-race": "Percentage of Families Below Poverty Level by Race and Ethnicity, 2023",
    "fig-economically-disadvantaged": "Enrollment Rate of Economically Disadvantaged Students, 2021-2024",
    "tbl-economically-disadvantaged": "Enrollment Rate of Economically Disadvantaged Students, 2021-2024",
    "tbl-alice-budget": "ALICE Household Survival Budget, New York State, 2023",
    "fig-alice-threshold": "ALICE Threshold Percentage, 2023",
    "tbl-alice-threshold": "ALICE Threshold Percentage, 2023",
    "fig-graduation-rate": "High School Graduation Rate, 2021-2023",
    "tbl-graduation-rate": "High School Graduation Rate, 2021-2023",
    "fig-graduation-by-race": "High School Graduation Rate, by Race and Ethnicity, 2023",
    "tbl-graduation-by-race": "High School Graduation Rate, by Race and Ethnicity, 2023",
    "fig-educational-attainment": "Rate of Education Attainment, 2023",
    "tbl-educational-attainment": "Rate of Education Attainment, 2023",
    "fig-language-proficiency": "Percentage of Population that Speaks English Less Than Very Well, 5 Years and Older, 2021-2023",
    "tbl-language-proficiency": "Percentage of Population that Speaks English Less Than Very Well, 5 Years and Older, 2021-2023",
    "fig-disconnected-youth": "Percentage of Disconnected Youth Ages 16-19, 2014-2023",
    "tbl-disconnected-youth": "Percentage of Disconnected Youth Ages 16-19, 2014-2023",
    "tbl-residential-segregation": "Index Score of Residential Segregation, 2013-2023",
    "fig-food-environment-index": "Food Environment Index, 2025",
    "tbl-food-environment-index": "Food Environment Index, 2025",
    "fig-limited-access-healthy-foods": "Percentage of Population with Limited Access to Healthy Foods, 2025",
    "tbl-limited-access-healthy-foods": "Percentage of Population with Limited Access to Healthy Foods, 2025",
    "fig-violent-crime": "Violent Crime Rate per 100,000 Population, 2018-2021",
    "tbl-violent-crime": "Violent Crime Rate per 100,000 Population, 2018-2021",
    "fig-air-pollution": "Average Daily Density of Fine Particulate Matter, 2014-2020",
    "tbl-air-pollution": "Average Daily Density of Fine Particulate Matter, 2014-2020",
    "fig-lead-testing": "Percentage of Children Tested for Lead at Least Twice Before 36 Months of Age, 2016-2019 Birth Cohorts",
    "tbl-lead-testing": "Percentage of Children Tested for Lead at Least Twice Before 36 Months of Age, 2016-2019 Birth Cohorts",
    "fig-severe-housing-problems": "Percentage of Households with Severe Housing Problems, 2025",
    "tbl-severe-housing-problems": "Percentage of Households with Severe Housing Problems, 2025",
    "fig-no-vehicles": "Percentage of Households with No Available Vehicles, 2023",
    "tbl-no-vehicles": "Percentage of Households with No Available Vehicles, 2023",
    "fig-modes-transportation": "Modes of Transportation to Work, 2023",
    "tbl-modes-transportation": "Modes of Transportation to Work, 2023",
    "fig-commute-time": "Average Commute Time to Work, 2023",
    "tbl-commute-time": "Average Commute Time to Work, 2023",
    # Health care access section
    "fig-residential-segregation": "Index Score of Residential Segregation, 2013-2023",
    "fig-medical-care-cost": "Percentage of Adults Who Did Not Receive Medical Care Due to Cost, 2016, 2018, 2021",
    "tbl-medical-care-cost": "Percentage of Adults Who Did Not Receive Medical Care Due to Cost, 2016, 2018, 2021",
    "fig-children-uninsured": "Percentage of Children Without Health Insurance, 2020-2023",
    "tbl-children-uninsured": "Percentage of Children Without Health Insurance, 2020-2023",
    "fig-adults-insured": "Percentage of Adults with Health Insurance, 2020-2023",
    "tbl-adults-insured": "Percentage of Adults with Health Insurance, 2020-2023",
    "tbl-medically-underserved": "Medically Underserved Areas and Medically Underserved Populations",
    "fig-primary-care-providers": "Ratio of Residents to Primary Care Providers, 2019-2021",
    "tbl-primary-care-providers": "Ratio of Residents to Primary Care Providers, 2019-2021",
    "fig-dentists": "Ratio of Residents to Dentists, 2019-2022",
    "tbl-dentists": "Ratio of Residents to Dentists, 2019-2022",
    "fig-mental-health-providers": "Ratio of Residents to Mental Health Providers, 2021-2024",
    "tbl-mental-health-providers": "Ratio of Residents to Mental Health Providers, 2021-2024",
    "fig-access-to-primary-care": "Percentage of Adults Who Have a Regular Primary Care Provider, 2016, 2018, and 2021",
    "tbl-access-to-primary-care": "Percentage of Adults Who Have a Regular Primary Care Provider, 2016, 2018, and 2021",
}

# The path setup block that appears as a setup chunk at the top of the file.
# We keep this; each indicator chunk does NOT need to repeat it.
WORKBOOK_PATH_SETUP = "CHA_WORKBOOK_PATH"


def _make_fig_body(label: str, caption: str) -> str:
    cap_line = f'#| fig-cap: "{caption}"'
    return (
        "#| echo: false\n"
        "#| warning: false\n"
        "#| message: false\n"
        f"#| label: {label}\n"
        f"{cap_line}\n"
        "from scripts.cha_registry_renderer import render_figure_object\n"
        f'render_figure_object(figure_id="{label}", workbook_path={WORKBOOK_PATH_SETUP}).show()\n"'
    ).rstrip('"')


def _make_tbl_body(label: str, caption: str) -> str:
    cap_line = f'#| tbl-cap: "{caption}"'
    return (
        "#| echo: false\n"
        "#| warning: false\n"
        "#| message: false\n"
        f"#| label: {label}\n"
        f"{cap_line}\n"
        "from scripts.cha_registry_renderer import render_table_object\n"
        f'render_table_object(object_id="{label}", workbook_path={WORKBOOK_PATH_SETUP})\n"'
    ).rstrip('"')


def _build_replacement_block(label: str) -> str:
    """Return the full ```{python} ... ``` replacement for a given label."""
    caption = CAPTIONS.get(label, label)
    if label.startswith("fig-"):
        body = _make_fig_body(label, caption)
    else:
        body = _make_tbl_body(label, caption)
    return f"```{{python}}\n{body}\n```"


def transform(text: str) -> str:
    """
    Parse the QMD text into segments and replace qualifying Python code blocks.

    A qualifying block:
    - starts with ```{python}
    - contains a ``#| label:`` line whose value is in MIGRATED_IDS
    - does NOT already contain render_figure_object / render_table_object
      (i.e., is still hardcoded)
    """
    # Pattern: ```{python}  ...content...  ```
    # Use a non-greedy match.  We rely on ``` being at start of line.
    code_block_re = re.compile(
        r"(```\{python\}\n)(.*?)(```)",
        re.DOTALL,
    )

    label_re = re.compile(r"#\|\s*label:\s*(\S+)")

    replaced = 0

    def replacer(m: re.Match) -> str:
        nonlocal replaced
        opener = m.group(1)   # ```{python}\n
        body   = m.group(2)   # block content
        closer = m.group(3)   # ```

        # Find label
        lm = label_re.search(body)
        if not lm:
            return m.group(0)

        label = lm.group(1)
        if label not in MIGRATED_IDS:
            return m.group(0)

        # Skip if already uses registry call
        if "render_figure_object" in body or "render_table_object" in body:
            return m.group(0)

        replaced += 1
        return _build_replacement_block(label)

    result = code_block_re.sub(replacer, text)
    print(f"  Replaced {replaced} hardcoded code blocks.")
    return result


def fix_stray_backticks(text: str) -> str:
    """
    Fix the known typo at line ~88 where a stray 't' sits before triple
    backticks, leaving a malformed non-block.  Replace ``t``` `` with nothing.
    """
    # Matches a line that is literally just  t```  (with optional surrounding whitespace)
    fixed = re.sub(r"^t```\s*$", "", text, flags=re.MULTILINE)
    if fixed != text:
        print("  Fixed stray 't```' typo.")
    return fixed


def add_fig_unemployment_block(text: str) -> str:
    """
    The original QMD was missing a fig-unemployed code block (only tbl-unemployed
    existed).  Insert it before the tbl-unemployment block if it's not present.
    """
    if "fig-unemployed" in text:
        return text

    fig_block = _build_replacement_block("fig-unemployed")
    # Insert before the tbl-unemployment block
    tbl_marker = "```{python}\n#| echo: false\n#| warning: false\n#| message: false\n#| label: tbl-unemployment"
    if tbl_marker in text:
        text = text.replace(tbl_marker, fig_block + "\n\n" + tbl_marker, 1)
        print("  Inserted missing fig-unemployed block.")
    return text


def main() -> None:
    if not QMD_PATH.exists():
        raise FileNotFoundError(f"QMD not found: {QMD_PATH}")

    print(f"Transforming: {QMD_PATH}")
    original = QMD_PATH.read_text(encoding="utf-8")

    text = fix_stray_backticks(original)
    text = add_fig_unemployment_block(text)
    text = transform(text)

    QMD_PATH.write_text(text, encoding="utf-8")
    print(f"Done.  Saved {QMD_PATH}")


if __name__ == "__main__":
    main()
