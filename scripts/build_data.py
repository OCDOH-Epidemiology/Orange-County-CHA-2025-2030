#!/usr/bin/env python3
"""Build processed datasets for the CHA book.

Expected input: data/raw/cha_metrics.xlsx
Sheets:
- metrics: columns = year, metric, value, county
- table: columns = indicator, county, value, year
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

RAW_FILE = Path("data/raw/cha_metrics.xlsx")
PROCESSED_DIR = Path("data/processed")


def sample_metrics() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "year": [2019, 2020, 2021, 2022],
            "metric": ["Physical Activity"] * 4,
            "value": [52.0, 49.5, 50.8, 53.2],
            "county": ["Orange County"] * 4,
        }
    )


def sample_table() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "indicator": [
                "Population",
                "Median Age",
                "Percent Hispanic",
                "Median Household Income",
            ],
            "county": [
                "Orange County",
                "Orange County",
                "Orange County",
                "Orange County",
            ],
            "value": [392000, 38.4, 25.1, 82500],
            "year": [2023, 2023, 2023, 2023],
        }
    )


def main() -> None:
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

    if RAW_FILE.exists():
        metrics = pd.read_excel(RAW_FILE, sheet_name="metrics")
        table = pd.read_excel(RAW_FILE, sheet_name="table")
    else:
        metrics = sample_metrics()
        table = sample_table()

    metrics.to_csv(PROCESSED_DIR / "cha_metrics.csv", index=False)
    table.to_csv(PROCESSED_DIR / "cha_table.csv", index=False)

    try:
        import pyarrow  # noqa: F401
    except ImportError:
        pyarrow = None

    if pyarrow is not None:
        metrics.to_parquet(PROCESSED_DIR / "cha_metrics.parquet", index=False)
        table.to_parquet(PROCESSED_DIR / "cha_table.parquet", index=False)

    print("Processed data written to data/processed/")


if __name__ == "__main__":
    main()
