"""
generate_qmd_stubs.py

Reads all indicator sheets from the CHA workbook and generates ready-to-paste
QMD code stubs for every object_id that does NOT already appear in a target
chapter file.

Usage:
    # Show stubs for everything missing from the ch04 QMD:
    python scripts/generate_qmd_stubs.py

    # Target a different chapter:
    python scripts/generate_qmd_stubs.py --chapter chapters/05-health-behaviors.qmd

    # Filter by section_tag (object_id prefix slug):
    python scripts/generate_qmd_stubs.py --section food

    # Write stubs to a file:
    python scripts/generate_qmd_stubs.py --out stubs.md

Options:
    --workbook PATH     Excel workbook (default: main CHA workbook)
    --chapter PATH      QMD to check for existing labels (default: ch04)
    --section SLUG      Only include objects whose section_tag contains SLUG
    --out PATH          Write output to file instead of stdout
    --all               Include objects already present in the chapter file
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from scripts.workbook_loader import load_cha_workbook, RegistryRecord

DEFAULT_WORKBOOK = (
    PROJECT_ROOT
    / "data"
    / "raw"
    / "Mid-Hudson Regional Community Health Assessment 2025 Data File.xlsx"
)
DEFAULT_CHAPTER = PROJECT_ROOT / "chapters" / "04-Social and Physical Determinants of Health.qmd"

WORKBOOK_VAR = "CHA_WORKBOOK_PATH"

# ── template builders ──────────────────────────────────────────────────────────

def _fig_stub(record: RegistryRecord, caption: str) -> str:
    label = record.object_id
    return (
        f"```{{python}}\n"
        f"#| echo: false\n"
        f"#| warning: false\n"
        f"#| message: false\n"
        f"#| label: {label}\n"
        f'#| fig-cap: "{caption}"\n'
        f"from scripts.cha_registry_renderer import render_figure_object\n"
        f'render_figure_object(figure_id="{label}", workbook_path={WORKBOOK_VAR}).show()\n'
        f"```"
    )


def _tbl_stub(record: RegistryRecord, caption: str) -> str:
    label = record.object_id
    return (
        f"```{{python}}\n"
        f"#| echo: false\n"
        f"#| warning: false\n"
        f"#| message: false\n"
        f"#| label: {label}\n"
        f'#| tbl-cap: "{caption}"\n'
        f"from scripts.cha_registry_renderer import render_table_object\n"
        f'render_table_object(object_id="{label}", workbook_path={WORKBOOK_VAR})\n'
        f"```"
    )


def _source_stub(object_id: str) -> str:
    return (
        f"```{{python}}\n"
        f"#| echo: false\n"
        f"#| warning: false\n"
        f"#| message: false\n"
        f"from IPython.display import Markdown, display\n"
        f"from scripts.cha_registry_renderer import render_source_callout_for_object\n"
        f'display(Markdown(render_source_callout_for_object("{object_id}", {WORKBOOK_VAR})))\n'
        f"```"
    )


def _section_stub(section_tag: str, records: list[RegistryRecord], captions: dict[str, str]) -> str:
    lines = [f"\n<!-- ── {section_tag} ── -->"]

    # Figure block first (if present)
    fig_rec = next((r for r in records if r.object_type == "figure"), None)
    tbl_rec = next((r for r in records if r.object_type == "table"), None)

    if fig_rec:
        caption = captions.get(fig_rec.object_id, fig_rec.caption)
        lines.append(_fig_stub(fig_rec, caption))

    if tbl_rec:
        caption = captions.get(tbl_rec.object_id, tbl_rec.caption)
        lines.append(_tbl_stub(tbl_rec, caption))

        # Source callout after the table
        lines.append(_source_stub(tbl_rec.object_id))

    return "\n\n".join(lines)


# ── main ──────────────────────────────────────────────────────────────────────

def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Generate QMD stubs for CHA indicators.")
    parser.add_argument("--workbook", default=str(DEFAULT_WORKBOOK), help="Excel workbook path")
    parser.add_argument("--chapter", default=str(DEFAULT_CHAPTER), help="Target QMD file to check")
    parser.add_argument("--section", default="", help="Filter by section_tag substring")
    parser.add_argument("--out", default="", help="Output file path (default: stdout)")
    parser.add_argument("--all", action="store_true", dest="include_all",
                        help="Include objects already present in the chapter")
    args = parser.parse_args(argv)

    workbook_path = Path(args.workbook)
    chapter_path = Path(args.chapter)

    if not workbook_path.exists():
        print(f"ERROR: Workbook not found: {workbook_path}", file=sys.stderr)
        sys.exit(1)

    model = load_cha_workbook(workbook_path)

    # Read existing labels from the chapter file
    existing_labels: set[str] = set()
    if chapter_path.exists() and not args.include_all:
        import re
        chapter_text = chapter_path.read_text(encoding="utf-8")
        existing_labels = set(re.findall(r"#\|\s*label:\s*(\S+)", chapter_text))

    # Group records by section_tag, preserving order_index sort
    from collections import defaultdict
    sections: dict[str, list[RegistryRecord]] = defaultdict(list)
    for record in sorted(model.registry.values(), key=lambda r: r.order_index):
        if args.section and args.section.lower() not in record.section_tag.lower():
            continue
        sections[record.section_tag].append(record)

    # Build caption map from source_specs + registry
    captions: dict[str, str] = {}
    for oid, record in model.registry.items():
        captions[oid] = record.caption

    output_parts: list[str] = []

    header_lines = [
        "<!-- QMD Stubs generated by scripts/generate_qmd_stubs.py -->",
        f"<!-- Workbook: {workbook_path.name} -->",
        f"<!-- Chapter:  {chapter_path.name} -->",
        "",
    ]

    skipped = 0
    included = 0

    for section_tag, records in sections.items():
        # Check if any record in this section is missing from the chapter
        missing = [r for r in records if r.object_id not in existing_labels]
        if not args.include_all and not missing:
            skipped += 1
            continue

        target_records = records if args.include_all else missing
        stub = _section_stub(section_tag, target_records, captions)
        output_parts.append(stub)
        included += 1

    summary = [
        f"<!-- {included} section(s) with missing objects -->",
        f"<!-- {skipped} section(s) already fully present in chapter (skipped) -->",
    ]

    full_output = "\n".join(header_lines + summary + [""] + output_parts)

    if args.out:
        out_path = Path(args.out)
        out_path.write_text(full_output, encoding="utf-8")
        print(f"Written to {out_path}  ({included} sections)")
    else:
        print(full_output)


if __name__ == "__main__":
    main()
