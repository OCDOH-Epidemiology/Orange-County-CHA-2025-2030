"""
Helpers for metadata-driven chapter object rendering.

This module groups workbook registry objects into indicator-level bundles and
emits canonical Quarto Python blocks for figure/table/source rendering.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
from typing import Iterable

from scripts.workbook_loader import WorkbookModel, load_cha_workbook


def _resolve_default_workbook_path() -> Path:
    # Convention: keep the project workbook at data/raw/workbook.xlsx.
    return Path("data/raw/workbook.xlsx")


DEFAULT_WORKBOOK_PATH = _resolve_default_workbook_path()


@dataclass(frozen=True)
class IndicatorGroup:
    section_tag: str
    base_slug: str
    order_index: int
    caption: str
    figure_id: str | None
    table_id: str | None

    @property
    def include_file_slug(self) -> str:
        return self.base_slug


def _strip_object_prefix(object_id: str) -> str:
    if object_id.startswith("fig-"):
        return object_id[4:]
    if object_id.startswith("tbl-"):
        return object_id[4:]
    return object_id


def _load_model(workbook_path: str | Path | None) -> WorkbookModel:
    return load_cha_workbook(Path(workbook_path) if workbook_path else DEFAULT_WORKBOOK_PATH)


def get_indicator_groups(
    workbook_path: str | Path | None = None,
    *,
    section_filter: str = "",
) -> list[IndicatorGroup]:
    """
    Group workbook registry records into indicator bundles.

    Groups are formed by registry.section_tag + object ID base slug.
    """
    model = _load_model(workbook_path)
    grouped: dict[tuple[str, str], dict[str, object]] = {}

    for record in sorted(model.registry.values(), key=lambda row: row.order_index):
        if section_filter and section_filter.lower() not in record.section_tag.lower():
            continue

        base_slug = _strip_object_prefix(record.object_id)
        key = (record.section_tag, base_slug)
        slot = grouped.setdefault(
            key,
            {
                "section_tag": record.section_tag,
                "base_slug": base_slug,
                "order_index": record.order_index,
                "caption": record.caption,
                "figure_id": None,
                "table_id": None,
            },
        )

        if record.order_index < int(slot["order_index"]):
            slot["order_index"] = record.order_index
        if not slot["caption"] and record.caption:
            slot["caption"] = record.caption

        if record.object_id.startswith("fig-"):
            slot["figure_id"] = record.object_id
        elif record.object_id.startswith("tbl-"):
            slot["table_id"] = record.object_id

    indicators = [
        IndicatorGroup(
            section_tag=str(item["section_tag"]),
            base_slug=str(item["base_slug"]),
            order_index=int(item["order_index"]),
            caption=str(item["caption"]),
            figure_id=str(item["figure_id"]) if item["figure_id"] else None,
            table_id=str(item["table_id"]) if item["table_id"] else None,
        )
        for item in grouped.values()
    ]

    return sorted(indicators, key=lambda item: (item.section_tag.lower(), item.order_index, item.base_slug))


def validate_indicator_groups(groups: Iterable[IndicatorGroup]) -> list[str]:
    """
    Return a list of validation errors for indicator groups.
    """
    errors: list[str] = []
    order_seen: set[tuple[str, int]] = set()

    for group in groups:
        order_key = (group.section_tag, group.order_index)
        if order_key in order_seen:
            errors.append(
                f"Duplicate order_index={group.order_index} in section_tag='{group.section_tag}'."
            )
        order_seen.add(order_key)

        if not group.caption.strip():
            errors.append(
                f"Missing caption for section_tag='{group.section_tag}', slug='{group.base_slug}'."
            )

        if not group.figure_id and not group.table_id:
            errors.append(
                f"Indicator '{group.base_slug}' in section_tag='{group.section_tag}' has no figure or table."
            )

        if group.figure_id and not group.figure_id.startswith("fig-"):
            errors.append(f"Invalid figure_id '{group.figure_id}'. Expected prefix 'fig-'.")
        if group.table_id and not group.table_id.startswith("tbl-"):
            errors.append(f"Invalid table_id '{group.table_id}'. Expected prefix 'tbl-'.")

    return errors


def _yaml_single_quoted(text: str) -> str:
    # Quarto chunk options are YAML; single-quoted scalars avoid conflicts with
    # embedded double quotes in workbook captions.
    return "'" + text.replace("'", "''") + "'"


def _normalize_caption_dashes(caption: str) -> str:
    """
    Normalize caption punctuation to use en dashes only for numeric ranges.
    """
    return re.sub(
        r"(?<![A-Za-z0-9_/-])(\d{1,4})\s*-\s*(\d{1,4})(?![-A-Za-z0-9_/])",
        r"\1–\2",
        caption,
    )


def _default_fig_alt(caption: str) -> str:
    """
    Derive a reasonable default alt text from the figure caption.
    """
    text = caption.strip()
    if not text:
        return "Data visualization."
    # Drop a leading "Figure <n>:" prefix if present.
    text = text.split(":", 1)[-1].strip() if text.lower().startswith("figure ") else text
    return f"Data visualization showing {text}."


def _render_figure_block(figure_id: str, caption: str, workbook_var: str) -> str:
    normalized_caption = _normalize_caption_dashes(caption)
    return (
        "```{python}\n"
        "#| echo: false\n"
        "#| warning: false\n"
        "#| message: false\n"
        f"#| label: {figure_id}\n"
        f"#| fig-cap: {_yaml_single_quoted(normalized_caption)}\n"
        f"#| fig-alt: {_yaml_single_quoted(_default_fig_alt(normalized_caption))}\n"
        "from scripts.cha_registry_renderer import render_figure_object\n"
        f"render_figure_object(figure_id=\"{figure_id}\", workbook_path={workbook_var}).show()\n"
        "```"
    )


def _render_table_block(table_id: str, caption: str, workbook_var: str) -> str:
    normalized_caption = _normalize_caption_dashes(caption)
    return (
        "```{python}\n"
        "#| echo: false\n"
        "#| warning: false\n"
        "#| message: false\n"
        f"#| label: {table_id}\n"
        f"#| tbl-cap: {_yaml_single_quoted(normalized_caption)}\n"
        "from scripts.cha_registry_renderer import render_table_object\n"
        f"render_table_object(object_id=\"{table_id}\", workbook_path={workbook_var})\n"
        "```"
    )


def _render_source_block(object_id: str, workbook_var: str) -> str:
    return (
        "```{python}\n"
        "#| echo: false\n"
        "#| warning: false\n"
        "#| message: false\n"
        "from IPython.display import Markdown, display\n"
        "from scripts.cha_registry_renderer import render_source_callout_for_object\n"
        f"display(Markdown(render_source_callout_for_object(\"{object_id}\", {workbook_var})))\n"
        "```"
    )


def render_indicator_blocks(
    group: IndicatorGroup,
    *,
    workbook_var: str = "CHA_WORKBOOK_PATH",
    include_source: bool = True,
) -> str:
    """
    Render canonical Quarto blocks for a single indicator group.
    """
    parts: list[str] = [f"<!-- indicator: {group.base_slug} -->"]
    if group.figure_id:
        parts.append(_render_figure_block(group.figure_id, group.caption, workbook_var))
    if group.table_id:
        parts.append(_render_table_block(group.table_id, group.caption, workbook_var))
    return "\n\n".join(parts)


def _render_source_note_markdown(source_text: str = "", note_text: str = "") -> list[str]:
    """Return markdown callout blocks for sheet-level note metadata."""
    parts: list[str] = []
    if note_text:
        parts.append(
            "::: {.callout-note}\n"
            f"{note_text}\n"
            ":::"
        )
    return parts


def render_figure_blocks(
    figure_id: str,
    caption: str,
    *,
    workbook_var: str = "CHA_WORKBOOK_PATH",
    include_source: bool = False,
    source_text: str = "",
    note_text: str = "",
) -> str:
    """
    Render canonical Quarto blocks for a single figure object.
    """
    parts: list[str] = [f"<!-- indicator: {figure_id} -->", _render_figure_block(figure_id, caption, workbook_var)]
    parts.extend(_render_source_note_markdown(source_text, note_text))
    return "\n\n".join(parts)


def render_table_blocks(
    table_id: str,
    caption: str,
    *,
    workbook_var: str = "CHA_WORKBOOK_PATH",
    include_source: bool = True,
    source_text: str = "",
    note_text: str = "",
) -> str:
    """
    Render canonical Quarto blocks for a single table object.
    """
    parts: list[str] = [f"<!-- indicator: {table_id} -->", _render_table_block(table_id, caption, workbook_var)]
    parts.extend(_render_source_note_markdown(source_text, note_text))
    return "\n\n".join(parts)
