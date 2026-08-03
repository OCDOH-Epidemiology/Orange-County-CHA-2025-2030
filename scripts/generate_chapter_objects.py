"""
Generate metadata-driven include files for chapter indicator objects.

This script writes one include file per workbook object ID (fig/tbl) and can
optionally rewrite a chapter QMD so repetitive object code blocks are replaced
with include directives.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from scripts.cha_chapter_renderer import (  # noqa: E402
    DEFAULT_WORKBOOK_PATH,
    get_indicator_groups,
    render_figure_blocks,
    render_table_blocks,
    validate_indicator_groups,
)
from scripts.workbook_loader import load_cha_workbook  # noqa: E402


DEFAULT_CHAPTER = PROJECT_ROOT / "chapters" / "13-Triangulate the Data.qmd"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "chapters" / "_generated" / "objects"
DEFAULT_CHAPTERS_DIR = PROJECT_ROOT / "chapters"


_CROSSREF_RE = re.compile(r"\[@(fig-[A-Za-z0-9+\-]+|tbl-[A-Za-z0-9+\-]+)\]")
_PY_BLOCK_RE = re.compile(r"```{python}\n.*?```", flags=re.DOTALL)
_LABEL_RE = re.compile(r"#\|\s*label:\s*([A-Za-z0-9+\-]+)")
_SOURCE_OBJ_RE = re.compile(r"render_source_callout_for_object\(\"([^\"]+)\"")
_FIG_CALL_RE = re.compile(r"render_figure_object\(\s*figure_id=\"([^\"]+)\"")
_TBL_CALL_RE = re.compile(r"render_table_object\(\s*object_id=\"([^\"]+)\"")
_INCLUDE_STEM_RE = re.compile(r"_generated/objects/((?:fig|tbl)-[A-Za-z0-9+\-]+)\.qmd")
_DUPLICATE_ORDER_ERROR_PREFIX = "Duplicate order_index="

# ch06-health-behaviors.qmd supplies markdown Source callouts for these; omit
# registry-driven render_source_callout chunks to avoid triple duplicates (fig
# + tbl + chapter) when fig/tbl live in separate indicator groups.
OBJECT_IDS_OMIT_EMBEDDED_SOURCE: frozenset[str] = frozenset(
    {
        "fig-brfss-leisure-pa",
        "tbl-brfss-adult-leisure-pa",
        "fig-brfss-fruit-veg",
        "tbl-brfss-adult-fruit-veg",
        "fig-brfss-sug-bev",
        "tbl-brfss-adult-sug-bev",
    }
)


def _iter_chapter_qmd_files(chapters_dir: Path):
    """Yield chapter files while skipping macOS sidecar metadata files."""
    for chapter_file in chapters_dir.glob("*.qmd"):
        if chapter_file.name.startswith("._"):
            continue
        yield chapter_file


def _include_stems_from_chapters(chapters_dir: Path) -> set[str]:
    """Collect fig-/tbl- stems from all chapter shortcode includes (any .qmd in chapters/)."""
    stems: set[str] = set()
    for chapter_file in _iter_chapter_qmd_files(chapters_dir):
        stems.update(_INCLUDE_STEM_RE.findall(chapter_file.read_text(encoding="utf-8")))
    return stems


def _file_hashes(output_dir: Path) -> dict[str, str]:
    import hashlib

    hashes: dict[str, str] = {}
    if not output_dir.exists():
        return hashes
    for include_file in output_dir.glob("*.qmd"):
        if include_file.name.startswith("._"):
            continue
        digest = hashlib.sha256(include_file.read_bytes()).hexdigest()
        hashes[include_file.stem] = digest
    return hashes


def _summarize_generation(
    before: dict[str, str],
    after: dict[str, str],
) -> tuple[list[str], list[str], list[str]]:
    before_stems = set(before)
    after_stems = set(after)
    added = sorted(after_stems - before_stems)
    removed = sorted(before_stems - after_stems)
    updated = sorted(stem for stem in before_stems & after_stems if before[stem] != after[stem])
    return added, removed, updated


def _print_generation_summary(
    added: list[str],
    removed: list[str],
    updated: list[str],
    total: int,
    output_dir: Path,
) -> None:
    print(f"Wrote {total} include files to {output_dir}")
    if added:
        print(f"  Added ({len(added)}): {', '.join(added)}")
    if updated:
        print(f"  Updated ({len(updated)}): {', '.join(updated)}")
    if removed:
        print(f"  Removed ({len(removed)}): {', '.join(removed)}")
    if not added and not updated and not removed:
        print("  No object files added, removed, or changed.")


def _validate_chapter_references(chapter_text: str, registry_ids: set[str]) -> list[str]:
    errors: list[str] = []
    referenced_ids = set(_CROSSREF_RE.findall(chapter_text))
    missing = sorted(referenced_ids - registry_ids)
    for object_id in missing:
        errors.append(f"Chapter cross-reference points to missing object_id '{object_id}'.")

    # Pairing sanity check: if both figure and table are referenced for the same
    # slug, both must exist in the workbook registry.
    slug_refs: dict[str, set[str]] = {}
    for object_id in referenced_ids:
        kind, slug = object_id.split("-", 1)
        slug_refs.setdefault(slug, set()).add(kind)
    for slug, kinds in slug_refs.items():
        if {"fig", "tbl"}.issubset(kinds):
            fig_id = f"fig-{slug}"
            tbl_id = f"tbl-{slug}"
            if fig_id not in registry_ids or tbl_id not in registry_ids:
                errors.append(
                    f"Broken fig/tbl pairing for slug '{slug}' (expected '{fig_id}' and '{tbl_id}' in workbook)."
                )
    return errors


def _validate_include_directives(
    chapter_text: str, output_dir: Path, written_stems: set[str]
) -> list[str]:
    """Check that include directives point to generated files that actually exist."""
    errors: list[str] = []
    include_stems = set(_INCLUDE_STEM_RE.findall(chapter_text))
    for stem in sorted(include_stems):
        target = output_dir / f"{stem}.qmd"
        if stem not in written_stems and not target.exists():
            errors.append(
                f"Include directive references '{stem}.qmd' but no generated file exists. "
                f"Available stems with similar prefix: "
                f"{sorted(s for s in written_stems if s.split('-')[0] == stem.split('-')[0])}"
            )
    return errors


def _validate_all_chapter_references(
    chapters_dir: Path,
    registry_ids: set[str],
    output_dir: Path,
    written_stems: set[str],
) -> list[str]:
    errors: list[str] = []
    for chapter_file in sorted(_iter_chapter_qmd_files(chapters_dir)):
        chapter_text = chapter_file.read_text(encoding="utf-8")
        chapter_errors = _validate_chapter_references(chapter_text, registry_ids)
        for err in chapter_errors:
            errors.append(f"{chapter_file.name}: {err}")
        include_errors = _validate_include_directives(chapter_text, output_dir, written_stems)
        for err in include_errors:
            errors.append(f"{chapter_file.name}: {err}")
    return errors


def _replace_python_blocks_with_includes(
    chapter_text: str,
    known_file_stems: set[str],
    include_prefix: str,
    output_dir: Path,
) -> str:
    inserted_includes: set[str] = set()

    def _include_line(file_stem: str) -> str:
        return f"{{{{< include {include_prefix}/{file_stem}.qmd >}}}}"

    def replacer(match: re.Match[str]) -> str:
        block = match.group(0)
        label_match = _LABEL_RE.search(block)

        object_candidates: list[str] = []
        if label_match:
            object_candidates.append(label_match.group(1))

        fig_call_match = _FIG_CALL_RE.search(block)
        if fig_call_match:
            object_candidates.append(fig_call_match.group(1))
        tbl_call_match = _TBL_CALL_RE.search(block)
        if tbl_call_match:
            object_candidates.append(tbl_call_match.group(1))

        for candidate in object_candidates:
            if candidate.startswith(("fig-", "tbl-")):
                if candidate in known_file_stems:
                    if candidate in inserted_includes:
                        return ""
                    inserted_includes.add(candidate)
                    return _include_line(candidate)
                # Fallback: keep unknown object blocks by extracting them to
                # generated include files named with the full object label.
                fallback_stem = candidate
                fallback_path = output_dir / f"{fallback_stem}.qmd"
                if not fallback_path.exists():
                    fallback_path.write_text(block + "\n", encoding="utf-8")
                if fallback_stem in inserted_includes:
                    return ""
                inserted_includes.add(fallback_stem)
                return _include_line(fallback_stem)

        source_match = _SOURCE_OBJ_RE.search(block)
        if source_match:
            source_id = source_match.group(1)
            if source_id in known_file_stems:
                return ""

        return block

    updated = _PY_BLOCK_RE.sub(replacer, chapter_text)
    # Keep file tidy after block removals.
    updated = re.sub(r"\n{4,}", "\n\n\n", updated)
    return updated


def _write_indicator_files(
    output_dir: Path,
    workbook_var: str,
    include_source: bool,
    section_filter: str,
    workbook_path: Path,
    preserve_stems: set[str] | None = None,
) -> list[str]:
    groups = get_indicator_groups(workbook_path=workbook_path, section_filter=section_filter)
    validation_errors = validate_indicator_groups(groups)
    if validation_errors:
        non_fatal_errors = [
            err for err in validation_errors if err.startswith(_DUPLICATE_ORDER_ERROR_PREFIX)
        ]
        fatal_errors = [
            err for err in validation_errors if not err.startswith(_DUPLICATE_ORDER_ERROR_PREFIX)
        ]

        if non_fatal_errors:
            print(
                "Warning: duplicate order indexes detected in workbook metadata; "
                "continuing object generation."
            )
            for err in non_fatal_errors:
                print(f"  {err}")

        if fatal_errors:
            formatted = "\n".join(f"- {err}" for err in fatal_errors)
            raise ValueError(f"Workbook indicator validation failed:\n{formatted}")

    output_dir.mkdir(parents=True, exist_ok=True)

    model = load_cha_workbook(workbook_path)

    written_stems: list[str] = []
    for group in groups:
        # Look up sheet-level source/note text.  When an indicator group has
        # both a figure and a table, attach source/note only to the table
        # include file to avoid rendering duplicates.
        def _source_note(object_id: str, attach: bool) -> tuple[str, str]:
            if not attach:
                return "", ""
            spec = model.source_specs.get(object_id)
            if not spec:
                return "", ""
            return spec.source_text, spec.note_text

        if group.figure_id:
            figure_path = output_dir / f"{group.figure_id}.qmd"
            fig_include_source = (
                include_source
                and not bool(group.table_id)
                and group.figure_id not in OBJECT_IDS_OMIT_EMBEDDED_SOURCE
            )
            fig_src, fig_note = _source_note(
                group.figure_id, attach=not bool(group.table_id)
            )
            figure_path.write_text(
                render_figure_blocks(
                    group.figure_id,
                    group.caption,
                    workbook_var=workbook_var,
                    include_source=fig_include_source,
                    source_text=fig_src,
                    note_text=fig_note,
                )
                + "\n",
                encoding="utf-8",
            )
            written_stems.append(group.figure_id)

        if group.table_id:
            table_path = output_dir / f"{group.table_id}.qmd"
            tbl_include_source = (
                include_source and group.table_id not in OBJECT_IDS_OMIT_EMBEDDED_SOURCE
            )
            tbl_src, tbl_note = _source_note(group.table_id, attach=True)
            table_path.write_text(
                render_table_blocks(
                    group.table_id,
                    group.caption,
                    workbook_var=workbook_var,
                    include_source=tbl_include_source,
                    source_text=tbl_src,
                    note_text=tbl_note,
                )
                + "\n",
                encoding="utf-8",
            )
            written_stems.append(group.table_id)

    written_set = set(written_stems)
    preserve_set = preserve_stems or set()
    # Keep output directories in sync with current workbook metadata.
    for existing_file in output_dir.glob("*.qmd"):
        if existing_file.name.startswith("._"):
            # AppleDouble sidecars on network volumes can vanish between glob and unlink.
            existing_file.unlink(missing_ok=True)
            continue
        if existing_file.stem not in written_set and existing_file.stem not in preserve_set:
            existing_file.unlink(missing_ok=True)

    # Ensure preserved include stems referenced by chapters always resolve,
    # even when workbook metadata no longer contains those object IDs.
    for stem in sorted(preserve_set - written_set):
        fallback_path = output_dir / f"{stem}.qmd"
        if fallback_path.exists():
            continue
        if stem.startswith("fig-"):
            fallback_path.write_text(
                "\n".join(
                    [
                        f"<!-- indicator: {stem} (fallback) -->",
                        "",
                        "```{python}",
                        "#| echo: false",
                        "#| warning: false",
                        "#| message: false",
                        f"#| label: {stem}",
                        f"#| fig-alt: 'Figure for {stem}'",
                        "from scripts.cha_registry_renderer import render_figure_object",
                        f'render_figure_object(figure_id="{stem}", workbook_path=CHA_WORKBOOK_PATH).show()',
                        "```",
                        "",
                    ]
                ),
                encoding="utf-8",
            )
        elif stem.startswith("tbl-"):
            fallback_path.write_text(
                "\n".join(
                    [
                        f"<!-- indicator: {stem} (fallback) -->",
                        "",
                        "```{python}",
                        "#| echo: false",
                        "#| warning: false",
                        "#| message: false",
                        f"#| label: {stem}",
                        "from scripts.cha_registry_renderer import render_table_object",
                        f'render_table_object(object_id="{stem}", workbook_path=CHA_WORKBOOK_PATH)',
                        "```",
                        "",
                    ]
                ),
                encoding="utf-8",
            )

    return sorted(written_set)


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Generate metadata-driven chapter object include files.")
    parser.add_argument("--workbook", default=str(DEFAULT_WORKBOOK_PATH), help="Workbook path")
    parser.add_argument("--chapter", default=str(DEFAULT_CHAPTER), help="Chapter QMD to validate/rewrite")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="Directory for include files")
    parser.add_argument("--workbook-var", default="CHA_WORKBOOK_PATH", help="Workbook variable used in code blocks")
    parser.add_argument("--section", default="", help="Optional section_tag filter")
    parser.add_argument(
        "--include-source",
        choices=["true", "false"],
        default="true",
        help="Whether generated include files should render source callouts",
    )
    parser.add_argument(
        "--validate-all-chapters",
        action="store_true",
        help="Validate cross-references and include directives in every chapter QMD.",
    )
    parser.add_argument(
        "--rewrite-chapter",
        action="store_true",
        help="Replace repeated python figure/table/source blocks with include directives.",
    )
    parser.add_argument(
        "--strict-refs",
        action="store_true",
        help="Fail if chapter cross-references point to object IDs missing from workbook.",
    )
    args = parser.parse_args(argv)

    workbook_path = Path(args.workbook)
    chapter_path = Path(args.chapter)
    output_dir = Path(args.output_dir)
    include_source = args.include_source == "true"

    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")
    if not chapter_path.exists():
        raise FileNotFoundError(f"Chapter not found: {chapter_path}")

    # Preserve any _generated include stubs referenced by chapter files.
    # Without this, pre-render generation can delete stubs whose object IDs
    # are missing from the workbook registry (e.g., when chapters reference
    # newer IDs than the current workbook contains).
    chapter_text = chapter_path.read_text(encoding="utf-8")
    preserve_stems: set[str] = set()
    chapters_dir = chapter_path.parent
    for chapter_file in _iter_chapter_qmd_files(chapters_dir):
        chapter_body = chapter_file.read_text(encoding="utf-8")
        include_prefix = output_dir.relative_to(chapter_file.parent).as_posix()
        preserve_re = re.compile(re.escape(include_prefix) + r"/([A-Za-z0-9+-]+)\.qmd")
        preserve_stems.update(preserve_re.findall(chapter_body))

    before_hashes = _file_hashes(output_dir) if output_dir.exists() else {}

    written_stems = _write_indicator_files(
        output_dir=output_dir,
        workbook_var=args.workbook_var,
        include_source=include_source,
        section_filter=args.section,
        workbook_path=workbook_path,
        preserve_stems=preserve_stems,
    )

    model = load_cha_workbook(workbook_path)
    fallback_registry_ids: set[str] = set()
    for include_file in output_dir.glob("*.qmd"):
        stem = include_file.stem
        if stem.startswith(("fig-", "tbl-")):
            fallback_registry_ids.add(stem)
    include_stems_all_chapters = _include_stems_from_chapters(chapters_dir)
    registry_ids = set(model.registry.keys()) | fallback_registry_ids | include_stems_all_chapters
    written_stem_set = set(written_stems) | fallback_registry_ids

    if args.validate_all_chapters:
        chapter_errors = _validate_all_chapter_references(
            chapters_dir,
            registry_ids,
            output_dir,
            written_stem_set,
        )
    else:
        chapter_errors = _validate_chapter_references(chapter_text, registry_ids)
        include_errors = _validate_include_directives(
            chapter_text, output_dir, written_stem_set
        )
        chapter_errors.extend(include_errors)

    if chapter_errors and args.strict_refs:
        formatted = "\n".join(f"- {err}" for err in chapter_errors)
        raise ValueError(f"Chapter reference validation failed:\n{formatted}")
    if chapter_errors and not args.strict_refs:
        print("Warning: unresolved cross-references detected:")
        for err in chapter_errors:
            print(f"  {err}")

    if args.rewrite_chapter:
        # Paths inside include directives should be relative to the chapter file.
        include_prefix = output_dir.relative_to(chapter_path.parent).as_posix()
        updated = _replace_python_blocks_with_includes(
            chapter_text,
            set(written_stems),
            include_prefix,
            output_dir,
        )
        chapter_path.write_text(updated, encoding="utf-8")
        print(f"Rewrote chapter with include directives: {chapter_path}")

    after_hashes = _file_hashes(output_dir)
    added, removed, updated = _summarize_generation(before_hashes, after_hashes)
    _print_generation_summary(added, removed, updated, len(written_stems), output_dir)


if __name__ == "__main__":
    main()
