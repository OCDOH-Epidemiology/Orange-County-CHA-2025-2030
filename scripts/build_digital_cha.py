#!/usr/bin/env python3
"""
Build pipeline for metadata-driven digital CHA output.

Pipeline stages:
1) Generate chapter include files from workbook metadata.
2) Optionally convert a DOCX narrative to QMD.
3) Optionally run Quarto render (chapter or full site).
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from scripts.docx_to_qmd import convert_docx_to_qmd  # noqa: E402
from scripts.generate_chapter_objects import (  # noqa: E402
    DEFAULT_CHAPTER,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_WORKBOOK_PATH,
    main as generate_chapter_objects_main,
)


def _run_quarto_render(target: str, chapter_path: Path) -> None:
    if target == "none":
        return
    command = ["quarto", "render"]
    if target == "chapter":
        command.append(str(chapter_path))
    subprocess.run(command, cwd=PROJECT_ROOT, check=True)


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Build metadata-driven digital CHA output.")
    parser.add_argument("--workbook", default=str(DEFAULT_WORKBOOK_PATH), help="Workbook path")
    parser.add_argument("--chapter", default=str(DEFAULT_CHAPTER), help="Chapter QMD path")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="Include output directory")
    parser.add_argument("--workbook-var", default="CHA_WORKBOOK_PATH", help="Workbook variable name in code blocks")
    parser.add_argument("--section", default="", help="Optional section_tag filter for include generation")
    parser.add_argument("--strict-refs", action="store_true", help="Fail on chapter refs missing from workbook")
    parser.add_argument("--rewrite-chapter", action="store_true", help="Rewrite chapter object blocks to include directives")
    parser.add_argument(
        "--include-source",
        choices=["true", "false"],
        default="true",
        help="Whether generated include files should render source callouts",
    )
    parser.add_argument("--docx", default="", help="Optional DOCX narrative input path")
    parser.add_argument("--docx-output", default="", help="Output QMD path for DOCX conversion")
    parser.add_argument(
        "--docx-object-render-mode",
        choices=["inline", "include"],
        default="include",
        help="How DOCX converter emits object references.",
    )
    parser.add_argument(
        "--render",
        choices=["none", "chapter", "site"],
        default="none",
        help="Quarto render target after generation.",
    )
    args = parser.parse_args(argv)

    # 1) Generate include files (and optionally rewrite chapter)
    generator_args = [
        "--workbook",
        args.workbook,
        "--chapter",
        args.chapter,
        "--output-dir",
        args.output_dir,
        "--workbook-var",
        args.workbook_var,
        "--section",
        args.section,
        "--include-source",
        args.include_source,
    ]
    if args.rewrite_chapter:
        generator_args.append("--rewrite-chapter")
    if args.strict_refs:
        generator_args.append("--strict-refs")
    generate_chapter_objects_main(generator_args)

    # 2) Optional DOCX -> QMD conversion
    if args.docx:
        docx_path = Path(args.docx)
        if not docx_path.exists():
            raise FileNotFoundError(f"DOCX file not found: {docx_path}")
        output_path = Path(args.docx_output) if args.docx_output else docx_path.with_suffix(".qmd")
        qmd_text = convert_docx_to_qmd(
            docx_path=docx_path,
            workbook_var=args.workbook_var,
            object_render_mode=args.docx_object_render_mode,
            include_dir=Path(args.output_dir).relative_to(Path(args.chapter).parent).as_posix(),
        )
        output_path.write_text(qmd_text, encoding="utf-8")
        print(f"Converted DOCX to QMD: {docx_path} -> {output_path}")

    # 3) Optional Quarto render
    _run_quarto_render(args.render, Path(args.chapter))
    if args.render != "none":
        print(f"Quarto render complete ({args.render}).")


if __name__ == "__main__":
    main()
