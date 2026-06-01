#!/usr/bin/env python3
"""Update _quarto.yml with new chapter files and render the book.

This script:
1. Scans the chapters/ directory for all .qmd files
2. Compares with what's in _quarto.yml
3. Updates _quarto.yml to include any new chapters (sorted by filename)
4. Renders the book using quarto render
"""
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

# Get the project root (parent of scripts/)
PROJECT_ROOT = Path(__file__).parent.parent
CHAPTERS_DIR = PROJECT_ROOT / "chapters"
QUARTO_YML = PROJECT_ROOT / "_quarto.yml"


def get_chapter_files() -> list[Path]:
    """Get all .qmd files from the chapters directory, sorted."""
    if not CHAPTERS_DIR.exists():
        print(f"Error: {CHAPTERS_DIR} does not exist")
        sys.exit(1)
    
    chapters = sorted(CHAPTERS_DIR.glob("*.qmd"))
    return chapters


def read_quarto_yml() -> str:
    """Read the current _quarto.yml file."""
    if not QUARTO_YML.exists():
        print(f"Error: {QUARTO_YML} does not exist")
        sys.exit(1)
    
    return QUARTO_YML.read_text()


def get_existing_chapters(content: str) -> list[str]:
    """Extract existing chapter paths from _quarto.yml."""
    # Find the chapters section
    chapters_match = re.search(
        r"chapters:\s*\n((?:\s+-.*\n)*)",
        content,
        re.MULTILINE
    )
    
    if not chapters_match:
        return []
    
    chapters_text = chapters_match.group(1)
    # Extract all chapter paths (lines starting with -)
    chapters = re.findall(r"-\s+(.+\.qmd)", chapters_text)
    return chapters


def update_quarto_yml(chapter_files: list[Path], existing_chapters: list[str]) -> bool:
    """Update _quarto.yml with new chapters. Returns True if changes were made."""
    # Convert Path objects to relative paths from project root
    chapter_paths = [f"chapters/{f.name}" for f in chapter_files]
    
    # Also include index.qmd if it exists
    index_path = PROJECT_ROOT / "index.qmd"
    if index_path.exists():
        chapter_paths.insert(0, "index.qmd")
    
    # Check if there are any new chapters (not in existing list)
    new_chapters = [ch for ch in chapter_paths if ch not in existing_chapters]
    
    # Check if there are any missing chapters (in existing list but file doesn't exist)
    existing_file_paths = {Path(PROJECT_ROOT / ch) for ch in existing_chapters if not ch.startswith("index")}
    missing_chapters = [ch for ch in existing_chapters if ch not in chapter_paths and ch != "index.qmd"]
    
    if not new_chapters and not missing_chapters:
        print("No new chapters found. All chapters are already in _quarto.yml")
        return False
    
    if new_chapters:
        print(f"Found {len(new_chapters)} new chapter(s):")
        for ch in new_chapters:
            print(f"  - {ch}")
    
    if missing_chapters:
        print(f"Found {len(missing_chapters)} missing chapter(s) (in _quarto.yml but file not found):")
        for ch in missing_chapters:
            print(f"  - {ch}")
        print("Note: These will be removed from _quarto.yml")
    
    # Read the current content
    content = read_quarto_yml()
    
    # Preserve existing order, but add new chapters at the end (before index.qmd if present)
    # Find the chapters block
    pattern = r"(chapters:\s*\n)((?:\s+-.*\n)*)"
    
    def replace_chapters(match):
        indent = "    "  # 4 spaces for YAML indentation
        
        # Start with existing chapters that still exist
        updated_chapters = []
        for existing_ch in existing_chapters:
            if existing_ch in chapter_paths:
                updated_chapters.append(existing_ch)
        
        # Add any new chapters that weren't in the existing list
        for new_ch in chapter_paths:
            if new_ch not in updated_chapters:
                updated_chapters.append(new_ch)
        
        new_chapters_text = "\n".join([f"{indent}- {ch}" for ch in updated_chapters])
        return match.group(1) + new_chapters_text + "\n"
    
    new_content = re.sub(pattern, replace_chapters, content, flags=re.MULTILINE)
    
    # Write back to file
    QUARTO_YML.write_text(new_content)
    print(f"\nUpdated {QUARTO_YML}")
    return True


def render_book() -> bool:
    """Render the Quarto book. Returns True if successful."""
    print("\nRendering Quarto book...")
    try:
        result = subprocess.run(
            ["quarto", "render"],
            cwd=PROJECT_ROOT,
            check=True,
            capture_output=False,  # Show output in real-time
            text=True
        )
        print("\n✓ Book rendered successfully!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"\n⚠ Warning: Book rendering encountered errors.")
        print("Some files may have rendered successfully. Check the output above for details.")
        print("\nCommon issues:")
        print("  - Files with 'jupyter: python3' but no code blocks should remove that line")
        print("  - Syntax errors in .qmd files")
        return False
    except FileNotFoundError:
        print("Error: 'quarto' command not found. Make sure Quarto is installed and in your PATH.")
        return False


def main():
    """Main function."""
    print("Scanning for chapter files...")
    chapter_files = get_chapter_files()
    
    if not chapter_files:
        print("No .qmd files found in chapters/ directory")
        sys.exit(1)
    
    print(f"Found {len(chapter_files)} chapter file(s) in chapters/")
    
    # Read current _quarto.yml
    content = read_quarto_yml()
    existing_chapters = get_existing_chapters(content)
    
    print(f"Found {len(existing_chapters)} existing chapter(s) in _quarto.yml")
    
    # Update _quarto.yml if needed
    updated = update_quarto_yml(chapter_files, existing_chapters)
    
    # Always render (in case files changed even if _quarto.yml didn't)
    if updated:
        print("\n" + "="*60)
        print("_quarto.yml has been updated. Rendering book...")
        print("="*60)
    
    success = render_book()
    
    if updated:
        print("\n" + "="*60)
        print("Summary:")
        print(f"  ✓ _quarto.yml updated with new chapters")
        if success:
            print("  ✓ Book rendered successfully")
        else:
            print("  ⚠ Book rendering had errors (see above)")
        print("="*60)
    
    # Don't exit with error code if only render failed but update succeeded
    # This allows the user to fix content issues and re-run render manually
    if not updated and not success:
        sys.exit(1)


if __name__ == "__main__":
    main()

