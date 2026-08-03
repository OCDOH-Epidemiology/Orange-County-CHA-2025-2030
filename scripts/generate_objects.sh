#!/usr/bin/env bash
# Regenerate figure/table include files from data/raw/workbook.xlsx.
#
# Run this after adding or updating indicators in the workbook, before the first
# render that references new object IDs. Also re-run whenever you change source,
# note, caption, or chart/table config in the workbook.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

exec python3 "${REPO_ROOT}/scripts/generate_chapter_objects.py" \
  --workbook "${REPO_ROOT}/data/raw/workbook.xlsx" \
  --chapter "${REPO_ROOT}/chapters/13-Triangulate the Data.qmd" \
  --output-dir "${REPO_ROOT}/chapters/_generated/objects" \
  --include-source true \
  --validate-all-chapters \
  "$@"
