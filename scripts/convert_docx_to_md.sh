#!/usr/bin/env bash
set -euo pipefail

SOURCE_PATH=${1:-"Source/2025 Regional CHA Document Orange County 12.17.2025.docx"}
OUT_DIR="draft"
MEDIA_DIR="media"

mkdir -p "$OUT_DIR" "$MEDIA_DIR"

quarto pandoc "$SOURCE_PATH" -t gfm --extract-media="$MEDIA_DIR" -o "$OUT_DIR/draft.md"

echo "Draft written to $OUT_DIR/draft.md"
