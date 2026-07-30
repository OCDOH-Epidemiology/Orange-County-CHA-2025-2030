# Digital CHA Pipeline Usage

This project supports a metadata-driven workflow where workbook metadata is the source of truth for indicator objects (figure/table/source).

## Default workflow (automatic object generation on render)

`_quarto.yml` is configured with a `project.pre-render` hook that runs object generation before each render:

```bash
quarto render
```

Pre-render command:

```bash
./generate-objects
```

This means you no longer need to run `generate_chapter_objects.py` manually for normal builds.

## 1) Generate include files from workbook metadata (manual/advanced)

```bash
python3 scripts/generate_chapter_objects.py \
  --workbook "data/raw/Mid-Hudson Regional Community Health Assessment 2025 Data File.xlsx" \
  --chapter "chapters/04-OCDOH.qmd" \
  --output-dir "chapters/_generated/ch04-objects"
```

Output:

- One `.qmd` include file per indicator in `chapters/_generated/ch04-objects/`
- Each file contains canonical rendering blocks (figure/table/source)

## 2) Rewrite chapter blocks to include directives (optional)

```bash
python3 scripts/generate_chapter_objects.py \
  --chapter "chapters/04-OCDOH.qmd" \
  --output-dir "chapters/_generated/ch04-objects" \
  --rewrite-chapter
```

This replaces repetitive object python chunks with include shortcodes.

## 3) Validate references strictly (optional)

```bash
python3 scripts/generate_chapter_objects.py \
  --chapter "chapters/04-OCDOH.qmd" \
  --strict-refs
```

Validation checks include:

- Missing workbook object IDs referenced in chapter cross-references
- Duplicate order indexes within a section
- Missing captions in indicator metadata
- Broken figure/table pairing when chapter references both fig/tbl for a slug

## 4) Word-first narrative conversion (optional)

Inline python object blocks:

```bash
python3 scripts/docx_to_qmd.py "my_narrative.docx" --output "chapters/03-demographics.qmd"
```

Include-shortcode object blocks:

```bash
python3 scripts/docx_to_qmd.py \
  "my_narrative.docx" \
  --output "chapters/03-demographics.qmd" \
  --object-render-mode include \
  --include-dir "_generated/ch04-objects"
```

## 5) One-command orchestration

Use the pipeline wrapper script for repeatable builds:

```bash
python3 scripts/build_digital_cha.py \
  --chapter "chapters/04-OCDOH.qmd" \
  --output-dir "chapters/_generated/ch04-objects" \
  --rewrite-chapter \
  --render chapter
```

`--render` options:

- `none` (default)
- `chapter`
- `site`

## New CHA Project Checklist

1. Update workbook metadata/data sheets for the new CHA.
2. Generate include files from metadata.
3. Rewrite chapter object blocks to include directives (or author with include mode from DOCX).
4. Run strict reference validation and fix any unresolved IDs.
5. Render chapter/site and review output.

## 6) PDF narrative extraction (native text, not OCR)

For a published CHA PDF with a text layer, extract prose + figure/table include stubs:

```bash
python3 scripts/extract_pdf_to_qmd.py \
  --pdf "/path/to/Orange County CHA.pdf" \
  --from-page 196 --to-page 228 \
  --out drafts/pdf-extract/ch08-remainder.qmd \
  --report drafts/pdf-extract/ch08-remainder.report.txt
```

**Page numbers:** `--from-page` / `--to-page` are **printed footer** pages by default.
This CHA’s offset is **PDF page = printed + 6** (auto-calibrated; override with
`--page-offset 6`). Use `--pdf-pages` to pass raw 1-based PDF indices instead.

**What it does:**
- Keeps narrative wording; only reflows soft line wraps
- Skips chart axes / table grids (workbook already renders those)
- Maps `Figure N` / `Table N` via workbook Master `F or T #` → Object ID
- Emits `fig-` then `tbl-` includes for figures; `tbl-` only for tables

**Workflow:** review the draft → paste/append into the chapter → run
`prompts/convert-figure-table-crossrefs.md` to turn `(Figure N)` into `[@fig-…]`.

Optional append after review:

```bash
python3 scripts/extract_pdf_to_qmd.py ... --append chapters/08-leading-health-issues.qmd
```

