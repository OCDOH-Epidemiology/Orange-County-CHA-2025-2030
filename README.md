# Community Health Assessment - Quarto Workbook Template

A reusable [Quarto](https://quarto.org) **book** template for building a Community Health Assessment (CHA) website + PDF, with a Python-based data/figure pipeline. Clone it, drop in your data, fill in the chapter skeletons, and render.

## What's in here

| Path | Purpose |
|---|---|
| `_quarto.yml` | Book config: title, chapter list, HTML/PDF formats, pre/post-render hooks. |
| `index.qmd` | Landing page (population focus + contacts). |
| `chapters/*.qmd` | Chapter skeletons - section headings + `TODO` placeholders ready to fill. |
| `chapters/_generated/objects/` | Auto-generated table/figure include files (created by the pipeline). |
| `scripts/` | Python/Node build pipeline + authoring guides (see below). |
| `includes/` | Accessibility skip-links and PDF LaTeX packages. |
| `templates/` | Word authoring template + `TEMPLATE_INSTRUCTIONS.md`. |
| `theme.scss`, `custom-citation-styles.css`, `references-dropdown.html` | Styling. |
| `references.bib` | Bibliography (starts empty - add your citations). |
| `data/raw/workbook.xlsx` | The project data workbook (a working starter ships with the template). |
| `chapters/13-example-data-objects.qmd` | Live demo of a figure + table generated from the workbook (delete when done). |
| `source/` | Optional source Word documents. |
| `media/`, `pdfs/` | Images and downloadable PDFs (e.g. maps). |
| `docs/` | Rendered output (generated; git-ignored). |

## Quick start

1. Install Python dependencies:
   ```bash
   python -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```
2. Render the book:
   ```bash
   quarto render
   ```
3. Open `docs/index.html` in a browser. The "Template Example: Live Data Objects" chapter shows a figure + table built from the starter `data/raw/workbook.xlsx`.

## How the data pipeline works

Tables and figures are **data-driven**, not hand-written. The Excel workbook is both the data and the configuration:

```
data/raw/workbook.xlsx  ──generate──>  chapters/_generated/objects/<id>.qmd  ──include──>  rendered figure/table
```

1. The workbook uses the **flat per-indicator format** (the same format as the Mid-Hudson CHA workbook): one sheet per indicator, with config cells in columns A/B (`Name`, `Table/Figure/Both`, `Object ID`, `Figure Type`, `X Column`, ...) and a data block to the right of an `Enter Data` cell. The starter `data/raw/workbook.xlsx` includes two example sheets (`Example Indicator`, `Example Trend`). See `scripts/WORKBOOK_SCHEMA.md`. (A normalized `_registry`/`_figure_specs`/`_table_specs` format is also supported.)
2. `scripts/generate_chapter_objects.py` reads the workbook and writes one include file per object into `chapters/_generated/objects/`. These files **are committed** to the repo so includes resolve on the very first render. The `pre-render` hook in `_quarto.yml` re-runs the generator on every `quarto render` to keep them in sync as the workbook changes.
3. A chapter pulls an object in with `{{< include _generated/objects/<object_id>.qmd >}}`, which reads the workbook at render time via the `CHA_WORKBOOK_PATH` defined in that chapter's setup block.

**Convention:** the workbook lives at `data/raw/workbook.xlsx`. Keep that name (or update `CHA_WORKBOOK_PATH` in your chapters and the `--workbook` path in the `_quarto.yml` pre-render hook). Just dropping an arbitrary spreadsheet in `data/raw/` is **not** enough — it must follow the flat-indicator schema in `scripts/WORKBOOK_SCHEMA.md`.

> **When you add a brand-new indicator**, run the generator once before rendering so its include file exists (Quarto resolves includes before the pre-render hook runs):
> ```bash
> python3 scripts/generate_chapter_objects.py --workbook "data/raw/workbook.xlsx" --chapter "chapters/13-example-data-objects.qmd" --output-dir "chapters/_generated/objects" --include-source true
> ```

## Spinning up a new workbook

1. **Rename the project.** Edit `_quarto.yml`: set `book.title`, `book.subtitle`, and `book.author`. Update `index.qmd` (population focus + contacts).
2. **Replace the data.** Put your real workbook at `data/raw/workbook.xlsx` (same schema as the starter). It feeds every figure/table.
3. **Wire your chapters into the pipeline.** Add a `pre-render` line in `_quarto.yml` for each chapter that uses objects (mirroring the example line), so its include files regenerate on render.
4. **Write narrative.** Each `chapters/*.qmd` keeps the standard section structure with `TODO` placeholders. Replace the placeholders with your text. (You can also author chapters in Word - see `templates/TEMPLATE_INSTRUCTIONS.md`.)
5. **Insert tables/figures.** Each chapter marks where data objects go with `<!-- OBJECT: <id> -->` comments. Add the indicator sheet to the workbook, run the generator once (see note above), then replace the marker with a real include `{{< include _generated/objects/<id>.qmd >}}`.
6. **Add citations.** Append BibTeX entries to `references.bib` and cite them in text with `[@citation-key]`.
7. **Remove the demo.** Delete `chapters/13-example-data-objects.qmd` and its line in `_quarto.yml` once you no longer need the reference example.

## Scripts

See `scripts/` for the full pipeline and guides:
- `scripts/WORKBOOK_SCHEMA.md`, `scripts/AUTOMATION_GUIDE.md`, `scripts/PIPELINE_USAGE.md` - how the data/object pipeline works.
- `scripts/FIGURE_TEMPLATE.md`, `scripts/TABLE_TEMPLATE.md` - figure/table authoring patterns.
- `scripts/generate_chapter_objects.py` - generate table/figure include files from the workbook.
- `scripts/docx_to_qmd.py` - convert a Word narrative draft into a chapter `.qmd`.
- `scripts/patch_quarto_search.js` - post-render search fix-up.

## Publishing to GitHub Pages

The book outputs to `docs/` for GitHub Pages.
1. Run `quarto render`.
2. Commit and push to `main`.
3. In GitHub: Settings -> Pages -> Build from branch `main`, folder `/docs`.

## Accessibility

`ACCESSIBILITY_CHANGES.md` documents the accessibility conventions baked into this template (skip-links, alt text, heading hierarchy, etc.). Keep them in mind as you author.
