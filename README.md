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
| `data/raw/` | Drop your source Excel workbook here. |
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
3. Open `docs/index.html` in a browser.

## Spinning up a new workbook

1. **Rename the project.** Edit `_quarto.yml`: set `book.title`, `book.subtitle`, and `book.author`. Update `index.qmd` (population focus + contacts).
2. **Add your data.** Put your Excel workbook in `data/raw/` and update the `CHA_WORKBOOK_PATH` placeholder at the top of each chapter that loads data.
3. **Write narrative.** Each `chapters/*.qmd` keeps the standard section structure with `TODO` placeholders. Replace the placeholders with your text. (You can also author chapters in Word - see `templates/TEMPLATE_INSTRUCTIONS.md`.)
4. **Insert tables/figures.** Each chapter marks data objects with `<!-- OBJECT: <id> ... -->` comments and includes one worked example. Use the pipeline to generate the real include files, then reference them with `{{< include _generated/objects/<id>.qmd >}}`.
5. **Add citations.** Append BibTeX entries to `references.bib` and cite them in text with `[@citation-key]`.
6. **Re-enable pipeline hooks** (optional). The `pre-render` block in `_quarto.yml` is commented out; enable it once your chapters use metadata-driven objects.

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
