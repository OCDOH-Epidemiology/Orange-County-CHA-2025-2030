# CHA Chapter Word Template – Instructions

Use the **CHA_Chapter_Template.docx** file to draft narrative text for any chapter of the Community Health Assessment. A Python script will convert your Word document into the format used by the Quarto publishing system.

## How To Use the Template

1. Open `templates/CHA_Chapter_Template.docx` in Microsoft Word.
2. **Delete the instructions page** (the first page with the red heading).
3. Replace "Chapter Title Goes Here" with your chapter's actual title.
4. Write your sections using the conventions described below.
5. Fill in the Citations Table at the end of the document.
6. Save your file (e.g. `Demographics_Narrative.docx`).
7. Run the converter:

```
python3 scripts/docx_to_qmd.py Demographics_Narrative.docx --output chapters/03-demographics.qmd
```

## Formatting Conventions

### Headings

Use Word's built-in heading styles from the Styles ribbon:

| Word Style | Converts To | Use For |
|---|---|---|
| Heading 1 | Chapter title (YAML) | Only use ONCE for the chapter title |
| Heading 2 | `##` section heading | Major sections (e.g. "Demographic Characteristics") |
| Heading 3 | `###` subsection heading | Subsections (e.g. "Population", "Age", "Income") |

### Table and Figure References

When your text refers to a data table or figure from the Excel workbook, type the reference in bold brackets:

- **Table:** Type **[Table: population-demographics]** (bold)
- **Figure:** Type **[Figure: labor-force]** (bold)

The ID after the colon must match the `object_id` in the workbook (without the `tbl-` or `fig-` prefix). The converter adds the prefix automatically and inserts the code block to render it.

### Citations (Superscript Numbers)

To cite a source in the body text:

1. Type the citation number (e.g. `7`).
2. Select the number and make it superscript: **Ctrl+Shift+=** (Windows) or **Cmd+Shift+=** (Mac).
3. Add the full citation details to the **Citations Table** at the end of the document.

The converter turns each superscript number into an HTML anchor link and builds the References section automatically.

### Notes (Callout Boxes)

To create a note/callout box, start a paragraph with the word **NOTE:** in bold:

> **NOTE:** The American Community Survey includes a question that intends to capture current sex; there are no questions about gender.

This becomes a styled callout box in the final output.

### Source Callouts

To create a collapsible source citation box, start a paragraph with **SOURCE:** in bold:

> **SOURCE:** US Census Bureau; American Community Survey, 2023 ACS 5-Year Estimates, Table S0101, April 2025

### Bulleted Lists

Use Word's standard bulleted list (the bullet button on the Home ribbon). Each bullet becomes a Markdown list item.

### Bold and Italic

- **Bold text** in Word becomes **bold** in the output.
- *Italic text* in Word becomes *italic* in the output.
- This works for individual words or phrases within a paragraph.

## Citations Table

The last table in the document must be the citations table with these columns:

| Column | Description |
|---|---|
| Number | Citation number matching the superscript in the text |
| Author / Organization | Name of the author or organization |
| Year | Publication year |
| Title | Title of the source |
| URL | Full URL to the source |
| Accessed | When the source was accessed (e.g. "August 2025") |

Add one row per citation. Leave rows empty if not needed. Only rows with a Number value are processed.

## What NOT To Include

- **Data tables** – these come from the Excel workbook and are rendered automatically.
- **Python code** – the converter inserts code blocks based on your `[Table: ...]` and `[Figure: ...]` references.
- **Manual HTML** – the converter generates all HTML anchors and formatting.

## Converter Command Reference

```
python3 scripts/docx_to_qmd.py <input.docx> [options]

Options:
  --output, -o PATH       Output .qmd file (default: same name with .qmd extension)
  --workbook-var NAME     Python variable name for workbook path (default: CHA_WORKBOOK_PATH)
  --object-render-mode    inline | include (default: inline)
  --include-dir PATH      Include directory when mode=include (default: _generated/ch04-objects)
```

## Metadata-Driven Build Workflow (Recommended)

For reusable digital CHA builds, keep indicator metadata in the workbook and generate chapter object blocks automatically.

1. Generate include files + rewrite chapter object blocks:

```
python3 scripts/generate_chapter_objects.py \
  --chapter "chapters/04-Social and Physical Determinants of Health.qmd" \
  --output-dir "chapters/_generated/ch04-objects" \
  --rewrite-chapter
```

2. Optional end-to-end orchestration (includes optional docx conversion and optional render):

```
python3 scripts/build_digital_cha.py \
  --chapter "chapters/04-Social and Physical Determinants of Health.qmd" \
  --rewrite-chapter \
  --render chapter
```

This keeps narrative authoring flexible (Word or direct QMD) while making figure/table/source rendering metadata-driven from the workbook.
