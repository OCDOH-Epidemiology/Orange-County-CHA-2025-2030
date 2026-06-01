# Citation Automation Guide

## Overview

You have a complete automation system for adding citations to QMD files! The system includes:

1. **`add_citations.py`** - Main script that adds citations to QMD files
2. **`extract_citations_from_draft.py`** - Helper to extract citations from draft.md
3. **`demographics_citations.json`** - Pre-made citations file for demographics chapter

## Quick Start

### For a New Chapter File

1. **Add citation markers** in your `.qmd` file:
   ```markdown
   The population grew 4.7% from 2010 to 2020.[7]
   Income affects health outcomes.[8]
   ```

2. **Create or use a citations JSON file**:
   ```json
   {
     "7": "United States Census Bureau, 2025, https://...",
     "8": "Robert Wood Johnson Foundation, 2013, https://..."
   }
   ```

3. **Run the script**:
   ```bash
   python scripts/add_citations.py chapters/your-file.qmd --citations scripts/your-citations.json
   ```

4. **Done!** Citations are now clickable links with a references section.

## Three Ways to Get Citations

### Option 1: Manual JSON File (Recommended for Control)

Create a JSON file with your citations:

```json
{
  "7": "United States Census Bureau, 2025, https://data.census.gov/..., accessed August 2025",
  "8": "Robert Wood Johnson Foundation, 2013, https://www.rwjf.org/..., accessed August 2025"
}
```

Then use it:
```bash
python scripts/add_citations.py chapters/your-file.qmd --citations your-citations.json
```

### Option 2: Extract from draft.md

If your citations are in `draft/draft.md` in `[^N]:` format:

```bash
# Extract all citations from draft.md
python scripts/extract_citations_from_draft.py --output scripts/all_citations.json

# Use the extracted citations
python scripts/add_citations.py chapters/your-file.qmd --citations scripts/all_citations.json
```

### Option 3: Extract from DOCX/PDF

If you have a source document with a references section:

```bash
# From DOCX
python scripts/add_citations.py chapters/your-file.qmd --docx "source/document.docx"

# From PDF
python scripts/add_citations.py chapters/your-file.qmd --pdf "source/document.pdf"
```

## Complete Workflow Example

Let's say you want to add citations to a new chapter:

```bash
# Step 1: Extract citations from draft.md (if available)
python scripts/extract_citations_from_draft.py --output scripts/my_citations.json

# Step 2: Edit your QMD file and add citation markers [7], [8], etc.

# Step 3: Run the citation script
python scripts/add_citations.py chapters/04-your-chapter.qmd --citations scripts/my_citations.json

# Step 4: Check the output - citations are now linked!
```

## What the Script Does

1. ✅ Finds all citation markers like `[7]`, `[8]` in your QMD file
2. ✅ Converts them to clickable superscript links: `<a href="#cite-7"><sup>7</sup></a>`
3. ✅ Creates a References section at the end with all citations
4. ✅ Formats URLs as `<https://...>` to match introduction.qmd style
5. ✅ Won't overwrite existing references (safety feature)

## File Structure

```
scripts/
├── add_citations.py              # Main citation script
├── extract_citations_from_draft.py  # Extract from draft.md
├── demographics_citations.json   # Pre-made citations (7-11)
├── citations_template.json       # Template for new citations
├── AUTOMATION_GUIDE.md          # This file
└── QUICK_START_CITATIONS.md     # Quick reference
```

## Pre-made Citation Files

- **`demographics_citations.json`** - Contains citations 7-11 used in demographics chapter
- **`citations_template.json`** - Template to create your own

## Troubleshooting

### "References section already exists"
The script won't overwrite existing references. To regenerate:
1. Remove the `## References` section from your QMD file
2. Run the script again

### "No citations extracted"
- Check your JSON file format (must be valid JSON)
- Make sure citation numbers match what's in your QMD file
- For DOCX/PDF: ensure the document has a "References" section

### "No citation markers found"
- Make sure you have `[7]`, `[8]`, etc. (not `[ 7 ]` with spaces)
- Check for typos in citation markers

### Script not found
Make sure you're running from the project root:
```bash
cd "Mid-HudsonRegionalCHA-2025"
python scripts/add_citations.py ...
```

## Advanced Usage

### Extract specific citations from draft.md

Edit `extract_citations_from_draft.py` to filter specific citation numbers, or manually edit the JSON output.

### Batch process multiple files

Create a simple bash script:
```bash
#!/bin/bash
for file in chapters/*.qmd; do
    python scripts/add_citations.py "$file" --citations scripts/all_citations.json
done
```

## Tips

1. **Keep citation JSON files organized** - Name them by chapter or topic
2. **Use version control** - Commit your citation JSON files
3. **Test on a copy first** - The script modifies files in place
4. **Check the output** - Always review the generated references section

## Need Help?

- See `QUICK_START_CITATIONS.md` for a quick reference
- See `CITATIONS_README.md` for detailed documentation
- Check the script help: `python scripts/add_citations.py --help`

