# Convert plain Figure/Table citations to Quarto cross-refs

Use this prompt **after a chapter’s prose, figure/table includes, and HTML comments are finished**.

---

## Prompt

```text
Convert remaining plain Figure/Table citations in prose to Quarto cross-references in:

chapters/<CHAPTER>.qmd

### Goal
Replace hard-coded citations like `(Figure 55)`, `(Table 37)`, split-line
`(Figure` + next-line `57).`, and typos like `(Figure F80)` with Quarto
cross-refs such as `[@fig-allsitesca-incidence-aar]` / `[@tbl-…]`.

Quarto will render the numbered “Figure N” / “Table N” link automatically.

### Do
1. Scope: only the named `.qmd` file. Leave HTML comments like
   `<!-- Figure 55: … -->` as documentation.
2. Map each prose citation from nearby HTML comments + includes, e.g.:
     <!-- Figure 55: … -->
     {{< include _generated/objects/fig-allsitesca-incidence-aar.qmd >}}
   → `(Figure 55)` becomes `[@fig-allsitesca-incidence-aar]`
3. Match existing chapter style (e.g. heart disease / CLRD already converted):
     … NYS overall [@fig-allsitesca-incidence-aar].
   not `(Figure 55)` and not `([@fig-…])`.
4. Prefer the figure label when both fig and tbl are included for the same
   indicator; use `[@tbl-…]` when the prose cites counts/tables only.
5. Fix include mismatches needed for links to resolve:
   - missing fig include when prose cites a figure but only the tbl is included
   - wrong/duplicate include paths (use the real file under
     `chapters/_generated/objects/`)
   - misnumbered prose (e.g. two different objects both called “Figure 73”)
     — map by meaning + nearest correct comment/include, not the wrong number
6. If the same hard-coded number is reused for different objects (e.g. two
   “Table 37” comments), map the prose cite to the object the sentence
   actually describes.
7. Verify:
   - no leftover `(Figure`, `(Table`, or `Figure F` in prose (comments OK)
   - every new `[@fig-…]` / `[@tbl-…]` matches an included object’s `#| label:`
   - report any missing generated object files you cannot fix

### Do not
- Edit the plan file or unrelated chapters
- Invent object ids that do not exist under `_generated/objects/`
- Rewrite narrative text beyond swapping citations (and minimal include fixes)
- Remove or renumber HTML figure/table comments unless fixing an obvious typo
  tied to a broken include path

### Example
Before:
  … compared with 466.8 in NYS overall (Figure 55).

After:
  … compared with 466.8 in NYS overall [@fig-allsitesca-incidence-aar].
```

---

## How to use

1. Replace `chapters/<CHAPTER>.qmd` with the finished chapter path.
2. Paste the prompt block into the agent chat.
3. Skim the diff, especially misnumbered cites and any include path fixes.
