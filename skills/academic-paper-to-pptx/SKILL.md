---
name: academic-paper-to-pptx
description: "Transform an academic research paper (PDF, DOCX, or Markdown) into a professional PowerPoint presentation styled for oncology or clinical audiences. Use this skill whenever the user uploads a scientific paper, clinical trial publication, or medical research article and asks for a presentation, slide deck, or summary slides. Also trigger when the user says 'make slides from this paper', 'turn this into a presentation', 'create a deck from this study', or references converting any academic manuscript into slides. Handles both quantitative (clinical trials, RCTs) and qualitative research papers. The output is a polished .pptx file with study design diagrams, recreated tables, extracted figures, and structured discussion slides — all in a consistent Moffitt Cancer Center academic style."
---

# Academic Paper to PowerPoint Skill

Transform a research paper into a professional, oncology-grade PowerPoint presentation using the **Moffitt Cancer Center slide template** with accurate data reproduction, extracted figures, and programmatically built study design diagrams.

## Architecture

This skill uses a **template-based approach**: a bundled `template.pptx` contains the Moffitt slide master (Master 0) with all branding baked in — top bar, bottom 4-color bar, Moffitt badge icon, title formatting with red underline, and bullet styling. Python-pptx opens the template and adds slides using the master's layouts. This gives pixel-perfect style matching with minimal code.

**Key library**: `python-pptx` (not PptxGenJS). All slides are built in Python.

## Environments & Paths

This skill runs in two environments. All paths below are relative to the **skill root** (the directory containing this SKILL.md):

| Environment | Skill root |
|---|---|
| Claude Code (local machine) | `~/.claude/skills/academic-paper-to-pptx/` |
| Claude desktop app (sandbox) | `/mnt/skills/user/academic-paper-to-pptx/` |

Write outputs (deck, figures, claims/QA files) to the user's working directory — never a temp path — and print the output paths at the end. In the desktop-app sandbox, `/home/claude/` is the working directory.

## Workflow — Four Phases

1. **Ingest** — Convert the paper to markdown, then deeply analyze it
2. **Extract** — Pull embedded figures (KM curves, forest plots, etc.) from the PDF
3. **Build** — Generate the .pptx using python-pptx with the Moffitt template
4. **QA** — Convert to images and visually inspect for defects

## Prerequisites

```bash
pip install markitdown python-pptx pdfplumber Pillow --break-system-packages
```

Also required on PATH: `pdftoppm`/`pdfimages` (poppler) and LibreOffice `soffice` for QA rendering (macOS: `brew install poppler && brew install --cask libreoffice`).

Desktop app only: the generic pptx skill at `/mnt/skills/public/pptx/SKILL.md` provides extra QA/image utilities. Locally, use `soffice` + `pdftoppm` directly (see Phase 4).

---

## Phase 1: Ingest & Analyze the Paper

### Step 1a: Convert to Markdown (Token Efficiency)

Converting to markdown first saves 5–8x tokens vs rasterizing every page.

```bash
# For PDF
markitdown paper.pdf > paper_text.md

# For DOCX
markitdown paper.docx > paper_text.md

# For MD — already done
cp paper.md paper_text.md
```

**Fallback** if markitdown produces garbled output: `pdftotext -layout paper.pdf paper_text.txt`

### Step 1b: Identify Pages with Figures

Figures cannot be captured as text. Scan the markdown for figure references and rasterize those pages:

```bash
pdfimages -list paper.pdf           # list embedded images
pdftoppm -jpeg -r 200 -f <PAGE> -l <PAGE> paper.pdf fig_page  # rasterize specific page
```

### Step 1c: Deep Paper Analysis

Read the full markdown and extract these elements. **Never invent data** — every number must come from the paper.

| Element | What to capture |
|---------|----------------|
| **Title** | Full paper title, authors, journal, year |
| **Research gap** | What problem does this paper address? |
| **Research question** | Core question or objective |
| **Literature context** | Key prior studies, current standard of care |
| **Hypotheses** | Explicit hypotheses (quantitative papers only) |
| **Study design** | Trial type, phase, NCT number, randomization ratio |
| **Population** | Eligibility criteria, sample size, stratification |
| **Treatment arms** | Each arm with drug names, doses, schedules |
| **Endpoints** | Primary and secondary endpoints |
| **Table 1 data** | Demographics — extract EVERY number exactly |
| **Primary results** | HR, CI, p-values, medians |
| **Secondary results** | ORR, DOR, other outcomes with exact numbers |
| **Safety/AEs** | Treatment-related AEs with percentages by arm |
| **Key figures** | Which figures exist and what they show |
| **Discussion** | Main implications, limitations, future directions |

**Data fidelity**: Cross-check extracted values against rasterized page images when uncertain. Never round, estimate, or infer. Use `[CHECK]` placeholders for anything uncertain and flag to the user.

### Step 1d: Write the Claims File

Persist every number destined for slides to `<deckname>_claims.json` in the working directory, following the schema in `references/qa-checklist.md`. One claim per verifiable statement (HR+CI+p, median, ORR, each Table 1 row, AE %s, arm sizes, NCT). Values verbatim from the paper; `slide` stays `null` until Phase 3. Keep the markdown text (`paper_text.md`) — the QA step uses it for the orphan-number check.

---

## Phase 2: Extract Figures from PDF

```bash
python "<skill-root>/scripts/extract_figures.py" paper.pdf figures/
```

Filters out small images (<400px or <50KB). For vector figures that `pdfimages` misses:

```bash
pdftoppm -png -r 300 -f <PAGE> -l <PAGE> paper.pdf figures/vector_fig
```

Crop figure regions with Pillow if needed.

---

## Phase 3: Build the Presentation

### Template Setup

Copy the bundled template and open it with python-pptx:

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import shutil, os

# Copy template to working directory (resolve <skill-root> per the Environments table)
SKILL_ROOT = os.path.expanduser("~/.claude/skills/academic-paper-to-pptx")  # or /mnt/skills/user/... in the desktop app
shutil.copy(os.path.join(SKILL_ROOT, "references/template.pptx"), "output.pptx")
prs = Presentation("output.pptx")

# Get Master 0 (Moffitt) layouts
master = prs.slide_masters[0]
LAYOUTS = {layout.name: layout for layout in master.slide_layouts}
```

Available layouts (all inherit Moffitt branding automatically):
- **`Title Slide`** — Center title + subtitle placeholders
- **`Title and Content`** — Title + bullet content (most slides)
- **`Title Only`** — Title placeholder only (study design, figures)
- **`Two Content`** — Title + two side-by-side content areas (split tables)
- **`Blank`** — No placeholders (fully custom slides)

### What the template gives you for free (DO NOT recreate):
- Top navy bar
- Bottom 4-color bar (navy, teal, light blue, green)
- Moffitt badge icon (top-right)
- Title formatting (dark blue, bold, with red underline)
- Bullet formatting with indent levels

### What you add manually per slide:
- Study name text badge (top-right, beside the icon)
- Citation text (bottom-right)
- Images, tables, shapes for content

Read `references/slide-builders.md` for complete python-pptx code patterns.

### Presentation Structure

| # | Slide | Layout | Notes |
|---|-------|--------|-------|
| 1 | Title | `Title Slide` | Trial name, author, affiliation, date |
| 2 | Research Objective | `Title and Content` | Research gap + core question |
| 3–4 | Literature Background | `Title and Content` | Max 2 slides |
| 5–6 | Hypotheses | `Title and Content` | Quantitative only, max 2 slides |
| 7 | Study Design Diagram | `Title Only` | **Build with python-pptx shapes** |
| 8 | Statistical Methods | `Title and Content` | Quantitative only |
| 9 | CONSORT Diagram | `Title Only` | If patient flow data available |
| 10 | Table 1 (Demographics) | `Two Content` or `Title Only` | Exact numbers from paper |
| 11–12 | Primary Endpoint Figures | `Title Only` | **Embed extracted images** |
| 13 | Safety / AE Table | `Title Only` | Recreate table or embed chart |
| 14–16 | Additional Findings | varies | Max 3 slides |
| 17–18 | Discussion | `Title and Content` | Max 2 slides |

### Study Design Diagram — Always Programmatic

Build with python-pptx shapes on a `Title Only` slide. Components:
- Left rounded rect: eligibility criteria
- Center circle: randomization ratio + N
- Right: treatment arm boxes (colored by arm)
- Far right: endpoints box
- Arrows connecting flow
- Bottom: treatment duration, NCT number

See `references/slide-builders.md` § "Study Design Diagram" for full code.

### Tables — Exact Data Reproduction

Use `slide.shapes.add_table()` for demographics and safety tables. Requirements:
- Column headers colored to match treatment arms
- **Every number must match the publication exactly**
- Split wide tables across two sub-tables on one slide
- **Tables must be horizontally centered** on the slide. Calculate: `table_x = (slide_width - total_table_width) // 2` where slide width = 12,192,000 EMU
- If a table has too many rows to fit at 16pt, split across two slides rather than shrinking below 14pt

### Font Size Requirements (MANDATORY)

These are hard minimums — never go below them regardless of content density. If content doesn't fit, **split across multiple slides** rather than shrinking fonts.

| Element | Preferred | Minimum | Notes |
|---------|-----------|---------|-------|
| **"Title and Content" body text (level 0)** | 18pt | 18pt | Main bullet points |
| **"Title and Content" sub-bullets (level 1)** | 16pt | 16pt | Indented sub-points |
| **Table cell text (headers)** | 16pt | 14pt | Column headers, row labels |
| **Table cell text (data)** | 16pt | 14pt | All data values |
| **Study design diagram text** | 14pt | 12pt | Eligibility, arms, endpoints, labels |
| **Key message / subtitle** | 14pt | 14pt | Italic text below slide title |
| **Conclusion key finding** | 18pt | 18pt | First bold bullet on conclusion |
| **Conclusion sub-bullets** | 14pt | 14pt | Supporting points |
| **Citation text** | 10pt | 10pt | Bottom-right reference (exception to minimum) |
| **Study badge** | 11pt | 11pt | Top-right study name (exception to minimum) |

**Why these sizes matter**: Presentations are projected in large rooms. Text below 14pt becomes illegible at typical viewing distances. The only exceptions are citation and badge text, which are reference elements not meant to be read during the talk.

### Figures — Embed Extracted Images

Use `slide.shapes.add_picture()` with extracted images from Phase 2. Always calculate aspect ratio to avoid distortion.

### Claims Bookkeeping

As each slide is built, fill in the `slide` number of every claim it carries in `<deckname>_claims.json`. Every claim must be assigned before Phase 4.

---

## Phase 4: QA

```bash
# Local: soffice on PATH (macOS fallback: /Applications/LibreOffice.app/Contents/MacOS/soffice)
# Desktop app: python /mnt/skills/public/pptx/scripts/office/soffice.py instead of soffice
soffice --headless --convert-to pdf output.pptx
rm -f slide-*.jpg
pdftoppm -jpeg -r 150 output.pdf slide
```

**Visual check for:**
1. Data accuracy — key numbers match the paper
2. Text overflow in tables and bullets
3. Image placement and aspect ratio
4. Study design diagram renders correctly
5. Branding elements present on every slide (inherited from template)
6. Citations on content slides

### Automated cross-check (mandatory)

Read `references/qa-checklist.md`, then run:

```bash
python "<skill-root>/scripts/qa_crosscheck.py" output.pptx \
    --claims <deckname>_claims.json --mode paper --source-text paper_text.md
```

The script exits non-zero on any ❌ (claim number missing from its slide, orphan number not in the source, unresolved `[CHECK]`, font below floors, full-slide image). **Fix every ❌ and re-run until clean.** The deck is delivered **together with** its `<deckname>_QA.md`; the remaining MANUAL lines are the user's eyeball pass against the source.

---

## Common Pitfalls

- **Inventing data**: Never fill in numbers not in the paper — use `[CHECK]` placeholders
- **Overcrowded slides**: Split dense content across multiple slides
- **Table column widths**: Test with real data — narrow columns cause wrapping
- **Figure aspect ratios**: Always calculate from original image dimensions
- **Study design complexity**: Adapt diagram layout for 2-arm, 3-arm, 4+ arm, crossover
- **Title slide styling**: The `Title Slide` layout uses default placeholder formatting — override font color/size for the trial name (see slide-builders.md)

## Reference Files

| File | Purpose |
|------|-------------|
| `references/template.pptx` | Moffitt slide master template (open with python-pptx) |
| `references/slide-builders.md` | Python-pptx code patterns for every slide type |
| `references/style-spec.md` | Quick reference for colors, positions, and manual elements |
| `references/qa-checklist.md` | QA workflow, claims-file schema, checklist rules |
| `references/Picture_3.x-wmf` | Moffitt Cancer Center logo (for title slide) |
| `scripts/extract_figures.py` | Extract images from PDF papers |
| `scripts/qa_crosscheck.py` | Deterministic QA cross-check (shared with pptx-to-pptx) |
