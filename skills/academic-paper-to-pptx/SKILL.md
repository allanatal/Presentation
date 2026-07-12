---
name: academic-paper-to-pptx
description: "Transform an academic research paper (PDF, DOCX, or Markdown) into a professional PowerPoint presentation styled for oncology or clinical audiences. Use this skill whenever the user uploads a scientific paper, clinical trial publication, or medical research article and asks for a presentation, slide deck, or summary slides. Also trigger when the user says 'make slides from this paper', 'turn this into a presentation', 'create a deck from this study', or references converting any academic manuscript into slides. Handles both quantitative (clinical trials, RCTs) and qualitative research papers. The output is a polished .pptx file with study design diagrams, recreated tables, extracted figures, and structured discussion slides — all in a consistent Moffitt Cancer Center academic style."
---

# Academic Paper to PowerPoint Skill

Transform a research paper into a professional, oncology-grade PowerPoint presentation using the **Moffitt Cancer Center slide template** with accurate data reproduction, extracted figures, and programmatically built study design diagrams.

## Architecture

This skill uses a **template-based approach**: a bundled `template.pptx` contains the Moffitt slide master (Master 0) with all branding baked in — top bar, bottom 4-color bar, Moffitt badge icon, title formatting with red underline, and bullet styling. Python-pptx opens the template and adds slides using the master's layouts. This gives pixel-perfect style matching with minimal code.

**Key library**: `python-pptx` (not PptxGenJS).

**Build model (deck-spec → fixed builder).** You do **not** hand-author python-pptx per deck.
You emit a compact typed **deck spec** (JSON, one entry per slide) that references claim IDs;
the committed `scripts/build_deck.py` consumes spec + claims and renders the deck. This cuts
the biggest per-deck token bucket and gives two integrity properties for free:
- **No sub-floor fonts** — the spec has no font field; sizes live only in `build_deck.py`, all
  ≥ the style-spec floors.
- **No hand-typed numbers in code** — prose uses `{{claim-id}}` tokens substituted from the
  claims file; tables inline values but declare claim IDs, and `qa_crosscheck.py` verifies
  they landed. The builder auto-fills each claim's `slide` field (you never maintain it).

Schema + one example per slide type: `references/deck-spec.md`.

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

Persist every number destined for slides to `<deckname>_claims.json` in the working directory, following the schema in `references/qa-checklist.md`. One claim per verifiable statement (HR+CI+p, median, ORR, each Table 1 row, AE %s, arm sizes, NCT). Values verbatim from the paper. Leave `slide` as `null` — Phase 3's builder derives it and writes a `<deckname>_claims_resolved.json` (you don't maintain `slide` by hand). Keep the markdown text (`paper_text.md`) — the QA step uses it for the orphan-number check.

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

You author a **deck spec** (JSON) and run the committed builder. Full schema and one worked
example per slide type: **`references/deck-spec.md`** — read it before authoring.

### Step 3a: Author `<deckname>_deck_spec.json`

One entry per slide in `slides`, each a typed object (`title` / `bullets` / `table` /
`figure` / `study_design` / `consort`). A `deck` block holds `study` (badge), `citation`,
and title-slide metadata. Put the paper's numbers on slides two ways:
- **`{{claim-id}}` tokens** in prose (key messages, bullets) → substituted from the claims file.
- **slide-level `"claims": [...]`** on tables / diagrams / figure captions whose numbers are
  inlined — `qa_crosscheck.py` verifies each declared claim's numbers actually landed.

**Contract:** every claim in the claims file must be referenced exactly once; `build_deck.py`
errors otherwise. This *is* the "every claim assigned before QA" rule, enforced mechanically.

Recommended slide order (map each to a spec `type`):

| # | Slide | Spec `type` | Notes |
|---|-------|-------------|-------|
| 1 | Title | `title` | Inherits `deck.{title,subtitle,authors,affiliation,date}` |
| 2 | Research Objective | `bullets` | Research gap + core question |
| 3–4 | Literature Background | `bullets` | Max 2 slides |
| 5 | Study Design Diagram | `study_design` | Eligibility → R → arms → endpoints |
| 6 | Statistical Methods | `bullets` | Quantitative only |
| 7 | CONSORT Diagram | `consort` | If patient-flow data available |
| 8–9 | Table 1 (Demographics) | `table` | Exact numbers; split if it would overflow |
| 10–12 | Primary/Secondary Figures | `figure` | Embed extracted images (Phase 2) |
| 13 | Safety / AE Table | `table` | Use `note` for footnote lines |
| 14–16 | Additional Findings | `bullets`/`table`/`figure` | |
| 17–18 | Discussion / Conclusions | `bullets` | Max 2 slides |

**Data reproduction:** every number must match the publication exactly. Tables are
auto-centered; if one won't fit, split across two `table` slides — never shrink below the
floor. Figures embed with preserved aspect ratio (source/extracted figures or programmatic
charts only — never a full-slide screenshot).

### Step 3b: Run the builder

```bash
python "<skill-root>/scripts/build_deck.py" <deckname>_deck_spec.json \
    --claims <deckname>_claims.json --out <deckname>.pptx \
    --resolved-claims <deckname>_claims_resolved.json
```

It copies the Moffitt template, renders every slide, substitutes claim tokens, auto-fills each
claim's `slide` in `<deckname>_claims_resolved.json`, and prints `N/M claims placed`. If it
reports unreferenced claims, add their references to the spec and re-run.

### Font Size Requirements (enforced by the builder)

Hard minimums — the spec has **no font field**, so these are baked into `build_deck.py` and
cannot be violated from the spec. If content doesn't fit, **split across slides**.

| Element | Minimum |
|---------|---------|
| Body text (bullets level 0) | 18pt |
| Sub-bullets (level 1) | 16pt |
| Table cell text (headers + data) | 14pt |
| Study-design / CONSORT diagram text | 12pt |
| Key message / italic subtitle | 14pt |
| Citation text | 10pt (exempt reference element) |
| Study badge | 11pt (exempt reference element) |

**Why:** presentations are projected in large rooms; text below 14pt is illegible at distance.
Citation and badge are reference elements not meant to be read during the talk.

`references/slide-builders.md` documents what each renderer draws (colors, positions, layouts)
if you need to adjust `build_deck.py` itself.

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
python "<skill-root>/scripts/qa_crosscheck.py" <deckname>.pptx \
    --claims <deckname>_claims_resolved.json --mode paper --source-text paper_text.md
```

Use the **resolved** claims file (`build_deck.py` filled in each claim's `slide`). The script
exits non-zero on any ❌ (claim number missing from its slide, orphan number not in the source,
unresolved `[CHECK]`, font below floors, full-slide image). **Fix every ❌ and re-run until
clean.** The deck is delivered **together with** its `<deckname>_QA.md`; the remaining MANUAL
lines are the user's eyeball pass against the source.

---

## Common Pitfalls

- **Inventing data**: Never put a number in the spec that isn't in the paper — every claim value comes from the source; use `[CHECK]` placeholders for anything uncertain.
- **Unreferenced claims**: `build_deck.py` errors if any claim is never placed. Reference each via a `{{claim-id}}` token or a slide-level `"claims"` list — don't silence it with `--allow-unreferenced`.
- **Overcrowded slides**: Split dense content across multiple `bullets`/`table` slides rather than cramming — the font floors are enforced and cannot be shrunk to fit.
- **Wrong figure image path**: `image` is resolved relative to `--images-base` (default: the spec's directory). The builder warns if the file is missing.
- **Study design / CONSORT complexity**: The renderers cover 2–3 arms with the standard topology. For an exotic flow (crossover, 4+ arms, non-standard CONSORT), render it as a `figure` from a pre-made image instead.

## Reference Files

| File | Purpose |
|------|-------------|
| `references/deck-spec.md` | **Deck-spec JSON schema** + one example per slide type (author from this) |
| `references/template.pptx` | Moffitt slide master template (open with python-pptx) |
| `references/slide-builders.md` | What each `build_deck.py` renderer draws (colors, positions, layouts) |
| `references/style-spec.md` | Quick reference for colors, positions, and manual elements |
| `references/qa-checklist.md` | QA workflow, claims-file schema, checklist rules |
| `references/Picture_3.x-wmf` | Moffitt Cancer Center logo (for title slide) |
| `scripts/build_deck.py` | **Committed builder**: deck spec + claims → editable Moffitt PPTX |
| `scripts/extract_figures.py` | Extract images from PDF papers |
| `scripts/qa_crosscheck.py` | Deterministic QA cross-check (shared with pptx-to-pptx) |
