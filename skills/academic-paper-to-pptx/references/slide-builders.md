# Slide Renderer Reference — `build_deck.py`

The per-slide python-pptx code lives in **`scripts/build_deck.py`** (one `render_*` function
per slide type). You do **not** hand-author these per deck — you author a deck spec and the
builder renders it. This file documents what each renderer draws and the constants it uses, so
you can adjust `build_deck.py` itself (a new slide type, a template variant) with full context.

- **Authoring a deck?** → `references/deck-spec.md` (the JSON schema, one example per type).
- **Editing the renderers?** → this file + the code in `scripts/build_deck.py`.

All renderers open the bundled `template.pptx` (Moffitt Master 0), which provides the top bar,
bottom 4-color bar, badge icon, title formatting, and bullet styling for free.

## Renderer map

| Spec `type` | `render_*` | Layout | Draws | Key spec fields |
|-------------|-----------|--------|-------|-----------------|
| `title` | `render_title` | `Title Slide` | Trial name (red), subtitle, author/affiliation/date block, logo | `title`, `subtitle`, `authors`, `affiliation`, `date`, `title_size` |
| `bullets` | `render_bullets` | `Title and Content` | Bullets (18pt) + sub-bullets (16pt), optional italic key message | `title`, `key_message`, `bullets[].{text,bold,sub}` |
| `table` | `render_table` | `Title Only` | Centered arm-colored table (cells ≥14pt), optional key message + note | `title`, `arms[].{label,color}`, `rows[].{label,vals,bold,hdr}`, `key_message`, `note`, `font`, `row_h`, `top` |
| `figure` | `render_figure` | `Title Only` | Aspect-preserved image + caption, optional key message | `title`, `image`, `key_message`, `caption[]`, `max_h`, `fig_top` |
| `study_design` | `render_study_design` | `Title Only` | Eligibility box → R circle → arm boxes → endpoints box, connectors (text ≥12pt) | `description`, `eligibility[]`, `randomization`, `n`, `arms[].{name,detail,color}`, `primary_endpoints[]`, `secondary_endpoints[]`, `registration`, `footer` |
| `consort` | `render_consort` | `Title Only` | Assessed → (excluded) → randomized → per-arm allocated/not-received/received/outcomes, connectors | `assessed`, `excluded[]`, `randomized`, `arms[].{color,allocated[],not_received[],received,outcomes[]}` |

Every content-slide renderer calls `add_badge_and_citation(ctx, slide)` (study badge top-right,
citation bottom-right). The title slide has neither.

## Font floors (baked into the renderers)

`build_deck.py` sizes every run through `pt(size, floor)`, clamping up to the floor. The spec
has no font field, so nothing can render below floor:

| Role | Floor | Constant |
|------|-------|----------|
| Body bullet (level 0) | 18 | `F_BODY` |
| Sub-bullet (level 1) | 16 | `F_SUB` |
| Table cell | 14 | `F_TABLE` |
| Diagram text (study-design / CONSORT) | 12 | `F_DIAG` |
| Key message / subtitle | 14 | `F_KEYMSG` |
| Citation | 10 | `F_CITE` (reference element) |
| Badge | 11 | `F_BADGE` (reference element) |

## Color palette (`PALETTE` in `build_deck.py`; spec uses the names)

| Name | RGB | Hex | Usage |
|------|-----|-----|-------|
| `titleblue` | (0,51,102) | `003366` | Badge text, subtitles, header cells |
| `red` | (255,0,0) | `FF0000` | Title-slide trial name |
| `body` | (51,51,51) | `333333` | Body text, shape text |
| `gray` / `control` | (128,128,128) | `808080` | Control-arm header, control boxes |
| `white` | (255,255,255) | `FFFFFF` | Text on colored fills |
| `teal` | (46,125,125) | `2E7D7D` | Experimental arm 1 |
| `blue` | (68,114,196) | `4472C4` | Experimental arm 2 |
| `navy` | (27,42,74) | `1B2A4A` | Highlight box |
| `lightbg` | (240,244,248) | `F0F4F8` | Eligibility/endpoint/section fills |
| `border` | (204,204,204) | `CCCCCC` | Box/table borders |
| — | (153,153,153) | `999999` | Citation text (`GRAY`) |

## Key dimensions (EMU)

Slide is 13.33" × 7.50" = 12192000 × 6858000 EMU.

| Element | EMU | Inches |
|---------|-----|--------|
| Content left margin | 838200 | 0.92" |
| Title placeholder top | 49440 | 0.05" |
| Badge box | left 9984826, top 157655 | ~10.92", 0.17" |
| Citation box | left 9987101, top 6463328 | ~10.92", 7.07" |
| Bottom bar top | 6721475 | 7.35" |
| Slide width / height | 12192000 / 6858000 | 13.33" / 7.50" |

Tables auto-center: `table_x = (12192000 - total_table_width) // 2`.

## Adding a new slide type

1. Write a `render_<type>(ctx, s, n)` in `build_deck.py` (use `pt(size, floor)` for every
   font; call `add_badge_and_citation(ctx, slide)` on content slides).
2. Register it in the `RENDERERS` dict.
3. Add a unit test in `scripts/tests/test_build_deck.py` and document the fields in
   `references/deck-spec.md`.
