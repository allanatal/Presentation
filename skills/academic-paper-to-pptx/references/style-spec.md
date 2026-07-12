# Moffitt Style Quick Reference

This is a quick reference for the elements you add **manually** per slide. Everything else (top bar, bottom color bar, badge icon, title formatting, bullet styling) comes from the template automatically.

## What the Template Gives You (DO NOT recreate)

- Top navy decorative bar (`Rectangle 9` on Master 0)
- Bottom 4-color bar as `Picture 9` on Master 0 (navy, teal, light blue, green)
- Moffitt badge icon (circular, top-right area on content layouts)
- Title placeholder formatting: dark blue, bold, with red underline
- Content placeholder: bullet and indent-level formatting
- Slide dimensions: 13.33" × 7.50" (widescreen, standard PowerPoint)

## What You Add Manually Per Slide

### Study Name Badge Text (content slides only)

Position and size extracted from model.pptx XML:

```python
# Top-right, beside the inherited Moffitt icon
x = Emu(9984826)   # ~10.92"
y = Emu(157655)     # ~0.17"
w = Emu(1671098)    # ~1.83"
h = Emu(369332)     # ~0.40"
font_size = Pt(11)
color = RGBColor(0x00, 0x33, 0x66)  # dark blue
alignment = PP_ALIGN.RIGHT
```

### Citation Text (content slides only)

```python
# Bottom-right, above the bottom color bar
x = Emu(9987101)    # ~10.92"
y = Emu(6463328)    # ~7.07"
w = Emu(2204899)    # ~2.41"
h = Emu(307777)     # ~0.34"
font_size = Pt(10)
color = RGBColor(0x99, 0x99, 0x99)  # gray
alignment = PP_ALIGN.RIGHT
```

### Title Slide Affiliations + Date

```python
# Centered below the subtitle placeholder
x = Emu(1154187)    # ~1.26"
y = Emu(4928445)    # ~5.39"
w = Emu(9340438)    # ~10.21"
h = Emu(1200329)    # ~1.31"
affiliation_font = Pt(14), color 666666
date_font = Pt(12), color 666666
alignment = PP_ALIGN.CENTER
```

### Moffitt Logo on Title Slide

```python
# Centered above the trial name
x = Emu(4790362)    # ~5.24"
y = Emu(798667)     # ~0.87"
w = Emu(2266391)    # ~2.48"
h = Emu(578214)     # ~0.63"
# File: references/Picture_3.x-wmf
```

## Color Palette (for manual elements)

| Name | RGB | Hex | Usage |
|------|-----|-----|-------|
| Title Blue | (0,51,102) | `003366` | Study badge text, subtitles |
| Red | (255,0,0) | `FF0000` | Title slide trial name |
| Body | (51,51,51) | `333333` | Body text, shape text |
| Gray | (153,153,153) | `999999` | Citation text |
| White | (255,255,255) | `FFFFFF` | Text on colored backgrounds |
| Arm Teal | (46,125,125) | `2E7D7D` | Experimental arm 1 header |
| Arm Blue | (68,114,196) | `4472C4` | Experimental arm 2 header |
| Control Gray | (128,128,128) | `808080` | Control arm header |
| Navy | (27,42,74) | `1B2A4A` | Conclusion highlight box |
| Light BG | (240,244,248) | `F0F4F8` | Eligibility/endpoint box fills |
| Border Gray | (204,204,204) | `CCCCCC` | Box borders, table borders |

## Available Layouts (Master 0)

| Layout Name | Placeholders | Best For |
|-------------|-------------|----------|
| `Title Slide` | idx=0 (center title), idx=1 (subtitle) | Slide 1 only |
| `Title and Content` | idx=0 (title), idx=1 (bullets) | Most content slides |
| `Title Only` | idx=0 (title) | Diagrams, tables, figures |
| `Two Content` | idx=0 (title), idx=1 (left), idx=2 (right) | Split demographics |
| `Blank` | none | Fully custom slides |
| `Comparison` | idx=0-4 (title, 2 text, 2 content) | Side-by-side analysis |

## Key Dimensions (EMU)

The model uses 13.33" × 7.50" slides (12192000 × 6858000 EMU). Key reference positions:

| Element | EMU | Inches |
|---------|-----|--------|
| Content left margin | 838200 | 0.92" |
| Content top (below title) | 1825625 | 2.00" |
| Title placeholder top | 49440 | 0.05" |
| Title placeholder height | 1325563 | 1.45" |
| Bottom bar top | 6721475 | 7.35" |
| Slide width | 12192000 | 13.33" |
| Slide height | 6858000 | 7.50" |

## Font Size Requirements (MANDATORY MINIMUMS)

These are hard minimums. If content doesn't fit, split across slides rather than shrinking.

| Element | Preferred | Minimum |
|---------|-----------|---------|
| Body text ("Title and Content" level 0) | 18pt | 18pt |
| Sub-bullets ("Title and Content" level 1) | 16pt | 16pt |
| Table cells (headers and data) | 16pt | 14pt |
| Study design diagram text (shapes/boxes) | 14pt | 12pt |
| Key message / italic subtitle | 14pt | 14pt |
| Citation text | 10pt | 10pt |
| Study badge text | 11pt | 11pt |

## Table Centering

Tables must be horizontally centered. Use this formula:

```python
total_table_w = sum(col_widths)
table_x = (Emu(12192000) - total_table_w) // 2  # center on slide
```
