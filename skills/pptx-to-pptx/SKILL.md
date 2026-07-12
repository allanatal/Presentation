---
name: pptx-to-pptx
description: "Convert an existing PowerPoint presentation into the user's own template and visual style. Use this skill when the user uploads a .pptx file and asks to reformat, restyle, or convert it to match their template. Triggers include: 'convert this presentation to my template', 'restyle these slides', 'reformat this deck in my style', 'apply my template to this presentation', 'make this match my slide format', or any request to transform an existing .pptx into a different visual identity while preserving scientific content. Does NOT trigger for creating presentations from PDFs or papers — use the academic-paper-to-pptx skill for that."
---

# PPTX-to-PPTX Style Conversion Skill

Convert an existing PowerPoint presentation into the **Moffitt Cancer Center slide template**, preserving scientific content and logical structure while applying consistent formatting, typography, and visual identity.

## Architecture

This skill **parses an existing .pptx** (the "source"), extracts its content **element-by-element**, classifies each element's type and role, and rebuilds the presentation using the shared Moffitt template and slide-builder patterns.

**Key library**: `python-pptx` for both reading and writing.

**Shared resources** — this skill references files from the `academic-paper-to-pptx` skill:
- `/mnt/skills/user/academic-paper-to-pptx/references/template.pptx` — Moffitt slide master
- `/mnt/skills/user/academic-paper-to-pptx/references/slide-builders.md` — python-pptx code patterns
- `/mnt/skills/user/academic-paper-to-pptx/references/style-spec.md` — colors, positions, font rules

**Always read both `slide-builders.md` and `style-spec.md` before writing any code.** They contain the authoritative font sizes, color values, and positioning constants.

## Editability Principle (MANDATORY)

**Never rasterize an entire slide.** Process each shape individually. Text, tables, and simple shapes must always be rebuilt as editable elements. Only rasterize the **specific shape** that cannot be recreated: embedded chart objects (`shape.has_chart`), and complex grouped-shape diagrams where spatial fidelity matters more than editability.

A slide with 5 elements — a title, two tables, a textbox, and a conclusion shape — must produce 5 rebuilt editable elements, not 1 screenshot.

**Element-level routing rules:**

| Element type | Action | Editable? |
|---|---|---|
| Text shape (title, body, subtitle, footnote, caption, callout) | Rebuild as textbox or placeholder | ✅ Yes |
| Table | Rebuild via `add_table()` | ✅ Yes |
| Simple auto-shape with text (rectangle, oval, rounded rect, arrow) | Rebuild via `add_shape()` with text | ✅ Yes |
| Simple auto-shape without text (lines, connectors, decorative) | Rebuild as shape, or skip if purely decorative from source template | ✅ Yes |
| Picture (photo, logo, illustration) | Re-embed extracted image via `add_picture()` | ✅ Yes (image) |
| Chart object (`shape.has_chart`) | Rasterize **only this shape's bounding box**, embed as image | ❌ No |
| Complex group shape (many nested children forming a flow diagram with connectors/arrows) | Rasterize **only this shape's bounding box**, embed as image | ❌ No |
| Simple group shape (few children, all text/rectangles) | Decompose and rebuild each child shape individually | ✅ Yes |

**The test for groups:** If a group shape contains ≤ 6 children and all children are simple shapes (rectangles, ovals, text boxes, connectors), decompose and rebuild each child. If it has > 6 children with complex nesting, curved connectors, or freeform paths, rasterize the group's bounding box only.

## Prerequisites

```bash
pip install python-pptx Pillow --break-system-packages
```

Read the pptx skill at `/mnt/skills/public/pptx/SKILL.md` for QA workflow and image conversion utilities.

---

## Workflow — Four Phases

1. **Parse** — Extract all content from the source presentation, element by element
2. **Classify** — Classify each element's type and role; choose a layout per slide
3. **Build** — Reconstruct each slide by placing every element individually
4. **QA** — Convert to images and visually inspect

---

## Phase 1: Parse the Source Presentation

### Step 1a: Visual Overview

Generate thumbnails of the source to understand the deck at a glance:

```bash
python /mnt/skills/public/pptx/scripts/thumbnail.py /path/to/source.pptx
```

View the thumbnail grid, then also extract text:

```bash
extract-text /path/to/source.pptx
```

### Step 1b: Extract Content with python-pptx

Open the source and iterate through every slide, extracting structured content:

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image
import os, json

src = Presentation("/path/to/source.pptx")
src_slide_width = src.slide_width   # EMU
src_slide_height = src.slide_height # EMU
slides_data = []

for slide_idx, slide in enumerate(src.slides):
    slide_info = {
        "index": slide_idx,
        "layout_name": slide.slide_layout.name if slide.slide_layout else "Unknown",
        "elements": []
    }

    for shape in slide.shapes:
        element = extract_shape(shape, slide_idx)
        if element:
            slide_info["elements"].append(element)

    slides_data.append(slide_info)
```

### Step 1c: Shape Extraction Logic

For each shape on a slide, determine its type and extract content. Also extract fill color, line color, and text alignment for shapes that will be rebuilt:

```python
def extract_shape(shape, slide_idx):
    """Extract content from a single shape."""
    el = {
        "shape_type": str(shape.shape_type),
        "name": shape.name,
        "position": {"left": shape.left, "top": shape.top,
                      "width": shape.width, "height": shape.height},
    }

    # ── Text shapes (titles, subtitles, body, captions) ──
    if shape.has_text_frame:
        paragraphs = []
        for para in shape.text_frame.paragraphs:
            p_data = {
                "text": para.text,
                "level": para.level,
                "bold": any(run.font.bold for run in para.runs if run.font.bold),
                "italic": any(run.font.italic for run in para.runs if run.font.italic),
                "font_size": None,
                "font_color": None,
                "alignment": str(para.alignment) if para.alignment else None,
            }
            # Capture font size and color from first run
            for run in para.runs:
                if run.font.size:
                    p_data["font_size"] = run.font.size.pt
                if run.font.color and run.font.color.rgb:
                    p_data["font_color"] = str(run.font.color.rgb)
                if p_data["font_size"]:
                    break
            paragraphs.append(p_data)

        el["type"] = "text"
        el["paragraphs"] = paragraphs
        el["full_text"] = shape.text_frame.text
        el["word_wrap"] = shape.text_frame.word_wrap
        el["role"] = classify_text_role(shape, paragraphs)

        # Extract fill for shapes that are styled boxes (conclusion boxes, callout boxes)
        if hasattr(shape, 'fill') and shape.fill.type is not None:
            try:
                el["fill_color"] = str(shape.fill.fore_color.rgb) if shape.fill.fore_color else None
            except:
                el["fill_color"] = None

        # Detect auto-shape type for styled text boxes (rectangles, rounded rects, ovals)
        if hasattr(shape, 'auto_shape_type'):
            el["auto_shape_type"] = str(shape.auto_shape_type)
            el["type"] = "shape_with_text"  # distinguish from plain text boxes

    # ── Tables ──
    elif shape.has_table:
        tbl = shape.table
        rows = []
        for row in tbl.rows:
            row_data = [cell.text for cell in row.cells]
            rows.append(row_data)

        # Also extract column widths to preserve proportions
        col_widths = [col.width for col in tbl.columns]

        # Extract header cell fill colors
        header_colors = []
        if len(tbl.rows) > 0:
            for cell in tbl.rows[0].cells:
                try:
                    if cell.fill and cell.fill.fore_color:
                        header_colors.append(str(cell.fill.fore_color.rgb))
                    else:
                        header_colors.append(None)
                except:
                    header_colors.append(None)

        el["type"] = "table"
        el["rows"] = rows
        el["n_rows"] = len(tbl.rows)
        el["n_cols"] = len(tbl.columns)
        el["col_widths"] = col_widths
        el["header_colors"] = header_colors

    # ── Images ──
    elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        image = shape.image
        ext = image.content_type.split("/")[-1]
        if ext == "jpeg":
            ext = "jpg"
        img_path = f"/home/claude/source_images/slide{slide_idx}_{shape.name}.{ext}"
        os.makedirs(os.path.dirname(img_path), exist_ok=True)
        with open(img_path, "wb") as f:
            f.write(image.blob)
        el["type"] = "image"
        el["image_path"] = img_path
        el["content_type"] = image.content_type

    # ── Charts (rasterize ONLY this shape) ──
    elif shape.has_chart:
        el["type"] = "chart"
        el["chart_type"] = str(shape.chart.chart_type)
        el["needs_rasterization"] = True

    # ── Grouped shapes ──
    elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        el["type"] = "group"
        el["child_count"] = len(shape.shapes)
        el["children"] = [extract_shape(child, slide_idx) for child in shape.shapes]
        el["children"] = [c for c in el["children"] if c]

        # Decide: decompose or rasterize?
        # Simple groups (≤6 children, all text/rect/oval): decompose
        # Complex groups (many children, freeforms, nested groups): rasterize
        simple_types = {"text", "shape_with_text", "shape", "image"}
        if el["child_count"] <= 6 and all(c.get("type") in simple_types for c in el["children"]):
            el["decompose"] = True
        else:
            el["decompose"] = False
            el["needs_rasterization"] = True

    # ── Other shapes (arrows, lines, simple auto-shapes) ──
    else:
        el["type"] = "shape"
        el["auto_shape_type"] = str(getattr(shape, 'auto_shape_type', 'unknown'))
        # Extract fill if present
        if hasattr(shape, 'fill') and shape.fill.type is not None:
            try:
                el["fill_color"] = str(shape.fill.fore_color.rgb) if shape.fill.fore_color else None
            except:
                el["fill_color"] = None

    return el
```

### Step 1d: Text Role Classification

Classify text elements by their semantic role using position, font size, and placeholder type:

```python
def classify_text_role(shape, paragraphs):
    """Determine if a text shape is a title, subtitle, body, caption, or footnote."""
    # Check if it's a placeholder
    if shape.is_placeholder:
        ph_type = shape.placeholder_format.type
        # Placeholder type constants:
        # 1 = CENTER_TITLE, 2 = SUBTITLE, 13 = TITLE, 15 = BODY
        if ph_type in (1, 13, 15):  # TITLE types
            return "title"
        elif ph_type == 2:
            return "subtitle"
        elif ph_type in (7, 8):  # BODY types
            return "body"

    # Heuristic classification based on position and size
    top_inches = shape.top / 914400  # EMU to inches
    height_inches = shape.height / 914400
    avg_font = None
    for p in paragraphs:
        if p.get("font_size"):
            avg_font = p["font_size"]
            break

    # Title: near top, large font
    if top_inches < 1.5 and avg_font and avg_font >= 20:
        return "title"

    # Subtitle / key message: near top, medium font, italic common
    if top_inches < 2.5 and avg_font and 12 <= avg_font < 20:
        return "subtitle"

    # Footnote / caption: near bottom, small font
    if top_inches > 5.5 or (avg_font and avg_font <= 10):
        return "footnote"

    # Default: body text
    return "body"
```

### Step 1e: Rasterize Individual Shapes (Charts and Complex Groups Only)

When a shape requires rasterization (chart object or complex group that cannot be decomposed), rasterize **only that shape's bounding box** from a full-slide image. Never embed a full-slide screenshot.

First, rasterize the source slides that contain shapes needing rasterization:

```bash
# Convert source to PDF, then rasterize affected slides at 300 DPI
python /mnt/skills/public/pptx/scripts/office/soffice.py --headless --convert-to pdf source.pptx
pdftoppm -jpeg -r 300 -f <SLIDE_NUM> -l <SLIDE_NUM> source.pdf /home/claude/source_images/raster_slide
```

Then crop **only the specific shape's region** using its EMU position:

```python
from PIL import Image

def crop_shape_from_slide(slide_image_path, shape_position, slide_width, slide_height):
    """Crop a single shape's bounding box from a rasterized slide image."""
    img = Image.open(slide_image_path)
    img_w, img_h = img.size

    # Convert EMU positions to pixel coordinates
    scale_x = img_w / slide_width
    scale_y = img_h / slide_height

    left = int(shape_position["left"] * scale_x)
    top = int(shape_position["top"] * scale_y)
    right = int((shape_position["left"] + shape_position["width"]) * scale_x)
    bottom = int((shape_position["top"] + shape_position["height"]) * scale_y)

    # Add small padding
    pad = 10
    left = max(0, left - pad)
    top = max(0, top - pad)
    right = min(img_w, right + pad)
    bottom = min(img_h, bottom + pad)

    cropped = img.crop((left, top, right, bottom))
    output_path = slide_image_path.replace(".jpg", f"_shape_{shape_position['left']}.jpg")
    cropped.save(output_path, quality=95)
    return output_path
```

After cropping, store the image path back on the element so the builder can embed it:

```python
element["image_path"] = crop_shape_from_slide(raster_path, element["position"], src_slide_width, src_slide_height)
element["type"] = "rasterized_shape"  # mark for the builder
```

---

## Phase 2: Classify Elements and Choose Layouts

### Element-Level Classification (NOT Slide-Level)

**Do NOT classify slides into a single type.** Instead, classify each element individually using the type assigned during extraction. The slide layout is chosen based on what combination of elements are present:

| Elements present on slide | Moffitt Layout |
|---|---|
| Only title + optional subtitle, no body/table/image | `Title Slide` |
| Title + body text in placeholder (no tables, no images) | `Title and Content` |
| Title + any combination of tables, images, shapes, textboxes | `Title Only` |
| Title + two distinct content regions side by side | `Two Content` |
| Mostly blank or decorative | `Title Slide` or skip |

```python
def choose_layout(slide_data):
    """Choose a Moffitt layout based on what elements are present."""
    elements = slide_data["elements"]

    has_title = any(e.get("role") == "title" for e in elements if e["type"] == "text")
    has_body_placeholder = any(
        e.get("role") == "body" and e["type"] == "text"
        for e in elements
    )
    has_table = any(e["type"] == "table" for e in elements)
    has_image = any(e["type"] in ("image", "chart", "rasterized_shape") for e in elements)
    has_shape_with_text = any(e["type"] == "shape_with_text" for e in elements)
    has_group = any(e["type"] == "group" for e in elements)

    # Count non-title, non-footnote content elements
    content_count = sum(1 for e in elements
                        if e["type"] in ("table", "image", "shape_with_text", "group", "chart", "rasterized_shape")
                        or (e["type"] == "text" and e.get("role") in ("body", "subtitle")))

    # Title-only slide
    if has_title and content_count == 0:
        return "Title Slide"

    # Simple content slide: title + body text, no tables/images/shapes
    if has_title and has_body_placeholder and not has_table and not has_image and not has_shape_with_text and not has_group:
        return "Title and Content"

    # Everything else: use Title Only and place elements manually
    return "Title Only"
```

### Content Trimming Strategy

Since the output must keep the **same slide count**, content that exceeds the Moffitt font size minimums must be trimmed. Rules:

1. **Tables**: If a table at 14pt minimum exceeds the available slide area, reduce the number of visible rows by prioritizing the header row and the most important data rows. Add a footnote like "Selected data shown; see original for full table."
2. **Bullet text**: If bullet content at 18pt doesn't fit, consolidate sub-bullets into fewer items. Preserve the key message of each bullet.
3. **Figures**: Always preserve. Resize to fit available area while maintaining aspect ratio.
4. **Never reduce font sizes below the minimums** defined in `style-spec.md`.

### Position Mapping

When rebuilding elements, map source positions to Moffitt template dimensions. The source may use a different slide size. Scale coordinates proportionally:

```python
# Target Moffitt dimensions
SLIDE_WIDTH = 12192000   # 13.33 inches in EMU
SLIDE_HEIGHT = 6858000   # 7.50 inches in EMU

def map_position(src_pos, src_slide_w, src_slide_h):
    """Scale source EMU positions to Moffitt slide dimensions."""
    scale_x = SLIDE_WIDTH / src_slide_w
    scale_y = SLIDE_HEIGHT / src_slide_h
    return {
        "left": int(src_pos["left"] * scale_x),
        "top": int(src_pos["top"] * scale_y),
        "width": int(src_pos["width"] * scale_x),
        "height": int(src_pos["height"] * scale_y),
    }
```

Use `map_position` for every element placed on the output slide. This preserves the relative layout from the source while fitting the Moffitt slide dimensions.

**Exception:** Title placeholders and body placeholders are placed using the template's built-in placeholder positions, not mapped from source. Only manually-placed elements (textboxes, tables, shapes, images) use position mapping.

---

## Phase 3: Build the Output Presentation

### Step 3a: Setup

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import shutil

# Read shared references first:
# view /mnt/skills/user/academic-paper-to-pptx/references/slide-builders.md
# view /mnt/skills/user/academic-paper-to-pptx/references/style-spec.md

TEMPLATE = "/mnt/skills/user/academic-paper-to-pptx/references/template.pptx"
shutil.copy(TEMPLATE, "/home/claude/output.pptx")
prs = Presentation("/home/claude/output.pptx")

master = prs.slide_masters[0]
LAYOUTS = {l.name: l for l in master.slide_layouts}

# Colors (from style-spec.md)
TITLE_BLUE = RGBColor(0x00, 0x33, 0x66)
RED = RGBColor(0xFF, 0x00, 0x00)
BODY = RGBColor(0x33, 0x33, 0x33)
GRAY = RGBColor(0x99, 0x99, 0x99)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
NAVY = RGBColor(0x1B, 0x2A, 0x4A)
LIGHT_BG = RGBColor(0xF0, 0xF4, 0xF8)

FONT = "Arial"
SLIDE_WIDTH = 12192000  # EMU
SLIDE_HEIGHT = 6858000  # EMU
```

### Step 3b: Composite Slide Builder (One Builder for All Slides)

Instead of separate builders per slide type, use **one composite builder** that iterates over every element on a slide and places each one individually. This ensures every editable element remains editable.

```python
def build_slide(prs, slide_data, study_name="", citation="",
                src_slide_w=None, src_slide_h=None):
    """Build a single slide by placing every element individually."""
    layout_name = choose_layout(slide_data)
    slide = prs.slides.add_slide(LAYOUTS[layout_name])

    elements = slide_data["elements"]

    # ── Place title ──
    title_el = next((e for e in elements if e["type"] == "text" and e.get("role") == "title"), None)
    if title_el and 0 in slide.placeholders:
        slide.placeholders[0].text = title_el["full_text"]
        set_title_font(slide.placeholders[0], 24)

    # ── Place body into placeholder (only for "Title and Content" layout) ──
    if layout_name == "Title and Content":
        body_els = [e for e in elements if e["type"] == "text" and e.get("role") == "body"]
        if body_els and 1 in slide.placeholders:
            tf = slide.placeholders[1].text_frame
            tf.clear()
            first = True
            for body_el in body_els:
                for para in body_el.get("paragraphs", []):
                    if not para["text"].strip():
                        continue
                    p = tf.paragraphs[0] if first else tf.add_paragraph()
                    first = False
                    p.text = para["text"]
                    p.level = min(para.get("level", 0), 1)
                    p.font.name = FONT
                    p.font.size = Pt(18) if p.level == 0 else Pt(16)
                    if para.get("bold"):
                        p.font.bold = True

    # ── Place remaining elements (tables, textboxes, shapes, images) ──
    if layout_name == "Title Only":
        for el in elements:
            if el["type"] == "text" and el.get("role") == "title":
                continue  # already placed
            pos = map_position(el["position"], src_slide_w, src_slide_h) if src_slide_w else el["position"]
            place_element(slide, el, pos)

    # ── Footnote elements (place regardless of layout) ──
    for el in elements:
        if el["type"] == "text" and el.get("role") == "footnote":
            pos = map_position(el["position"], src_slide_w, src_slide_h) if src_slide_w else el["position"]
            place_textbox(slide, el, pos)

    add_badge_and_citation(slide, study_name, citation)
    return slide
```

### Step 3c: Element Placement Functions

Each element type has its own placement function:

```python
def place_element(slide, el, pos):
    """Route an element to the correct placement function."""
    if el["type"] == "text" and el.get("role") != "title":
        place_textbox(slide, el, pos)
    elif el["type"] == "shape_with_text":
        place_shape_with_text(slide, el, pos)
    elif el["type"] == "table":
        place_table(slide, el, pos)
    elif el["type"] == "image":
        place_image(slide, el, pos)
    elif el["type"] in ("chart", "rasterized_shape"):
        place_rasterized(slide, el, pos)
    elif el["type"] == "group":
        place_group(slide, el, pos)
    elif el["type"] == "shape":
        place_simple_shape(slide, el, pos)


def place_textbox(slide, el, pos):
    """Rebuild a text shape as an editable textbox."""
    tb = slide.shapes.add_textbox(pos["left"], pos["top"], pos["width"], pos["height"])
    tb.text_frame.word_wrap = el.get("word_wrap", True)
    first = True
    for para in el.get("paragraphs", []):
        if not para["text"].strip():
            continue
        p = tb.text_frame.paragraphs[0] if first else tb.text_frame.add_paragraph()
        first = False
        p.text = para["text"]
        p.level = para.get("level", 0)
        p.font.name = FONT

        # Apply font size with minimums
        src_size = para.get("font_size")
        role = el.get("role", "body")
        if role == "footnote":
            p.font.size = Pt(max(src_size or 10, 10))
        elif role == "subtitle":
            p.font.size = Pt(max(src_size or 14, 14))
        else:
            p.font.size = Pt(max(src_size or 18, 16))

        if para.get("bold"):
            p.font.bold = True
        if para.get("italic"):
            p.font.italic = True

        # Apply color — use Moffitt palette equivalent
        p.font.color.rgb = BODY


def place_shape_with_text(slide, el, pos):
    """Rebuild an auto-shape (rectangle, oval, rounded rect) with text as an editable shape."""
    # Map source auto-shape type to MSO_SHAPE enum
    shape_type_str = el.get("auto_shape_type", "RECTANGLE")
    shape_map = {
        "RECTANGLE (1)": MSO_SHAPE.RECTANGLE,
        "ROUNDED_RECTANGLE (5)": MSO_SHAPE.ROUNDED_RECTANGLE,
        "OVAL (9)": MSO_SHAPE.OVAL,
    }
    mso_type = MSO_SHAPE.ROUNDED_RECTANGLE  # safe default
    for key, val in shape_map.items():
        if key in shape_type_str:
            mso_type = val
            break

    shape = slide.shapes.add_shape(mso_type, pos["left"], pos["top"], pos["width"], pos["height"])

    # Apply fill color
    fill_color = el.get("fill_color")
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(fill_color)
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = LIGHT_BG

    shape.line.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
    shape.line.width = Pt(0.5)

    # Add text
    shape.text_frame.word_wrap = True
    first = True
    for para in el.get("paragraphs", []):
        if not para["text"].strip():
            continue
        p = shape.text_frame.paragraphs[0] if first else shape.text_frame.add_paragraph()
        first = False
        p.text = para["text"]
        p.font.name = FONT
        p.font.size = Pt(max(para.get("font_size") or 14, 12))
        if para.get("bold"):
            p.font.bold = True
        p.font.color.rgb = BODY


def place_table(slide, el, pos):
    """Rebuild a table as an editable table shape."""
    rows = el["rows"]
    n_rows = len(rows)
    n_cols = len(rows[0]) if rows else 0
    if n_rows == 0 or n_cols == 0:
        return

    # Use extracted column widths if available, otherwise distribute evenly
    if el.get("col_widths"):
        # Scale column widths to Moffitt slide
        src_total = sum(el["col_widths"])
        available = Inches(11.5)
        col_widths = [int(w / src_total * available) for w in el["col_widths"]]
    else:
        col_w = int(Inches(11.5) / n_cols)
        col_widths = [col_w] * n_cols

    total_w = sum(col_widths)
    table_x = (SLIDE_WIDTH - total_w) // 2  # center horizontally

    # Limit rows to fit at 14pt
    max_rows = 14
    trimmed = False
    if n_rows > max_rows:
        rows = rows[:max_rows]
        n_rows = max_rows
        trimmed = True

    tbl_shape = slide.shapes.add_table(
        n_rows, n_cols,
        table_x, pos["top"],
        total_w, Emu(n_rows * 320000)
    )
    tbl = tbl_shape.table

    for i, w in enumerate(col_widths):
        tbl.columns[i].width = w

    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = str(val)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(14)  # MINIMUM 14pt
                p.font.name = FONT
                if r == 0:
                    p.font.bold = True
                    p.font.color.rgb = WHITE
                    p.alignment = PP_ALIGN.CENTER
                else:
                    p.font.color.rgb = BODY
                    if c >= 1:
                        p.alignment = PP_ALIGN.CENTER

            # Header row fill
            if r == 0:
                cell.fill.solid()
                # Use extracted header color if available
                header_colors = el.get("header_colors", [])
                if c < len(header_colors) and header_colors[c]:
                    try:
                        cell.fill.fore_color.rgb = RGBColor.from_string(header_colors[c])
                    except:
                        cell.fill.fore_color.rgb = NAVY
                else:
                    cell.fill.fore_color.rgb = NAVY

    if trimmed:
        note = slide.shapes.add_textbox(
            table_x, Emu(pos["top"] + n_rows * 320000 + 100000),
            total_w, Emu(300000)
        )
        p = note.text_frame.paragraphs[0]
        p.text = "Selected data shown; see original for full table."
        p.font.size = Pt(12)
        p.font.italic = True
        p.font.color.rgb = GRAY
        p.font.name = FONT


def place_image(slide, el, pos):
    """Re-embed an extracted image."""
    from PIL import Image as PILImage

    img_path = el.get("image_path")
    if not img_path or not os.path.exists(img_path):
        return

    img = PILImage.open(img_path)
    img_w, img_h = img.size
    ratio = img_w / img_h

    # Use mapped position but constrain to available area
    max_w = pos["width"]
    max_h = pos["height"]

    w = max_w
    h = int(w / ratio)
    if h > max_h:
        h = max_h
        w = int(h * ratio)

    # Center within the allocated area
    x = pos["left"] + (max_w - w) // 2
    y = pos["top"] + (max_h - h) // 2

    slide.shapes.add_picture(img_path, x, y, w, h)


def place_rasterized(slide, el, pos):
    """Place a rasterized shape region (chart or complex group)."""
    img_path = el.get("image_path")
    if not img_path or not os.path.exists(img_path):
        return
    place_image(slide, el, pos)


def place_group(slide, el, pos):
    """Handle group shapes: decompose simple ones, rasterize complex ones."""
    if el.get("decompose", False):
        # Rebuild each child shape individually
        for child in el.get("children", []):
            child_pos = child.get("position", pos)
            place_element(slide, child, child_pos)
    else:
        # Complex group — use rasterized image
        place_rasterized(slide, el, pos)


def place_simple_shape(slide, el, pos):
    """Rebuild a simple auto-shape without text (line, arrow, decorative rectangle)."""
    # Skip shapes that are clearly source-template decorations
    name = el.get("name", "").lower()
    if any(skip in name for skip in ["background", "rectangle 9", "picture 9"]):
        return  # source template decoration, skip

    shape_type_str = el.get("auto_shape_type", "RECTANGLE")
    if "LINE" in shape_type_str or "CONNECTOR" in shape_type_str:
        return  # skip connectors/lines for simplicity

    # Only rebuild if the shape appears to carry content (has fill, is visible)
    fill = el.get("fill_color")
    if fill:
        shape_map = {"RECTANGLE": MSO_SHAPE.RECTANGLE, "ROUNDED_RECTANGLE": MSO_SHAPE.ROUNDED_RECTANGLE, "OVAL": MSO_SHAPE.OVAL}
        mso_type = MSO_SHAPE.RECTANGLE
        for key, val in shape_map.items():
            if key in shape_type_str:
                mso_type = val
                break
        shape = slide.shapes.add_shape(mso_type, pos["left"], pos["top"], pos["width"], pos["height"])
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(fill)
        shape.line.fill.background()
```

### Step 3d: Title Font Helper

Explicitly set title font size to prevent the template's auto-sizing from creating oversized titles:

```python
def set_title_font(placeholder, font_size=24):
    """Set explicit title font size to prevent auto-sizing overflow."""
    placeholder.text_frame.auto_size = None
    for para in placeholder.text_frame.paragraphs:
        para.font.size = Pt(font_size)
        para.font.name = FONT
```

### Step 3e: Badge and Citation

Add the study badge and citation on every content slide (not the title slide). If the source presentation has a visible citation or reference, use that text. Otherwise use a generic format:

```python
def add_badge_and_citation(slide, study_name="", citation_text=""):
    """Add study name badge (top-right) and citation (bottom-right)."""
    if study_name:
        badge = slide.shapes.add_textbox(Emu(9984826), Emu(157655), Emu(1671098), Emu(369332))
        p = badge.text_frame.paragraphs[0]
        p.text = study_name
        p.font.size = Pt(11)
        p.font.color.rgb = TITLE_BLUE
        p.alignment = PP_ALIGN.RIGHT

    if citation_text:
        cite = slide.shapes.add_textbox(Emu(9987101), Emu(6463328), Emu(2204899), Emu(307777))
        p = cite.text_frame.paragraphs[0]
        p.text = citation_text
        p.font.size = Pt(10)
        p.font.color.rgb = GRAY
        p.alignment = PP_ALIGN.RIGHT
```

---

## Phase 4: QA

```bash
python /mnt/skills/public/pptx/scripts/office/soffice.py --headless --convert-to pdf output.pptx
rm -f slide-*.jpg
pdftoppm -jpeg -r 150 output.pdf slide
ls -1 "$PWD"/slide-*.jpg
```

**Check for:**
1. Same slide count as source
2. All titles transferred correctly
3. Tables readable at 14pt+ and centered
4. Images/charts present and properly sized
5. No text overflow or cut-off
6. Branding elements on every slide (inherited from template)
7. Badge and citation on content slides
8. **All tables, textboxes, and shapes are editable (not rasterized)**

**Compare source vs output side-by-side** by viewing both thumbnail grids.

Fix issues and re-verify once.

---

## Handling Special Cases

### Slides with Mixed Content (Text + Image + Table)

Every element is placed individually, so mixed-content slides are handled naturally by the composite builder. Each table, textbox, image, and shape gets its own placement call. Use `map_position` to preserve the source layout.

When elements overlap or don't fit at minimum font sizes:
1. Title (always preserved via placeholder)
2. Tables (always rebuilt as editable)
3. Images (always re-embedded)
4. Body textboxes (trim text if needed to make room)
5. Decorative shapes (skip if needed)

### Decorative Elements

Skip decorative shapes that don't carry content: background rectangles, decorative lines, logo images from the source template, empty text boxes. These are replaced by the Moffitt template's built-in branding.

Detect source-template decorations by checking:
- Shape name contains "background", "logo", "footer", "header" (from source template)
- Shape has no text and no meaningful fill
- Shape covers the full slide width (likely a background bar)

### Speaker Notes

If the source has speaker notes, preserve them on the output slides:

```python
if source_slide.notes_slide and source_slide.notes_slide.notes_text_frame:
    notes_text = source_slide.notes_slide.notes_text_frame.text
    if notes_text.strip():
        output_slide.notes_slide.notes_text_frame.text = notes_text
```

### Animations and Transitions

These are **not preserved**. The output is a clean, static presentation. Mention this to the user if the source appears animation-heavy.

---

## Common Pitfalls

- **Never rasterize entire slides**: Every text element, table, and simple shape must remain editable. Only rasterize the specific chart object or complex group shape that cannot be recreated — and crop only its bounding box from the rasterized slide image.
- **Chart objects**: Chart objects (`shape.has_chart`) cannot be reliably recreated programmatically across templates. Rasterize only the chart shape's bounding box and embed as an image. All other elements on the same slide (titles, textboxes, tables) must still be rebuilt as editable.
- **Group shapes**: First check if the group contains ≤ 6 children that are all simple shapes (rectangles, ovals, text boxes). If so, decompose and rebuild each child individually. Only rasterize if the group is truly complex (many nested children, freeform paths, curved connectors) where spatial fidelity matters more than editability.
- **Source template artifacts**: The source may have logos, bars, footers, or watermarks baked into the slide master — these will be replaced by the Moffitt template automatically, but check for logos placed as regular shapes (not in the master) that should be removed
- **Font substitution**: The source may use fonts not available in the build environment — python-pptx will set the font name, but LibreOffice may substitute during PDF conversion for QA. This doesn't affect the .pptx output opened in PowerPoint
- **Encoding**: Watch for special characters (em-dashes, smart quotes, superscripts) in source text — preserve them exactly
- **Empty slides**: Some presentations have blank separator slides — convert these to Moffitt `Title Slide` layout with the section name, or skip if truly empty
- **Title overflow**: Always call `set_title_font(placeholder, 24)` on title placeholders to prevent the template's auto-sizing from creating oversized text that overflows the slide

## Font Size Requirements (MANDATORY)

These minimums are inherited from the shared `style-spec.md` and must never be violated:

| Element | Minimum |
|---------|---------|
| Body text (level 0) | 18pt |
| Sub-bullets (level 1) | 16pt |
| Table cells | 14pt |
| Study design / diagram text | 12pt |
| Key message subtitle | 14pt |
| Citation | 10pt |
| Study badge | 11pt |

If source content doesn't fit at these sizes, **trim content rather than shrink fonts**.
