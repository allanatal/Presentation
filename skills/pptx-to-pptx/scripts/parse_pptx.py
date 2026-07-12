#!/usr/bin/env python3
"""
Parse a source .pptx and output structured JSON describing each slide's content.

Usage:
    python parse_pptx.py source.pptx [--output parsed.json] [--images-dir /home/claude/source_images]

Output: JSON array where each element represents a slide with:
- index: slide number (0-based)
- layout_name: source layout name
- slide_type: classified type (title_slide, content_slide, table_slide, etc.)
- elements: list of extracted shapes with type, role, content
"""

import argparse
import json
import os
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def classify_text_role(shape, paragraphs):
    """Determine semantic role of a text shape."""
    # Check placeholder type first
    if shape.is_placeholder:
        ph_idx = shape.placeholder_format.idx
        ph_type = shape.placeholder_format.type
        # Standard placeholder types
        if ph_type is not None:
            pt_val = int(ph_type)
            if pt_val in (1, 13, 15):  # CENTER_TITLE, TITLE, BODY (as title)
                return "title"
            elif pt_val == 2:  # SUBTITLE
                return "subtitle"
        # By index convention: 0 = title, 1 = body/content
        if ph_idx == 0:
            return "title"
        elif ph_idx == 1:
            return "body"

    # Heuristic: position + font size
    top_inches = shape.top / 914400 if shape.top else 0
    avg_font = None
    for p in paragraphs:
        if p.get("font_size"):
            avg_font = p["font_size"]
            break

    if top_inches < 1.5 and avg_font and avg_font >= 20:
        return "title"
    if top_inches < 2.5 and avg_font and 12 <= avg_font < 20:
        return "subtitle"
    if top_inches > 5.5 or (avg_font and avg_font <= 10):
        return "footnote"

    return "body"


def extract_shape(shape, slide_idx, images_dir):
    """Extract content from a single shape."""
    el = {
        "name": shape.name,
        "position": {
            "left": shape.left,
            "top": shape.top,
            "width": shape.width,
            "height": shape.height,
        },
    }

    # Text shapes
    if shape.has_text_frame:
        paragraphs = []
        for para in shape.text_frame.paragraphs:
            p_data = {
                "text": para.text,
                "level": para.level,
                "bold": any(
                    run.font.bold for run in para.runs if run.font.bold is True
                ),
                "font_size": None,
            }
            for run in para.runs:
                if run.font.size:
                    p_data["font_size"] = run.font.size.pt
                    break
            paragraphs.append(p_data)

        el["type"] = "text"
        el["paragraphs"] = paragraphs
        el["full_text"] = shape.text_frame.text.strip()
        el["role"] = classify_text_role(shape, paragraphs)

        # Skip empty text shapes
        if not el["full_text"]:
            return None

    # Tables
    elif shape.has_table:
        tbl = shape.table
        rows = []
        for row in tbl.rows:
            row_data = [cell.text for cell in row.cells]
            rows.append(row_data)
        el["type"] = "table"
        el["rows"] = rows
        el["n_rows"] = len(tbl.rows)
        el["n_cols"] = len(tbl.columns)

    # Images
    elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        try:
            image = shape.image
            ext = image.content_type.split("/")[-1]
            if ext == "jpeg":
                ext = "jpg"
            img_filename = f"slide{slide_idx}_{shape.name.replace(' ', '_')}.{ext}"
            img_path = os.path.join(images_dir, img_filename)
            with open(img_path, "wb") as f:
                f.write(image.blob)
            el["type"] = "image"
            el["image_path"] = img_path
            el["content_type"] = image.content_type
        except Exception as e:
            el["type"] = "image"
            el["image_path"] = None
            el["error"] = str(e)

    # Charts
    elif shape.has_chart:
        el["type"] = "chart"
        try:
            el["chart_type"] = str(shape.chart.chart_type)
        except Exception:
            el["chart_type"] = "unknown"
        el["needs_rasterization"] = True

    # Groups
    elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        el["type"] = "group"
        el["child_count"] = len(shape.shapes)
        children = []
        for child in shape.shapes:
            child_el = extract_shape(child, slide_idx, images_dir)
            if child_el:
                children.append(child_el)
        el["children"] = children

    # Other shapes
    else:
        el["type"] = "shape"
        el["auto_shape_type"] = str(
            getattr(shape, "auto_shape_type", "unknown")
        )
        # Skip purely decorative shapes with no text
        if not shape.has_text_frame:
            return None

    return el


def classify_slide(slide_info):
    """Determine slide type from its elements."""
    elements = slide_info["elements"]

    has_title = any(
        e.get("role") == "title" for e in elements if e["type"] == "text"
    )
    has_body = any(
        e.get("role") == "body" for e in elements if e["type"] == "text"
    )
    has_table = any(e["type"] == "table" for e in elements)
    has_image = any(e["type"] == "image" for e in elements)
    has_chart = any(e["type"] == "chart" for e in elements)
    has_group = any(e["type"] == "group" for e in elements)

    image_count = sum(1 for e in elements if e["type"] == "image")
    body_count = sum(
        1 for e in elements if e["type"] == "text" and e.get("role") == "body"
    )

    # Title slide: title only, no substantive body/data
    if has_title and not has_body and not has_table and not has_chart and image_count <= 1:
        return "title_slide"

    if has_table:
        return "table_slide"

    if has_chart:
        return "chart_slide"

    if has_image and image_count >= 1 and body_count <= 1:
        return "figure_slide"

    if has_group and not has_table:
        return "diagram_slide"

    if has_title and has_body:
        return "content_slide"

    return "content_slide"


def parse_presentation(pptx_path, images_dir):
    """Parse a PPTX file and return structured slide data."""
    src = Presentation(pptx_path)
    os.makedirs(images_dir, exist_ok=True)

    slides_data = []
    for slide_idx, slide in enumerate(src.slides):
        slide_info = {
            "index": slide_idx,
            "layout_name": (
                slide.slide_layout.name if slide.slide_layout else "Unknown"
            ),
            "elements": [],
            "has_notes": False,
            "notes_text": "",
        }

        # Extract shapes
        for shape in slide.shapes:
            element = extract_shape(shape, slide_idx, images_dir)
            if element:
                slide_info["elements"].append(element)

        # Extract speaker notes
        try:
            if slide.has_notes_slide:
                notes = slide.notes_slide.notes_text_frame
                if notes and notes.text.strip():
                    slide_info["has_notes"] = True
                    slide_info["notes_text"] = notes.text.strip()
        except Exception:
            pass

        # Classify slide type
        slide_info["slide_type"] = classify_slide(slide_info)

        slides_data.append(slide_info)

    return {
        "source_file": pptx_path,
        "slide_count": len(slides_data),
        "slide_width": src.slide_width,
        "slide_height": src.slide_height,
        "slides": slides_data,
    }


def main():
    parser = argparse.ArgumentParser(description="Parse a PPTX into structured JSON")
    parser.add_argument("pptx_path", help="Path to source .pptx file")
    parser.add_argument(
        "--output", "-o", default=None, help="Output JSON path (default: stdout)"
    )
    parser.add_argument(
        "--images-dir",
        default="/home/claude/source_images",
        help="Directory to save extracted images",
    )
    args = parser.parse_args()

    if not os.path.exists(args.pptx_path):
        print(f"Error: File not found: {args.pptx_path}", file=sys.stderr)
        sys.exit(1)

    result = parse_presentation(args.pptx_path, args.images_dir)

    # Print summary to stderr
    print(f"Parsed {result['slide_count']} slides:", file=sys.stderr)
    for s in result["slides"]:
        types = [e["type"] for e in s["elements"]]
        title = next(
            (
                e["full_text"][:60]
                for e in s["elements"]
                if e["type"] == "text" and e.get("role") == "title"
            ),
            "(no title)",
        )
        print(
            f"  Slide {s['index']}: {s['slide_type']:15s} | {title} | elements: {types}",
            file=sys.stderr,
        )

    # Output JSON
    output_json = json.dumps(result, indent=2, default=str)
    if args.output:
        with open(args.output, "w") as f:
            f.write(output_json)
        print(f"Saved to {args.output}", file=sys.stderr)
    else:
        print(output_json)


if __name__ == "__main__":
    main()
