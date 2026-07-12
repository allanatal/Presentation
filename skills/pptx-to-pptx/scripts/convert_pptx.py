#!/usr/bin/env python3
"""
End-to-end PPTX-to-PPTX conversion pipeline.

Chains: parse source → rasterize chart/diagram slides → build output.

Usage:
    python convert_pptx.py source.pptx [--output output.pptx] \
        [--study-name "Study Name"] [--citation "Author et al. 2026"]
"""

import argparse
import json
import os
import shutil
import subprocess
import sys

# Add scripts directory to path
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from parse_pptx import parse_presentation
from build_from_parsed import build_presentation


def find_soffice():
    """Locate a LibreOffice conversion entry point.

    Returns a command-list prefix, or None if LibreOffice is unavailable.
    Prefers the Claude desktop-app sandbox helper when present, otherwise
    the local soffice binary (PATH, then the macOS app bundle).
    """
    sandbox_helper = "/mnt/skills/public/pptx/scripts/office/soffice.py"
    if os.path.exists(sandbox_helper):
        return ["python3", sandbox_helper]
    soffice = shutil.which("soffice")
    if soffice:
        return [soffice]
    mac_soffice = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if os.path.exists(mac_soffice):
        return [mac_soffice]
    return None


SOFFICE_CMD = find_soffice()


def rasterize_slides(source_pptx, parsed_data, images_dir):
    """Rasterize slides that contain charts or complex diagrams."""
    slides_to_rasterize = []
    for s in parsed_data.get("slides", []):
        if s["slide_type"] in ("chart_slide", "diagram_slide"):
            slides_to_rasterize.append(s["index"])

    if not slides_to_rasterize:
        print("No chart/diagram slides to rasterize.", file=sys.stderr)
        return

    print(f"Rasterizing {len(slides_to_rasterize)} slides with charts/diagrams...", file=sys.stderr)

    if SOFFICE_CMD is None:
        print("WARNING: LibreOffice (soffice) not found — cannot rasterize chart/diagram "
              "slides. Install it (e.g. `brew install --cask libreoffice`) and re-run.",
              file=sys.stderr)
        return

    # Convert source to PDF
    pdf_path = os.path.join(images_dir, "source.pdf")
    result = subprocess.run(
        SOFFICE_CMD + ["--headless", "--convert-to", "pdf", source_pptx],
        capture_output=True, text=True, cwd=images_dir,
    )

    # soffice outputs to same directory as input; move if needed
    source_basename = os.path.splitext(os.path.basename(source_pptx))[0]
    generated_pdf = os.path.join(os.path.dirname(source_pptx), f"{source_basename}.pdf")
    if not os.path.exists(generated_pdf):
        # Try current dir
        generated_pdf = f"{source_basename}.pdf"

    if os.path.exists(generated_pdf) and generated_pdf != pdf_path:
        os.rename(generated_pdf, pdf_path)
    elif not os.path.exists(pdf_path):
        # Try the images_dir
        alt = os.path.join(images_dir, f"{source_basename}.pdf")
        if os.path.exists(alt):
            os.rename(alt, pdf_path)

    if not os.path.exists(pdf_path):
        print(f"WARNING: Could not generate PDF for rasterization. Tried: {generated_pdf}", file=sys.stderr)
        return

    # Rasterize each chart/diagram slide
    for slide_idx in slides_to_rasterize:
        page_num = slide_idx + 1  # pdftoppm is 1-indexed
        prefix = os.path.join(images_dir, f"slide_raster")
        subprocess.run(
            ["pdftoppm", "-jpeg", "-r", "300", "-f", str(page_num), "-l", str(page_num),
             pdf_path, prefix],
            capture_output=True,
        )
        # pdftoppm outputs slide_raster-NN.jpg
        expected = os.path.join(images_dir, f"slide_raster-{page_num:02d}.jpg")
        if os.path.exists(expected):
            print(f"  ✓ Rasterized slide {slide_idx} → {expected}", file=sys.stderr)
        else:
            print(f"  ✗ Rasterization failed for slide {slide_idx}", file=sys.stderr)


def main():
    parser = argparse.ArgumentParser(description="Convert a PPTX to Moffitt template style")
    parser.add_argument("source", help="Path to source .pptx file")
    parser.add_argument("--output", "-o", default="output.pptx", help="Output PPTX path")
    parser.add_argument("--study-name", default="", help="Study name for badge")
    parser.add_argument("--citation", default="", help="Citation text for footer")
    parser.add_argument("--images-dir", default="source_images", help="Working directory for images")
    parser.add_argument("--parsed-json", default=None, help="Save/load parsed JSON (skip re-parsing)")
    args = parser.parse_args()

    if not os.path.exists(args.source):
        print(f"Error: File not found: {args.source}", file=sys.stderr)
        sys.exit(1)

    os.makedirs(args.images_dir, exist_ok=True)
    json_path = args.parsed_json or os.path.join(args.images_dir, "parsed.json")

    # Phase 1: Parse
    print("=" * 60, file=sys.stderr)
    print("Phase 1: Parsing source presentation", file=sys.stderr)
    print("=" * 60, file=sys.stderr)
    parsed_data = parse_presentation(args.source, args.images_dir)
    with open(json_path, "w") as f:
        json.dump(parsed_data, f, indent=2, default=str)
    print(f"Parsed JSON saved to {json_path}", file=sys.stderr)

    # Phase 1.5: Rasterize charts/diagrams
    print("\n" + "=" * 60, file=sys.stderr)
    print("Phase 1.5: Rasterizing charts and diagrams", file=sys.stderr)
    print("=" * 60, file=sys.stderr)
    rasterize_slides(args.source, parsed_data, args.images_dir)

    # Phase 2: Classification already done during parse

    # Phase 3: Build
    print("\n" + "=" * 60, file=sys.stderr)
    print("Phase 3: Building output presentation", file=sys.stderr)
    print("=" * 60, file=sys.stderr)
    build_presentation(
        parsed_data,
        args.output,
        study_name=args.study_name,
        citation=args.citation,
        images_dir=args.images_dir,
    )

    # Summary
    print("\n" + "=" * 60, file=sys.stderr)
    print("Done!", file=sys.stderr)
    print(f"  Source: {args.source} ({parsed_data['slide_count']} slides)", file=sys.stderr)
    print(f"  Output: {args.output}", file=sys.stderr)
    print(f"  Images: {args.images_dir}", file=sys.stderr)
    print("=" * 60, file=sys.stderr)
    print(f"\nQA: Run the following to visually inspect:", file=sys.stderr)
    soffice_hint = " ".join(SOFFICE_CMD) if SOFFICE_CMD else "soffice"
    print(f"  {soffice_hint} --headless --convert-to pdf {args.output}", file=sys.stderr)
    print(f"  pdftoppm -jpeg -r 150 {os.path.splitext(args.output)[0]}.pdf slide", file=sys.stderr)


if __name__ == "__main__":
    main()
