#!/usr/bin/env python3
"""
Extract figures from a research paper PDF.

Usage:
    python extract_figures.py paper.pdf /home/claude/figures/

Extracts embedded raster images and filters out small logos/icons.
Real figures (KM curves, forest plots, bar charts) are typically >50KB
and >400px in both dimensions.
"""

import sys
import os
import subprocess
import json
from pathlib import Path


def extract_raster_images(pdf_path, output_dir):
    """Extract embedded raster images using pdfimages."""
    os.makedirs(output_dir, exist_ok=True)
    prefix = os.path.join(output_dir, "img")

    # Extract as PNG for consistent format
    subprocess.run(
        ["pdfimages", "-png", pdf_path, prefix],
        check=True, capture_output=True
    )

    # List what was extracted with metadata
    result = subprocess.run(
        ["pdfimages", "-list", pdf_path],
        capture_output=True, text=True
    )

    # Find all extracted image files
    extracted = sorted(Path(output_dir).glob("img-*.png"))
    return extracted, result.stdout


def filter_figures(image_paths, min_width=400, min_height=400, min_bytes=50000):
    """
    Filter out small images (logos, icons, decorative elements).
    Real figures are typically >400px and >50KB.
    """
    figures = []
    discarded = []

    for img_path in image_paths:
        file_size = img_path.stat().st_size

        # Get image dimensions
        try:
            result = subprocess.run(
                ["identify", "-format", "%wx%h", str(img_path)],
                capture_output=True, text=True, check=True
            )
            dims = result.stdout.strip()
            if "x" in dims:
                w, h = map(int, dims.split("x"))
            else:
                w, h = 0, 0
        except (subprocess.CalledProcessError, ValueError):
            w, h = 0, 0

        info = {
            "path": str(img_path),
            "filename": img_path.name,
            "width": w,
            "height": h,
            "size_bytes": file_size,
            "size_kb": round(file_size / 1024, 1),
        }

        if w >= min_width and h >= min_height and file_size >= min_bytes:
            info["status"] = "figure"
            figures.append(info)
        else:
            info["status"] = "discarded"
            discarded.append(info)

    return figures, discarded


def main():
    if len(sys.argv) < 3:
        print("Usage: python extract_figures.py <paper.pdf> <output_dir>")
        sys.exit(1)

    pdf_path = sys.argv[1]
    output_dir = sys.argv[2]

    if not os.path.exists(pdf_path):
        print(f"Error: PDF not found: {pdf_path}")
        sys.exit(1)

    print(f"Extracting images from: {pdf_path}")
    print(f"Output directory: {output_dir}")

    # Step 1: Extract all raster images
    extracted, listing = extract_raster_images(pdf_path, output_dir)
    print(f"\nExtracted {len(extracted)} raster images")

    if listing:
        print("\n--- pdfimages listing ---")
        # Show first 30 lines
        for line in listing.strip().split("\n")[:30]:
            print(line)
        if listing.count("\n") > 30:
            print(f"  ... ({listing.count(chr(10)) - 30} more rows)")

    # Step 2: Filter
    figures, discarded = filter_figures(extracted)

    print(f"\n--- Results ---")
    print(f"Figures (likely real content): {len(figures)}")
    for f in figures:
        print(f"  ✓ {f['filename']}  {f['width']}x{f['height']}  ({f['size_kb']} KB)")

    print(f"\nDiscarded (logos/icons/small): {len(discarded)}")
    for d in discarded:
        print(f"  ✗ {d['filename']}  {d['width']}x{d['height']}  ({d['size_kb']} KB)")

    # Save manifest
    manifest = {
        "pdf": pdf_path,
        "figures": figures,
        "discarded": discarded,
    }
    manifest_path = os.path.join(output_dir, "manifest.json")
    with open(manifest_path, "w") as f:
        json.dump(manifest, f, indent=2)
    print(f"\nManifest saved to: {manifest_path}")

    if not figures:
        print("\n⚠️  No large figures found via raster extraction.")
        print("    The paper may use vector-drawn figures (matplotlib, R, Excel).")
        print("    Use pdftoppm to rasterize specific pages instead:")
        print("    pdftoppm -png -r 300 -f <PAGE> -l <PAGE> paper.pdf /home/claude/figures/vector_fig")


if __name__ == "__main__":
    main()
