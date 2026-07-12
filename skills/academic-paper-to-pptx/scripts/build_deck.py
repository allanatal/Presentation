#!/usr/bin/env python3
"""
build_deck.py — build a Moffitt academic deck from a compact deck spec + claims JSON.

The model authors a deck_spec.json (compact, typed slides that reference claim IDs from
the claims file); this committed builder consumes spec + claims and emits an editable
Moffitt PPTX. Numbers are never hand-authored in python-pptx: prose uses {{claim-id}}
tokens substituted from claims.json (extracted by the strong model), and tables/captions
inline values but declare their claim IDs. Either way the builder records which slide each
claim lands on and writes a resolved claims file (auto-populated `slide`) for qa_crosscheck.

Font sizes are baked in and always >= the style-spec floors — the model cannot set a
sub-floor font because the deck spec has no font field.

Usage:
  python build_deck.py deck_spec.json --claims claims.json --out DECK.pptx \
      [--resolved-claims DECK_claims_resolved.json] [--images-base DIR] [--skill-root DIR] \
      [--allow-unreferenced]

Exit 0 on success; non-zero on spec/claim errors (unknown claim id, unreferenced claim,
unknown slide type). qa_crosscheck.py remains the content gate.
"""
import argparse
import copy
import json
import os
import re
import shutil
import sys

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR

try:
    from PIL import Image
except ImportError:
    Image = None

# ── Colors (style-spec.md) ──
TITLE_BLUE = RGBColor(0x00, 0x33, 0x66)
RED = RGBColor(0xFF, 0x00, 0x00)
BODY = RGBColor(0x33, 0x33, 0x33)
GRAY = RGBColor(0x99, 0x99, 0x99)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ARM_TEAL = RGBColor(0x2E, 0x7D, 0x7D)
ARM_BLUE = RGBColor(0x44, 0x72, 0xC4)
CONTROL_GRAY = RGBColor(0x80, 0x80, 0x80)
NAVY = RGBColor(0x1B, 0x2A, 0x4A)
LIGHT_BG = RGBColor(0xF0, 0xF4, 0xF8)
BORDER = RGBColor(0xCC, 0xCC, 0xCC)
CIRCLE_BLUE = RGBColor(0x4A, 0x6A, 0x8A)
CONN_GRAY = RGBColor(0x88, 0x88, 0x88)
FONT = "Arial"
SLIDE_W = 12192000  # EMU

PALETTE = {
    "teal": ARM_TEAL, "blue": ARM_BLUE, "gray": CONTROL_GRAY, "control": CONTROL_GRAY,
    "navy": NAVY, "lightbg": LIGHT_BG, "titleblue": TITLE_BLUE, "body": BODY,
    "white": WHITE, "red": RED, "border": BORDER,
}


def col(name):
    """Map a palette name (or pass through an RGBColor) to an RGBColor."""
    if isinstance(name, RGBColor):
        return name
    return PALETTE.get(name, ARM_TEAL)


# ── Font-size floors (pt) — hard minimums from style-spec.md ──
F_BODY, F_SUB, F_TABLE, F_DIAG, F_KEYMSG, F_CITE, F_BADGE = 18, 16, 14, 12, 14, 10, 11


def pt(size, floor=None):
    """Point size clamped up to `floor` (hard minimum). The model never picks sizes;
    renderer literals are fixed and already >= floor. This makes the floor structural."""
    if floor is not None and size < floor:
        size = floor
    return Pt(size)


CLAIM_TOKEN_RE = re.compile(r"\{\{([A-Za-z0-9_\-]+)\}\}")


class Claims:
    """Claims file wrapper: value lookup, {{id}} substitution, slide bookkeeping."""

    def __init__(self, data):
        self.data = data
        self.by_id = {c["id"]: c for c in data.get("claims", [])}
        self.referenced = {}  # id -> first slide number it was placed on

    def value(self, cid):
        if cid not in self.by_id:
            raise KeyError(f"unknown claim id: {cid}")
        return self.by_id[cid]["value"]

    def assign(self, cid, slide_no):
        if cid not in self.by_id:
            raise KeyError(f"unknown claim id: {cid}")
        self.referenced.setdefault(cid, slide_no)

    def substitute(self, text, slide_no):
        """Replace every {{id}} in text with the claim value; assign id -> slide_no."""
        def repl(m):
            cid = m.group(1)
            self.assign(cid, slide_no)
            return self.value(cid)
        return CLAIM_TOKEN_RE.sub(repl, text)

    def resolved_json(self):
        """Deep copy of the claims file with `slide` populated from bookkeeping."""
        out = copy.deepcopy(self.data)
        for c in out.get("claims", []):
            c["slide"] = self.referenced.get(c["id"])
        return out

    def unreferenced(self):
        return [cid for cid in self.by_id if cid not in self.referenced]
