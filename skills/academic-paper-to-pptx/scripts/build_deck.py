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


def substitute_in_place(node, claims, slide_no):
    """Recursively replace {{id}} tokens in every string within a dict/list tree.
    Assigns each referenced claim to slide_no. Mutates `node` in place."""
    if isinstance(node, dict):
        for k, v in node.items():
            if isinstance(v, str):
                node[k] = claims.substitute(v, slide_no)
            else:
                substitute_in_place(v, claims, slide_no)
    elif isinstance(node, list):
        for i, v in enumerate(node):
            if isinstance(v, str):
                node[i] = claims.substitute(v, slide_no)
            else:
                substitute_in_place(v, claims, slide_no)


def assign_slide_claims(sdata, claims, slide_no):
    """Assign every id in a slide's explicit `claims` list to slide_no."""
    for cid in sdata.get("claims", []):
        claims.assign(cid, slide_no)


# ═══════════════════════════════════════════════════
#  Build context + shared helpers
# ═══════════════════════════════════════════════════

class Ctx:
    def __init__(self, prs, layouts, study, citation, skill_root, images_base):
        self.prs = prs
        self.layouts = layouts
        self.study = study
        self.citation = citation
        self.skill_root = skill_root
        self.images_base = images_base


def add_badge_and_citation(ctx, slide):
    badge = slide.shapes.add_textbox(Emu(9984826), Emu(157655), Emu(1671098), Emu(369332))
    p = badge.text_frame.paragraphs[0]
    p.text = ctx.study
    p.font.size = pt(11, F_BADGE); p.font.bold = True
    p.font.color.rgb = TITLE_BLUE; p.font.name = FONT
    p.alignment = PP_ALIGN.RIGHT
    cite = slide.shapes.add_textbox(Emu(9987101), Emu(6463328), Emu(2204899), Emu(307777))
    p = cite.text_frame.paragraphs[0]
    p.text = ctx.citation
    p.font.size = pt(10, F_CITE); p.font.color.rgb = GRAY; p.font.name = FONT
    p.alignment = PP_ALIGN.RIGHT


def key_message(slide, text, y=Emu(1400000), size=14):
    msg = slide.shapes.add_textbox(Emu(838200), y, Emu(10515600), Emu(500000))
    msg.text_frame.word_wrap = True
    p = msg.text_frame.paragraphs[0]
    p.text = text
    p.font.size = pt(size, F_KEYMSG); p.font.italic = True
    p.font.color.rgb = TITLE_BLUE; p.font.name = FONT
    return msg


# ═══════════════════════════════════════════════════
#  Renderers — one per slide type (render_x(ctx, s, n))
# ═══════════════════════════════════════════════════

def render_title(ctx, s, n):
    slide = ctx.prs.slides.add_slide(ctx.layouts["Title Slide"])
    t = slide.placeholders[0]
    t.text = s["title"]
    for para in t.text_frame.paragraphs:
        para.font.size = Pt(s.get("title_size", 26)); para.font.bold = True
        para.font.color.rgb = RED; para.font.name = FONT
    if s.get("subtitle"):
        sub = slide.placeholders[1]
        sub.text = s["subtitle"]
        for para in sub.text_frame.paragraphs:
            para.font.size = Pt(20); para.font.bold = True
            para.font.color.rgb = BODY; para.font.name = FONT
    lines = [ln for ln in [s.get("authors"), s.get("affiliation"), "", s.get("date")]
             if ln is not None]
    if lines:
        box = slide.shapes.add_textbox(Emu(1154187), Emu(4928445), Emu(9340438), Emu(1200329))
        box.text_frame.word_wrap = True
        for i, line in enumerate(lines):
            p = box.text_frame.paragraphs[0] if i == 0 else box.text_frame.add_paragraph()
            p.text = line
            p.font.size = Pt(14) if line != s.get("date") else Pt(13)
            p.font.color.rgb = RGBColor(0x66, 0x66, 0x66); p.font.name = FONT
            p.alignment = PP_ALIGN.CENTER
    logo = os.path.join(ctx.skill_root, "references", "Picture_3.x-wmf")
    if os.path.exists(logo):
        try:
            slide.shapes.add_picture(logo, Emu(4790362), Emu(798667), Emu(2266391), Emu(578214))
        except Exception as e:
            print("logo skip:", e, file=sys.stderr)
    return slide


def render_bullets(ctx, s, n):
    slide = ctx.prs.slides.add_slide(ctx.layouts["Title and Content"])
    slide.placeholders[0].text = s["title"]
    if s.get("key_message"):
        key_message(slide, s["key_message"])
    tf = slide.placeholders[1].text_frame
    tf.clear()
    first = True
    for b in s["bullets"]:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.text = b["text"]; p.level = 0
        p.font.name = FONT; p.font.size = pt(18, F_BODY)
        if b.get("bold"):
            p.font.bold = True
        for sub in b.get("sub", []):
            ps = tf.add_paragraph()
            ps.text = sub; ps.level = 1
            ps.font.name = FONT; ps.font.size = pt(16, F_SUB)
    add_badge_and_citation(ctx, slide)
    return slide


def render_figure(ctx, s, n):
    slide = ctx.prs.slides.add_slide(ctx.layouts["Title Only"])
    slide.placeholders[0].text = s["title"]
    if s.get("key_message"):
        key_message(slide, s["key_message"])
    fig_top = s.get("fig_top", 1950000)
    max_h = s.get("max_h", 3400000)
    image_path = s["image"]
    if not os.path.isabs(image_path):
        image_path = os.path.join(ctx.images_base, image_path)
    if Image and os.path.exists(image_path):
        img = Image.open(image_path); iw, ih = img.size; ratio = iw / ih
        max_w = Emu(10500000)
        w = max_w; h = int(w / ratio)
        if h > max_h:
            h = max_h; w = int(h * ratio)
        fig_x = Emu(838200) + (max_w - w) // 2
        slide.shapes.add_picture(image_path, fig_x, Emu(fig_top), w, h)
        cap_y = fig_top + h + 120000
    else:
        print(f"WARNING: figure image not found: {image_path}", file=sys.stderr)
        cap_y = fig_top + 200000
    cap = None
    for i, line in enumerate(s.get("caption", []) or []):
        if i == 0:
            cap = slide.shapes.add_textbox(Emu(838200), Emu(cap_y), Emu(10515600), Emu(900000))
            cap.text_frame.word_wrap = True
            p = cap.text_frame.paragraphs[0]
        else:
            p = cap.text_frame.add_paragraph()
        p.text = line; p.font.size = Pt(14); p.font.name = FONT; p.font.color.rgb = BODY
    add_badge_and_citation(ctx, slide)
    return slide


def render_table(ctx, s, n):
    slide = ctx.prs.slides.add_slide(ctx.layouts["Title Only"])
    slide.placeholders[0].text = s["title"]
    font = s.get("font", 15)
    row_h = s.get("row_h", 360000)
    top = s.get("top", 1900000)
    if s.get("key_message"):
        key_message(slide, s["key_message"])
        top = max(top, 1980000)
    arms = s["arms"]
    rows = s["rows"]
    n_cols = 1 + len(arms)
    n_rows = 1 + len(rows)
    label_w = Inches(4.6)
    data_w = int((Inches(11.0) - label_w) / max(len(arms), 1))
    col_w = [label_w] + [data_w] * len(arms)
    total_w = sum(col_w)
    table_x = (Emu(SLIDE_W) - total_w) // 2
    shp = slide.shapes.add_table(n_rows, n_cols, table_x, Emu(top), total_w,
                                 Emu(n_rows * row_h))
    tbl = shp.table
    tbl.first_row = False; tbl.horz_banding = False
    for i, w in enumerate(col_w):
        tbl.columns[i].width = int(w)
    # header
    c0 = tbl.cell(0, 0); c0.text = s.get("label_header", "Characteristic")
    for p in c0.text_frame.paragraphs:
        p.font.size = pt(font, F_TABLE); p.font.bold = True
        p.font.color.rgb = WHITE; p.font.name = FONT
    c0.fill.solid(); c0.fill.fore_color.rgb = TITLE_BLUE
    for j, arm in enumerate(arms):
        c = tbl.cell(0, j + 1); c.text = arm["label"]
        c.fill.solid(); c.fill.fore_color.rgb = col(arm.get("color", "teal"))
        for p in c.text_frame.paragraphs:
            p.font.size = pt(font, F_TABLE); p.font.bold = True
            p.font.color.rgb = WHITE; p.font.name = FONT; p.alignment = PP_ALIGN.CENTER
    for r in tbl.rows:
        for c in r.cells:
            c.margin_top = Emu(16000); c.margin_bottom = Emu(16000)
    for i, row in enumerate(rows):
        is_hdr = row.get("hdr", False)
        c = tbl.cell(i + 1, 0); c.text = row["label"]
        for p in c.text_frame.paragraphs:
            p.font.size = pt(font, F_TABLE); p.font.name = FONT
            p.font.bold = is_hdr or row.get("bold", False)
            p.font.color.rgb = TITLE_BLUE if is_hdr else BODY
        vals = row.get("vals", [""] * len(arms))
        for j, v in enumerate(vals):
            c = tbl.cell(i + 1, j + 1); c.text = v
            for p in c.text_frame.paragraphs:
                p.font.size = pt(font, F_TABLE); p.font.name = FONT
                p.font.color.rgb = BODY; p.alignment = PP_ALIGN.CENTER
        if is_hdr:
            for j in range(n_cols):
                cc = tbl.cell(i + 1, j)
                cc.fill.solid(); cc.fill.fore_color.rgb = LIGHT_BG
    for r in tbl.rows:
        r.height = Emu(row_h)
    if s.get("note"):
        note_y = top + n_rows * row_h + 120000
        nb = slide.shapes.add_textbox(Emu(838200), Emu(note_y), Emu(10515600), Emu(750000))
        nb.text_frame.word_wrap = True
        p = nb.text_frame.paragraphs[0]
        p.text = s["note"]
        p.font.size = Pt(14); p.font.name = FONT; p.font.italic = True; p.font.color.rgb = BODY
    add_badge_and_citation(ctx, slide)
    return slide
