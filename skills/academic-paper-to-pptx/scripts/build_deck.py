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


def render_study_design(ctx, s, n):
    slide = ctx.prs.slides.add_slide(ctx.layouts["Title Only"])
    slide.placeholders[0].text = s.get("title", "Study Design")
    desc = slide.shapes.add_textbox(Emu(838200), Emu(1420000), Emu(10515600), Emu(600000))
    desc.text_frame.word_wrap = True
    p = desc.text_frame.paragraphs[0]
    p.text = s["description"]
    p.font.size = pt(14, F_DIAG); p.font.italic = True
    p.font.color.rgb = TITLE_BLUE; p.font.name = FONT

    content_y = Emu(2250000)
    # Eligibility box (left)
    elig = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(280000), content_y,
                                  Emu(3350000), Emu(3550000))
    elig.fill.solid(); elig.fill.fore_color.rgb = LIGHT_BG
    elig.line.color.rgb = BORDER; elig.line.width = Pt(0.75)
    tf = elig.text_frame; tf.word_wrap = True
    ph = tf.paragraphs[0]; ph.text = s.get("eligibility_header", "Key Eligibility")
    ph.font.size = Pt(13); ph.font.bold = True; ph.font.color.rgb = BODY; ph.font.name = FONT
    for crit in s.get("eligibility", []):
        pe = tf.add_paragraph(); pe.text = "• " + crit
        pe.font.size = pt(12, F_DIAG); pe.font.color.rgb = BODY; pe.font.name = FONT
    for extra in s.get("eligibility_notes", []):
        pn = tf.add_paragraph(); pn.text = extra
        pn.font.size = pt(12, F_DIAG); pn.font.italic = True
        pn.font.color.rgb = GRAY; pn.font.name = FONT

    # Randomization circle
    circ_x, circ_y, circ_d = Emu(3850000), content_y + Emu(1150000), Emu(950000)
    circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, circ_x, circ_y, circ_d, circ_d)
    circ.fill.solid(); circ.fill.fore_color.rgb = CIRCLE_BLUE; circ.line.fill.background()
    ctf = circ.text_frame
    ctf.paragraphs[0].text = "R"
    ctf.paragraphs[0].font.size = Pt(18); ctf.paragraphs[0].font.bold = True
    ctf.paragraphs[0].font.color.rgb = WHITE; ctf.paragraphs[0].alignment = PP_ALIGN.CENTER
    ctf.paragraphs[0].font.name = FONT
    pr = ctf.add_paragraph(); pr.text = s.get("randomization", "")
    pr.font.size = Pt(13); pr.font.bold = True; pr.font.color.rgb = WHITE
    pr.alignment = PP_ALIGN.CENTER; pr.font.name = FONT
    nb = slide.shapes.add_textbox(circ_x - Emu(150000), circ_y + circ_d + Emu(40000),
                                  circ_d + Emu(300000), Emu(300000))
    pn = nb.text_frame.paragraphs[0]; pn.text = s.get("n", "")
    pn.font.size = pt(13, F_DIAG); pn.font.bold = True; pn.font.color.rgb = BODY
    pn.alignment = PP_ALIGN.CENTER; pn.font.name = FONT

    # Arms
    arms = s["arms"]
    arm_x, arm_w = Emu(5150000), Emu(3450000)
    span = Emu(3300000)
    n_arms = len(arms)
    arm_h = min(Emu(1400000), span // n_arms - Emu(80000)) if n_arms else Emu(1400000)
    step = span // n_arms if n_arms else span
    for i, arm in enumerate(arms):
        ay = content_y + Emu(250000) + step * i
        a = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, arm_x, ay, arm_w, arm_h)
        a.fill.solid(); a.fill.fore_color.rgb = col(arm.get("color", "teal"))
        a.line.fill.background()
        atf = a.text_frame; atf.word_wrap = True
        pa = atf.paragraphs[0]; pa.text = arm["name"]
        pa.font.size = Pt(14); pa.font.bold = True; pa.font.color.rgb = WHITE; pa.font.name = FONT
        if arm.get("detail"):
            pd = atf.add_paragraph(); pd.text = arm["detail"]
            pd.font.size = pt(12, F_DIAG); pd.font.color.rgb = WHITE; pd.font.name = FONT

    # Endpoints box (right)
    ep = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(8850000),
                                content_y + Emu(250000), Emu(2950000), Emu(3050000))
    ep.fill.solid(); ep.fill.fore_color.rgb = LIGHT_BG
    ep.line.color.rgb = BORDER; ep.line.width = Pt(0.75)
    etf = ep.text_frame; etf.word_wrap = True
    pe = etf.paragraphs[0]
    pe.text = "Primary Endpoint" + ("s" if len(s.get("primary_endpoints", [])) > 1 else "")
    pe.font.size = Pt(13); pe.font.bold = True; pe.font.color.rgb = BODY; pe.font.name = FONT
    for e in s.get("primary_endpoints", []):
        p = etf.add_paragraph(); p.text = "• " + e
        p.font.size = pt(12, F_DIAG); p.font.color.rgb = BODY; p.font.name = FONT
    ps = etf.add_paragraph(); ps.text = "Secondary Endpoints"
    ps.font.size = Pt(13); ps.font.bold = True; ps.font.color.rgb = TITLE_BLUE; ps.font.name = FONT
    for e in s.get("secondary_endpoints", []):
        p = etf.add_paragraph(); p.text = "• " + e
        p.font.size = pt(12, F_DIAG); p.font.color.rgb = BODY; p.font.name = FONT
    if s.get("registration"):
        preg = etf.add_paragraph(); preg.text = "Registration: " + s["registration"]
        preg.font.size = pt(12, F_DIAG); preg.font.color.rgb = GRAY; preg.font.name = FONT

    # arrows
    def arrow(x1, y1, x2, y2):
        cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(x1), Emu(y1), Emu(x2), Emu(y2))
        cn.line.color.rgb = CONN_GRAY; cn.line.width = Pt(1.5)
    arrow(3630000, content_y + Emu(1600000), 3850000, content_y + Emu(1600000))
    for i in range(n_arms):
        ay = content_y + Emu(250000) + step * i + arm_h // 2
        arrow(4800000, content_y + Emu(1600000), 5150000, ay)
        arrow(8600000, ay, 8850000, content_y + Emu(1600000))

    if s.get("footer"):
        dur = slide.shapes.add_textbox(Emu(280000), content_y + Emu(3650000),
                                       Emu(11600000), Emu(400000))
        pd = dur.text_frame.paragraphs[0]; pd.text = s["footer"]
        pd.font.size = pt(12, F_DIAG); pd.font.color.rgb = BODY
        pd.alignment = PP_ALIGN.CENTER; pd.font.name = FONT
    add_badge_and_citation(ctx, slide)
    return slide


def render_consort(ctx, s, n):
    slide = ctx.prs.slides.add_slide(ctx.layouts["Title Only"])
    slide.placeholders[0].text = s.get("title", "CONSORT Flow Diagram")

    def box(x, y, w, h, lines, fill=WHITE, edge=BORDER):
        b = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y), Emu(w), Emu(h))
        b.fill.solid(); b.fill.fore_color.rgb = fill
        b.line.color.rgb = edge; b.line.width = Pt(0.75)
        tf = b.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_top = Emu(30000); tf.margin_bottom = Emu(30000)
        for i, (txt, bold, colr, sz) in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = txt; p.font.size = pt(sz, F_DIAG); p.font.bold = bold
            p.font.color.rgb = colr; p.font.name = FONT
        return b

    def conn(x1, y1, x2, y2):
        cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(x1), Emu(y1), Emu(x2), Emu(y2))
        cn.line.color.rgb = CONN_GRAY; cn.line.width = Pt(1.5)

    arms = s["arms"]
    centers = [3350000, 8850000] if len(arms) == 2 else \
        [int(2200000 + (7800000 / max(len(arms) - 1, 1)) * i) for i in range(len(arms))]
    bw = 3450000
    top_c = sum(centers) // len(centers)

    box(top_c - 1900000, 1500000, 3800000, 470000, [(s["assessed"], True, BODY, 13)])
    ex = s.get("excluded", [])
    if ex:
        box(top_c + 2100000, 1470000, 3350000, 560000,
            [(ex[0], True, BODY, 12)] + [(t, False, BODY, 12) for t in ex[1:]])
        conn(top_c + 1900000, 1735000, top_c + 2100000, 1735000)
    box(top_c - 1900000, 2180000, 3800000, 470000, [(s["randomized"], True, BODY, 13)])
    conn(top_c, 1970000, top_c, 2180000)

    for arm, cx in zip(arms, centers):
        x = cx - bw // 2
        alloc = arm["allocated"]
        box(x, 2870000, bw, 560000,
            [(alloc[0], True, WHITE, 12)] + [(t, False, WHITE, 12) for t in alloc[1:]],
            fill=col(arm.get("color", "teal")), edge=col(arm.get("color", "teal")))
        conn(top_c, 2650000, cx, 2870000)
        nr = arm.get("not_received", [])
        if nr:
            box(x + 300000, 3560000, bw - 300000, 470000,
                [(nr[0], False, BODY, 12)] + [(t, False, GRAY, 12) for t in nr[1:]])
            conn(cx, 3430000, cx, 3560000)
        box(x, 4130000, bw, 470000, [(arm["received"], True, BODY, 12)], fill=LIGHT_BG)
        conn(cx, 4030000, cx, 4130000)
        oc = arm.get("outcomes", [])
        if oc:
            box(x, 4720000, bw, 1500000, [(t, False, BODY, 12) for t in oc], fill=WHITE)
            conn(cx, 4600000, cx, 4720000)
    add_badge_and_citation(ctx, slide)
    return slide
