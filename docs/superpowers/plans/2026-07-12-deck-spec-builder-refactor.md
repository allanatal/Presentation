# Deck-Spec → Fixed Builder Refactor — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Stop having the model hand-author ~28 KB of python-pptx per deck. The model emits a compact typed **deck-spec JSON** (ordered slides referencing claim IDs); a single committed `scripts/build_deck.py` consumes spec + claims JSON → editable Moffitt PPTX.

**Architecture:** `build_deck.py` is the fixed builder. It copies the Moffitt `template.pptx`, dispatches each spec slide to a typed renderer (title/bullets/table/figure/study_design/consort), substitutes `{{claim-id}}` tokens from `claims.json`, records which slide each claim landed on, writes a **resolved claims file** (auto-populated `slide` field), and saves the deck. `qa_crosscheck.py` (unchanged) remains the gate. Numbers are never typed in python-pptx; font sizes live only in the committed builder and are always ≥ the style-spec floors, so the model cannot emit a sub-floor font — the spec has no font field.

**Tech Stack:** Python 3, `python-pptx`, `Pillow`, `pytest` (dev only). Mirrors the existing `pptx-to-pptx/scripts/build_from_parsed.py` (structured-JSON → committed builder → deck → qa_crosscheck).

**Integrity guardrails (non-negotiable, from CLAUDE.md + spec.md):**
- Claim extraction (reading the paper, HR/CI/median/Table 1) stays on the strong model — this refactor changes only the mechanical *build*, never extraction.
- `qa_crosscheck.py` gates every build; its claim-token check still catches a mistyped inlined table value (the claim value is ground truth).
- Font floors stay hard: body 18 / sub-bullet 16 / table 14 / diagram 12; citation 10 & badge 11 exempt.
- Repo is the only editable source. After changes under `skills/`, run `tools/sync-skills.sh` and remind Allan to re-upload `dist/*.skill`.

---

## File structure

- **Create** `skills/academic-paper-to-pptx/scripts/build_deck.py` — the committed builder (loader + Claims + renderers + CLI). One responsibility: turn a validated spec + claims into a deck and a resolved claims file.
- **Create** `skills/academic-paper-to-pptx/references/deck-spec.md` — the model-facing deck-spec schema with one worked example per slide type.
- **Create** `skills/academic-paper-to-pptx/scripts/tests/test_build_deck.py` — pytest unit + integration tests (dev-only; not shipped in `.skill` but kept in repo).
- **Create** `skills/academic-paper-to-pptx/scripts/tests/fixtures/mini_spec.json`, `mini_claims.json` — tiny synthetic deck for the self-test (public/synthetic content only).
- **Modify** `skills/academic-paper-to-pptx/SKILL.md` — Phase 3 becomes "author deck-spec JSON → run build_deck.py"; add the integrity note about no-font-field.
- **Modify** `skills/academic-paper-to-pptx/references/slide-builders.md` — replace the big per-deck code blocks with a pointer to `build_deck.py` + the deck-spec schema (kept as the "how the builder renders each type" reference).
- **Create** `build/DRAGON01_deck_spec.json` — the DRAGON-01 deck expressed as a spec (the regression proof).
- **Modify** `spec.md` — mark the Token-Efficiency option 1 as IMPLEMENTED; update status.
- **Modify** `CLAUDE.md` — append a Lessons-Learned entry if anything non-obvious surfaces.

---

### Task 1: Claims resolution core (loader + token substitution)

**Files:**
- Create: `skills/academic-paper-to-pptx/scripts/build_deck.py`
- Test: `skills/academic-paper-to-pptx/scripts/tests/test_build_deck.py`

- [ ] **Step 1: Write failing tests for the `Claims` class**

```python
# tests/test_build_deck.py
import os, sys, json, copy
import pytest
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import build_deck as bd

CLAIMS = {
    "deck": {"title": "T", "source": "p.pdf", "nct": "NCT00000000", "date": "2026-07-12"},
    "claims": [
        {"id": "eff-01", "category": "efficacy", "label": "Median OS IP",
         "value": "19.4 months (95% CI 17.1-22.9)", "slide": None, "source_hint": "Fig 2A"},
        {"id": "eff-03", "category": "efficacy", "label": "OS HR",
         "value": "HR 0.67 (95% CI 0.50-0.90); P = .01", "slide": None, "source_hint": "Fig 2A"},
        {"id": "dem-01", "category": "demographics", "label": "Age",
         "value": "IP 60 (24-70); PS 56 (23-74)", "slide": None, "source_hint": "Table 1"},
    ],
}

def test_value_lookup():
    c = bd.Claims(copy.deepcopy(CLAIMS))
    assert c.value("eff-01") == "19.4 months (95% CI 17.1-22.9)"

def test_unknown_claim_raises():
    c = bd.Claims(copy.deepcopy(CLAIMS))
    with pytest.raises(KeyError):
        c.value("nope-99")

def test_substitute_expands_and_assigns():
    c = bd.Claims(copy.deepcopy(CLAIMS))
    out = c.substitute("Median OS {{eff-01}}; {{eff-03}}", slide_no=10)
    assert out == "Median OS 19.4 months (95% CI 17.1-22.9); HR 0.67 (95% CI 0.50-0.90); P = .01"
    assert c.referenced["eff-01"] == 10
    assert c.referenced["eff-03"] == 10

def test_substitute_unknown_token_raises():
    c = bd.Claims(copy.deepcopy(CLAIMS))
    with pytest.raises(KeyError):
        c.substitute("bad {{nope-99}}", slide_no=1)

def test_explicit_assign():
    c = bd.Claims(copy.deepcopy(CLAIMS))
    c.assign("dem-01", 8)
    assert c.referenced["dem-01"] == 8

def test_first_assignment_wins():
    c = bd.Claims(copy.deepcopy(CLAIMS))
    c.assign("eff-01", 5)
    c.assign("eff-01", 9)
    assert c.referenced["eff-01"] == 5

def test_resolved_json_fills_slide_and_preserves_original():
    data = copy.deepcopy(CLAIMS)
    c = bd.Claims(data)
    c.assign("eff-01", 10); c.assign("dem-01", 8)
    resolved = c.resolved_json()
    by = {x["id"]: x for x in resolved["claims"]}
    assert by["eff-01"]["slide"] == 10
    assert by["dem-01"]["slide"] == 8
    assert by["eff-03"]["slide"] is None
    # original untouched
    assert data["claims"][0]["slide"] is None

def test_unreferenced_lists_unused_ids():
    c = bd.Claims(copy.deepcopy(CLAIMS))
    c.assign("eff-01", 10)
    assert set(c.unreferenced()) == {"eff-03", "dem-01"}
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `cd skills/academic-paper-to-pptx/scripts && python -m pytest tests/test_build_deck.py -k "value or substitute or assign or resolved or unreferenced" -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'build_deck'` (or `AttributeError: module has no attribute 'Claims'`).

- [ ] **Step 3: Create `build_deck.py` header + colors + `Claims`**

```python
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
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `cd skills/academic-paper-to-pptx/scripts && python -m pytest tests/test_build_deck.py -k "value or substitute or assign or resolved or unreferenced" -v`
Expected: PASS (8 tests).

- [ ] **Step 5: Commit**

```bash
rm -f .git/index.lock
git add skills/academic-paper-to-pptx/scripts/build_deck.py skills/academic-paper-to-pptx/scripts/tests/test_build_deck.py
git commit -m "feat(build_deck): claims resolution core (token substitution + slide bookkeeping)"
```

---

### Task 2: Recursive spec substitution + slide-level claim assignment

**Files:**
- Modify: `skills/academic-paper-to-pptx/scripts/build_deck.py`
- Test: `skills/academic-paper-to-pptx/scripts/tests/test_build_deck.py`

- [ ] **Step 1: Write failing tests for `substitute_in_place`**

```python
def test_substitute_in_place_walks_nested_structures():
    c = bd.Claims(copy.deepcopy(CLAIMS))
    sdata = {
        "type": "figure",
        "key_message": "OS {{eff-01}}",
        "caption": ["deaths line", "HR is {{eff-03}}"],
        "rows": [{"label": "Age", "vals": ["{{dem-01}}"]}],
    }
    bd.substitute_in_place(sdata, c, slide_no=10)
    assert sdata["key_message"] == "OS 19.4 months (95% CI 17.1-22.9)"
    assert sdata["caption"][1] == "HR is HR 0.67 (95% CI 0.50-0.90); P = .01"
    assert sdata["rows"][0]["vals"][0] == "IP 60 (24-70); PS 56 (23-74)"
    assert c.referenced == {"eff-01": 10, "eff-03": 10, "dem-01": 10}

def test_slide_level_claims_assigned():
    c = bd.Claims(copy.deepcopy(CLAIMS))
    sdata = {"type": "table", "claims": ["dem-01"], "rows": []}
    bd.assign_slide_claims(sdata, c, slide_no=8)
    assert c.referenced["dem-01"] == 8
```

- [ ] **Step 2: Run to verify they fail**

Run: `cd skills/academic-paper-to-pptx/scripts && python -m pytest tests/test_build_deck.py -k "in_place or slide_level" -v`
Expected: FAIL — `AttributeError: module 'build_deck' has no attribute 'substitute_in_place'`.

- [ ] **Step 3: Add the helpers to `build_deck.py`** (after the `Claims` class)

```python
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
```

- [ ] **Step 4: Run to verify they pass**

Run: `cd skills/academic-paper-to-pptx/scripts && python -m pytest tests/test_build_deck.py -k "in_place or slide_level" -v`
Expected: PASS (2 tests).

- [ ] **Step 5: Commit**

```bash
rm -f .git/index.lock
git add skills/academic-paper-to-pptx/scripts/build_deck.py skills/academic-paper-to-pptx/scripts/tests/test_build_deck.py
git commit -m "feat(build_deck): recursive spec token substitution + slide-level claim assignment"
```

---

### Task 3: Shared render helpers + title / bullets / figure renderers + Ctx

**Files:**
- Modify: `skills/academic-paper-to-pptx/scripts/build_deck.py`
- Test: `skills/academic-paper-to-pptx/scripts/tests/test_build_deck.py`

- [ ] **Step 1: Write failing tests for the simple renderers**

```python
# --- deck build fixtures ---
import shutil
from pptx import Presentation

def _skill_root():
    # scripts/ -> skill root
    return os.path.dirname(os.path.dirname(os.path.abspath(bd.__file__)))

def _new_ctx(tmp_path):
    out = str(tmp_path / "out.pptx")
    shutil.copy(os.path.join(_skill_root(), "references", "template.pptx"), out)
    prs = Presentation(out)
    layouts = {l.name: l for l in prs.slide_masters[0].slide_layouts}
    ctx = bd.Ctx(prs=prs, layouts=layouts, study="TEST-01",
                 citation="Author et al. 2026", skill_root=_skill_root(),
                 images_base=str(tmp_path))
    return ctx, out

def _texts(slide):
    out = []
    for sh in slide.shapes:
        if sh.has_text_frame:
            out.append(sh.text_frame.text)
    return "\n".join(out)

def test_render_title_sets_trial_name(tmp_path):
    ctx, out = _new_ctx(tmp_path)
    s = {"type": "title", "title": "TEST-01: A Trial", "subtitle": "Phase 3",
         "authors": "Author A", "affiliation": "Center X", "date": "Journal 2026"}
    bd.render_title(ctx, s, 1)
    slide = ctx.prs.slides[0]
    assert "TEST-01: A Trial" in _texts(slide)
    assert "Phase 3" in _texts(slide)

def test_render_bullets_counts_and_badge(tmp_path):
    ctx, out = _new_ctx(tmp_path)
    s = {"type": "bullets", "title": "Background",
         "bullets": [{"text": "Point one"}, {"text": "Point two", "sub": ["sub a"]}]}
    bd.render_bullets(ctx, s, 2)
    slide = ctx.prs.slides[0]
    txt = _texts(slide)
    assert "Point one" in txt and "Point two" in txt and "sub a" in txt
    assert "TEST-01" in txt          # badge
    assert "Author et al. 2026" in txt  # citation

def test_render_bullets_font_floors(tmp_path):
    ctx, out = _new_ctx(tmp_path)
    s = {"type": "bullets", "title": "T",
         "bullets": [{"text": "L0"}, {"text": "L0b", "sub": ["L1"]}]}
    bd.render_bullets(ctx, s, 2)
    slide = ctx.prs.slides[0]
    body = None
    for sh in slide.shapes:
        if sh.is_placeholder and sh.placeholder_format.idx == 1:
            body = sh
    sizes = {}
    for p in body.text_frame.paragraphs:
        for r in p.runs:
            if r.font.size:
                sizes.setdefault(p.level, set()).add(r.font.size.pt)
    assert all(x >= 18 for x in sizes.get(0, {18}))
    assert all(x >= 16 for x in sizes.get(1, {16}))

def test_render_figure_embeds_picture(tmp_path):
    # make a tiny PNG
    from PIL import Image as PILImage
    p = tmp_path / "fig.png"
    PILImage.new("RGB", (800, 600), "white").save(p)
    ctx, out = _new_ctx(tmp_path)
    s = {"type": "figure", "title": "OS", "image": "fig.png",
         "key_message": "median", "caption": ["line1"]}
    bd.render_figure(ctx, s, 3)
    slide = ctx.prs.slides[0]
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    assert any(sh.shape_type == MSO_SHAPE_TYPE.PICTURE for sh in slide.shapes)
```

- [ ] **Step 2: Run to verify they fail**

Run: `cd skills/academic-paper-to-pptx/scripts && python -m pytest tests/test_build_deck.py -k "render_title or render_bullets or render_figure" -v`
Expected: FAIL — `AttributeError: module 'build_deck' has no attribute 'Ctx'`.

- [ ] **Step 3: Add `Ctx`, shared helpers, and the three renderers to `build_deck.py`**

```python
# ═══════════ Build context + shared helpers ═══════════

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


# ═══════════ Renderers ═══════════

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
```

- [ ] **Step 4: Run to verify they pass**

Run: `cd skills/academic-paper-to-pptx/scripts && python -m pytest tests/test_build_deck.py -k "render_title or render_bullets or render_figure" -v`
Expected: PASS (4 tests).

- [ ] **Step 5: Commit**

```bash
rm -f .git/index.lock
git add skills/academic-paper-to-pptx/scripts/build_deck.py skills/academic-paper-to-pptx/scripts/tests/test_build_deck.py
git commit -m "feat(build_deck): Ctx + shared helpers + title/bullets/figure renderers"
```

---

### Task 4: Table renderer (2–3 arms, centered, optional note)

**Files:**
- Modify: `skills/academic-paper-to-pptx/scripts/build_deck.py`
- Test: `skills/academic-paper-to-pptx/scripts/tests/test_build_deck.py`

- [ ] **Step 1: Write failing tests**

```python
def test_render_table_dims_center_and_floor(tmp_path):
    ctx, out = _new_ctx(tmp_path)
    s = {"type": "table", "title": "Baseline",
         "label_header": "Characteristic",
         "arms": [{"label": "IP (n=148)", "color": "teal"},
                  {"label": "PS (n=74)", "color": "gray"}],
         "rows": [{"label": "Age", "vals": ["60 (24-70)", "56 (23-74)"], "bold": True},
                  {"label": "Sex", "hdr": True},
                  {"label": "  Female", "vals": ["68 (45.9)", "36 (48.6)"]}],
         "note": "No treatment-related deaths."}
    bd.render_table(ctx, s, 8)
    slide = ctx.prs.slides[0]
    tbl = None
    for sh in slide.shapes:
        if sh.has_table:
            tbl = sh
    assert tbl is not None
    assert len(tbl.table.rows) == 4       # header + 3
    assert len(tbl.table.columns) == 3    # label + 2 arms
    # centered: left margin symmetric within 3 EMU
    total = sum(c.width for c in tbl.table.columns)
    assert abs(tbl.left - (bd.SLIDE_W - total) // 2) <= 3
    # font floors
    for row in tbl.table.rows:
        for cell in row.cells:
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.size:
                        assert r.font.size.pt >= 14
    # note present
    txt = _texts(slide)
    assert "No treatment-related deaths." in txt
```

- [ ] **Step 2: Run to verify it fails**

Run: `cd skills/academic-paper-to-pptx/scripts && python -m pytest tests/test_build_deck.py -k "render_table" -v`
Expected: FAIL — `AttributeError: module 'build_deck' has no attribute 'render_table'`.

- [ ] **Step 3: Add `render_table` to `build_deck.py`**

```python
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
```

- [ ] **Step 4: Run to verify it passes**

Run: `cd skills/academic-paper-to-pptx/scripts && python -m pytest tests/test_build_deck.py -k "render_table" -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
rm -f .git/index.lock
git add skills/academic-paper-to-pptx/scripts/build_deck.py skills/academic-paper-to-pptx/scripts/tests/test_build_deck.py
git commit -m "feat(build_deck): table renderer (2-3 arms, centered, optional note)"
```

---

### Task 5: Study-design + CONSORT renderers (parameterized, 2-arm topology)

**Files:**
- Modify: `skills/academic-paper-to-pptx/scripts/build_deck.py`
- Test: `skills/academic-paper-to-pptx/scripts/tests/test_build_deck.py`

- [ ] **Step 1: Write failing tests**

```python
def test_render_study_design_diagram_floor_and_content(tmp_path):
    ctx, out = _new_ctx(tmp_path)
    s = {"type": "study_design",
         "description": "Multicenter phase 3 trial in 222 patients",
         "eligibility": ["Age 18-75 y", "Gastric adenocarcinoma"],
         "randomization": "2:1", "n": "N = 222",
         "arms": [{"name": "IP group (n=148)", "detail": "IP+IV paclitaxel + S-1", "color": "teal"},
                  {"name": "PS group (n=74)", "detail": "IV paclitaxel + S-1", "color": "gray"}],
         "primary_endpoints": ["Overall survival"],
         "secondary_endpoints": ["PFS", "Safety"],
         "registration": "ChiCTR-IIR-16009802",
         "footer": "Treatment until progression"}
    bd.render_study_design(ctx, s, 5)
    slide = ctx.prs.slides[0]
    txt = _texts(slide)
    for tok in ["N = 222", "2:1", "IP group (n=148)", "Overall survival",
                "ChiCTR-IIR-16009802"]:
        assert tok in txt
    # diagram text floor 12pt
    for sh in slide.shapes:
        if sh.has_text_frame and not (sh.is_placeholder and sh.placeholder_format.idx == 0):
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.size:
                        assert r.font.size.pt >= 12

def test_render_consort_content(tmp_path):
    ctx, out = _new_ctx(tmp_path)
    s = {"type": "consort",
         "assessed": "246 Assessed for eligibility",
         "excluded": ["8 Excluded", "6 Did not meet criteria; 2 declined"],
         "randomized": "238 Randomized (2:1)",
         "arms": [
            {"color": "teal",
             "allocated": ["158 Allocated to IP group", "IP + IV paclitaxel + S-1"],
             "not_received": ["10 Did not receive intervention", "declined IP port"],
             "received": "148 Received intervention (mITT)",
             "outcomes": ["Conversion surgery: 75 (63 R0, 12 R2)", "132 died"]},
            {"color": "gray",
             "allocated": ["80 Allocated to PS group", "IV paclitaxel + S-1"],
             "not_received": ["6 Did not receive intervention", "local treatment"],
             "received": "74 Received intervention (mITT)",
             "outcomes": ["Conversion surgery: 26 (22 R0, 4 R2)", "69 died"]}]}
    bd.render_consort(ctx, s, 7)
    slide = ctx.prs.slides[0]
    txt = _texts(slide)
    for tok in ["246", "238", "158", "80", "148", "74", "75", "26"]:
        assert tok in txt
    for p in slide.shapes:
        pass  # smoke
```

- [ ] **Step 2: Run to verify they fail**

Run: `cd skills/academic-paper-to-pptx/scripts && python -m pytest tests/test_build_deck.py -k "study_design or consort" -v`
Expected: FAIL — no `render_study_design`.

- [ ] **Step 3: Add `render_study_design` and `render_consort` to `build_deck.py`**

(Ported from `build/build_dragon01.py`'s `build_design` / `build_consort`, made data-driven over `arms`. Vertical arm spacing computed so 2 arms match the current layout; ≥3 arms compress evenly.)

```python
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
    pe = etf.paragraphs[0]; pe.text = "Primary Endpoint" + ("s" if len(s.get("primary_endpoints", [])) > 1 else "")
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
    # column centers spread across the slide
    centers = [3350000, 8850000] if len(arms) == 2 else \
        [int(2200000 + (7800000 / max(len(arms) - 1, 1)) * i) for i in range(len(arms))]
    bw = 3450000
    top_c = sum(centers) // len(centers)

    box(top_c - 1900000, 1500000, 3800000, 470000,
        [(s["assessed"], True, BODY, 13)])
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
```

- [ ] **Step 4: Run to verify they pass**

Run: `cd skills/academic-paper-to-pptx/scripts && python -m pytest tests/test_build_deck.py -k "study_design or consort" -v`
Expected: PASS (2 tests).

- [ ] **Step 5: Commit**

```bash
rm -f .git/index.lock
git add skills/academic-paper-to-pptx/scripts/build_deck.py skills/academic-paper-to-pptx/scripts/tests/test_build_deck.py
git commit -m "feat(build_deck): parameterized study-design + CONSORT renderers"
```

---

### Task 6: `build_presentation` orchestrator + CLI + `main`

**Files:**
- Modify: `skills/academic-paper-to-pptx/scripts/build_deck.py`
- Create: `skills/academic-paper-to-pptx/scripts/tests/fixtures/mini_claims.json`
- Create: `skills/academic-paper-to-pptx/scripts/tests/fixtures/mini_spec.json`
- Test: `skills/academic-paper-to-pptx/scripts/tests/test_build_deck.py`

- [ ] **Step 1: Create the fixtures** (public/synthetic content only)

`tests/fixtures/mini_claims.json`:

```json
{
  "deck": {"title": "MINI-01", "source": "synthetic.md", "nct": "NCT00000000", "date": "2026-07-12"},
  "claims": [
    {"id": "eff-01", "category": "efficacy", "label": "Median OS A",
     "value": "19.4 months (95% CI 17.1-22.9)", "slide": null, "source_hint": "synthetic"},
    {"id": "eff-02", "category": "efficacy", "label": "Median OS B",
     "value": "13.9 months (95% CI 10.3-16.1)", "slide": null, "source_hint": "synthetic"},
    {"id": "eff-03", "category": "efficacy", "label": "OS HR",
     "value": "HR 0.67 (95% CI 0.50-0.90); P = .01", "slide": null, "source_hint": "synthetic"},
    {"id": "dem-01", "category": "demographics", "label": "Age",
     "value": "A 60 (24-70); B 56 (23-74)", "slide": null, "source_hint": "synthetic"}
  ]
}
```

`tests/fixtures/mini_spec.json`:

```json
{
  "deck": {"study": "MINI-01", "citation": "Synthetic et al. 2026",
           "title": "MINI-01: Synthetic Trial", "subtitle": "A Phase 3 Synthetic Trial",
           "authors": "Synthetic A, Synthetic B", "affiliation": "Nowhere Center",
           "date": "Synthetic Journal 2026"},
  "slides": [
    {"type": "title"},
    {"type": "bullets", "title": "Background",
     "bullets": [{"text": "A synthetic point"}, {"text": "Another", "sub": ["a subpoint"]}]},
    {"type": "table", "title": "Baseline", "label_header": "Characteristic",
     "key_message": "Balanced arms.",
     "arms": [{"label": "A (n=148)", "color": "teal"}, {"label": "B (n=74)", "color": "gray"}],
     "rows": [{"label": "Age, median (range)", "vals": ["60 (24-70)", "56 (23-74)"], "bold": true}],
     "claims": ["dem-01"]},
    {"type": "bullets", "title": "Primary Result",
     "bullets": [{"text": "Median OS {{eff-01}} vs {{eff-02}}; {{eff-03}}", "bold": true}]}
  ]
}
```

- [ ] **Step 2: Write the failing integration test**

```python
def test_build_presentation_end_to_end(tmp_path):
    here = os.path.dirname(os.path.abspath(__file__))
    spec = os.path.join(here, "fixtures", "mini_spec.json")
    claims = os.path.join(here, "fixtures", "mini_claims.json")
    out = str(tmp_path / "mini.pptx")
    resolved = str(tmp_path / "mini_claims_resolved.json")
    rc = bd.build_from_files(spec, claims, out, resolved_path=resolved,
                             images_base=os.path.dirname(spec))
    assert rc == 0
    assert os.path.exists(out)
    # every claim assigned a slide in the resolved file
    with open(resolved) as f:
        rdata = json.load(f)
    for c in rdata["claims"]:
        assert c["slide"] is not None, c["id"]
    # deck opens and has 4 slides
    prs = Presentation(out)
    assert len(prs.slides._sldIdLst) == 4

def test_build_unreferenced_claim_fails_without_flag(tmp_path, monkeypatch):
    # remove the reference to dem-01 to leave it unreferenced
    here = os.path.dirname(os.path.abspath(__file__))
    with open(os.path.join(here, "fixtures", "mini_spec.json")) as f:
        spec = json.load(f)
    for sl in spec["slides"]:
        sl.pop("claims", None)
    spec_path = tmp_path / "spec_missing.json"
    spec_path.write_text(json.dumps(spec))
    claims = os.path.join(here, "fixtures", "mini_claims.json")
    out = str(tmp_path / "x.pptx")
    rc = bd.build_from_files(str(spec_path), claims, out, resolved_path=None,
                             images_base=here, allow_unreferenced=False)
    assert rc != 0  # dem-01 never referenced
```

- [ ] **Step 3: Run to verify the tests fail**

Run: `cd skills/academic-paper-to-pptx/scripts && python -m pytest tests/test_build_deck.py -k "end_to_end or unreferenced_claim_fails" -v`
Expected: FAIL — no `build_from_files`.

- [ ] **Step 4: Add orchestrator + CLI to `build_deck.py`**

```python
RENDERERS = {
    "title": render_title,
    "bullets": render_bullets,
    "table": render_table,
    "figure": render_figure,
    "study_design": render_study_design,
    "consort": render_consort,
}


def resolve_skill_root(cli):
    if cli:
        return cli
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_root = os.path.normpath(os.path.join(script_dir, ".."))
    if os.path.isdir(os.path.join(default_root, "references")):
        return default_root
    return "/mnt/skills/user/academic-paper-to-pptx"


def build_from_files(spec_path, claims_path, out_path, resolved_path=None,
                     images_base=None, skill_root=None, allow_unreferenced=False):
    with open(spec_path) as f:
        spec = json.load(f)
    with open(claims_path) as f:
        claims_data = json.load(f)
    if images_base is None:
        images_base = os.path.dirname(os.path.abspath(spec_path))
    skill_root = resolve_skill_root(skill_root)

    claims = Claims(claims_data)
    template = os.path.join(skill_root, "references", "template.pptx")
    shutil.copy(template, out_path)
    prs = Presentation(out_path)
    layouts = {l.name: l for l in prs.slide_masters[0].slide_layouts}

    deck_meta = spec.get("deck", {})
    ctx = Ctx(prs=prs, layouts=layouts,
              study=deck_meta.get("study", ""),
              citation=deck_meta.get("citation", ""),
              skill_root=skill_root, images_base=images_base)

    for i, sdata in enumerate(spec.get("slides", []), start=1):
        # title slide inherits deck meta fields it doesn't override
        if sdata.get("type") == "title":
            for k in ("title", "subtitle", "authors", "affiliation", "date"):
                sdata.setdefault(k, deck_meta.get(k))
        substitute_in_place(sdata, claims, i)
        assign_slide_claims(sdata, claims, i)
        t = sdata.get("type")
        if t not in RENDERERS:
            print(f"ERROR: slide {i}: unknown slide type {t!r} "
                  f"(known: {sorted(RENDERERS)})", file=sys.stderr)
            return 3
        RENDERERS[t](ctx, sdata, i)

    prs.save(out_path)

    unref = claims.unreferenced()
    if unref and not allow_unreferenced:
        print(f"ERROR: {len(unref)} claim(s) never referenced by any slide: "
              f"{', '.join(unref)}\n  Every claim must appear on a slide (use a "
              f"{{{{id}}}} token or a slide-level \"claims\" list), or pass "
              f"--allow-unreferenced.", file=sys.stderr)
        return 2

    if resolved_path:
        with open(resolved_path, "w") as f:
            json.dump(claims.resolved_json(), f, indent=2, ensure_ascii=False)
        print(f"Resolved claims: {os.path.abspath(resolved_path)}")

    print(f"Saved deck: {os.path.abspath(out_path)} "
          f"({len(spec.get('slides', []))} slides, "
          f"{len(claims.referenced)}/{len(claims.by_id)} claims placed)")
    if unref:
        print(f"WARNING: unreferenced claims (allowed): {', '.join(unref)}", file=sys.stderr)
    return 0


def main():
    ap = argparse.ArgumentParser(description="Build a Moffitt deck from a deck spec + claims")
    ap.add_argument("spec", help="deck_spec.json")
    ap.add_argument("--claims", required=True, help="claims JSON (qa-checklist.md schema)")
    ap.add_argument("--out", "-o", required=True, help="output .pptx path")
    ap.add_argument("--resolved-claims", default=None,
                    help="write claims JSON with slide numbers filled (for qa_crosscheck)")
    ap.add_argument("--images-base", default=None,
                    help="base dir for relative figure paths (default: spec dir)")
    ap.add_argument("--skill-root", default=None, help="override skill root")
    ap.add_argument("--allow-unreferenced", action="store_true",
                    help="do not fail when some claims are never placed on a slide")
    args = ap.parse_args()
    for path in (args.spec, args.claims):
        if not os.path.exists(path):
            print(f"Error: not found: {path}", file=sys.stderr)
            sys.exit(2)
    rc = build_from_files(args.spec, args.claims, args.out,
                          resolved_path=args.resolved_claims,
                          images_base=args.images_base, skill_root=args.skill_root,
                          allow_unreferenced=args.allow_unreferenced)
    sys.exit(rc)


if __name__ == "__main__":
    main()
```

- [ ] **Step 5: Run to verify the tests pass, then run the whole file**

Run: `cd skills/academic-paper-to-pptx/scripts && python -m pytest tests/test_build_deck.py -v`
Expected: PASS (all tests).

- [ ] **Step 6: Commit**

```bash
rm -f .git/index.lock
git add skills/academic-paper-to-pptx/scripts/build_deck.py skills/academic-paper-to-pptx/scripts/tests/
git commit -m "feat(build_deck): orchestrator + CLI; end-to-end + unreferenced-claim guard tests"
```

---

### Task 7: DRAGON-01 regression — author the deck spec, build, gate on qa_crosscheck

**Files:**
- Create: `build/DRAGON01_deck_spec.json`
- Uses (read-only): `build/DRAGON01_claims.json`, `build/paper_text.md`, `build/figures/*.png`

- [ ] **Step 1: Author `build/DRAGON01_deck_spec.json`**

Express the 16 DRAGON-01 slides as a spec, faithfully porting content from `build/build_dragon01.py`:
- 1 `title`; 2–4 `bullets`; 5 `study_design`; 6 `bullets` (stats); 7 `consort`; 8–9 `table`; 10–12 `figure`; 13 `bullets`; 14 `table` (with `note`); 15–16 `bullets`.
- Key messages and figure captions use `{{claim-id}}` tokens where a claim's value fits (e.g. slide 10 `key_message`: `"Median OS {{eff-01}} vs {{eff-02}}; {{eff-03}}"`).
- Table slides (8, 9, 14) inline the cell values (as in `build_dragon01.py`) and declare their claim IDs in a slide-level `"claims": [...]` list (dem-00…dem-21 on 8/9; saf-01…saf-09 + the port/postop notes on 14).
- `study_design` (slide 5) carries des-01…des-09 via tokens/inline + `"claims"`; `consort` (slide 7) carries con-01…con-10; slide 6 carries des-07/des-08; slide 13 carries eff-11…eff-17.
- Every one of the 71 claim IDs must be referenced exactly once so the builder's unreferenced check passes.

- [ ] **Step 2: Build the deck from the spec**

Run:
```bash
cd build && python ../skills/academic-paper-to-pptx/scripts/build_deck.py DRAGON01_deck_spec.json \
    --claims DRAGON01_claims.json --out DRAGON01_deck_v2.pptx \
    --resolved-claims DRAGON01_claims_resolved.json
```
Expected: `Saved deck: .../DRAGON01_deck_v2.pptx (16 slides, 71/71 claims placed)` and exit 0. If it reports unreferenced claims, add their references to the spec and re-run until 71/71.

- [ ] **Step 3: Run the QA gate against the resolved claims**

Run:
```bash
cd build && python ../skills/academic-paper-to-pptx/scripts/qa_crosscheck.py DRAGON01_deck_v2.pptx \
    --claims DRAGON01_claims_resolved.json --mode paper --source-text paper_text.md \
    --output DRAGON01_deck_v2_QA.md
```
Expected: `Result: 71 pass / 0 fail / N warn` (warnings only for the figure-picture confirmations, matching the original gate). Fix every ❌ (usually a missing claim reference or an inline typo) and re-run until 0 fail.

- [ ] **Step 4: Render + eyeball a couple of slides (visual sanity)**

Run:
```bash
cd build && soffice --headless --convert-to pdf DRAGON01_deck_v2.pptx && \
    pdftoppm -jpeg -r 100 -f 5 -l 7 DRAGON01_deck_v2.pdf sd_check
```
Read `sd_check-05.jpg` (study design) and `sd_check-07.jpg` (CONSORT); confirm no overflow/cutoff. (macOS soffice fallback: `/Applications/LibreOffice.app/Contents/MacOS/soffice`.)

- [ ] **Step 5: Commit the regression proof**

```bash
rm -f .git/index.lock
git add build/DRAGON01_deck_spec.json build/DRAGON01_deck_v2.pptx build/DRAGON01_deck_v2_QA.md build/DRAGON01_claims_resolved.json
git commit -m "test(regression): DRAGON-01 rebuilt from deck-spec via build_deck.py — 71 pass / 0 fail"
```

---

### Task 8: Update SKILL.md + slide-builders.md + deck-spec schema doc

**Files:**
- Create: `skills/academic-paper-to-pptx/references/deck-spec.md`
- Modify: `skills/academic-paper-to-pptx/SKILL.md` (Phase 3 section)
- Modify: `skills/academic-paper-to-pptx/references/slide-builders.md`

- [ ] **Step 1: Write `references/deck-spec.md`**

Document the deck-spec schema: the `deck` block, each slide `type` and its fields, the `{{claim-id}}` substitution rule, the "tables inline values + slide-level `claims` list" rule, the "every claim must be referenced" contract, and one minimal worked example per type (copy from `mini_spec.json` + the DRAGON study_design/consort blocks). State explicitly: **the spec has no font field — sizes are fixed in `build_deck.py` and always ≥ floors.**

- [ ] **Step 2: Rewrite SKILL.md Phase 3**

Replace "hand-author python-pptx" with:
1. Author `<deckname>_deck_spec.json` (schema: `references/deck-spec.md`), referencing claim IDs from `<deckname>_claims.json`.
2. `python "<skill-root>/scripts/build_deck.py" <deckname>_deck_spec.json --claims <deckname>_claims.json --out <deckname>.pptx --resolved-claims <deckname>_claims_resolved.json`
3. Phase 4 QA runs against `--claims <deckname>_claims_resolved.json` (slide numbers auto-filled by the builder — the model no longer maintains the `slide` field by hand).

Keep the data-fidelity, font-floor, and figure-embedding notes. Add the token-efficiency rationale (one deck ≈ compact spec, not ~28 KB of code) and the "no font field ⇒ no sub-floor fonts" integrity property.

- [ ] **Step 3: Trim slide-builders.md**

Replace the large per-type code blocks with: a one-paragraph note that the renderers now live in `scripts/build_deck.py`, and a compact table mapping each spec slide `type` → what the renderer draws → which fields it reads. Keep the color palette and EMU reference tables (still useful context). Point to `references/deck-spec.md` for the authoring schema.

- [ ] **Step 4: Sanity-check the SKILL still self-describes a full build**

Read the edited SKILL.md end-to-end; confirm Phases 1→4 form a complete, runnable pipeline with the new build step and that all referenced paths/filenames exist.

- [ ] **Step 5: Commit**

```bash
rm -f .git/index.lock
git add skills/academic-paper-to-pptx/SKILL.md skills/academic-paper-to-pptx/references/deck-spec.md skills/academic-paper-to-pptx/references/slide-builders.md
git commit -m "docs(skill): Phase 3 = deck-spec -> build_deck.py; add deck-spec schema; trim slide-builders"
```

---

### Task 9: Sync skills, run the qa_crosscheck self-test, update spec.md/CLAUDE.md, checkpoint

**Files:**
- Run: `tools/sync-skills.sh`
- Modify: `spec.md`, `CLAUDE.md` (Lessons Learned if warranted)

- [ ] **Step 1: Run the qa_crosscheck self-test (regression on the tool itself, per CLAUDE.md)**

Confirm a clean synthetic deck exits 0 and an injected wrong-value/orphan/`[CHECK]`/sub-floor/full-slide deck flags all five. (Use the existing self-test procedure; the new builder doesn't change qa_crosscheck, but verify nothing regressed.)

- [ ] **Step 2: Run `tools/sync-skills.sh`**

Run: `bash tools/sync-skills.sh`
Expected: installs to `~/.claude/skills/` and repackages `dist/*.skill`. Confirm `build_deck.py`, `deck-spec.md`, and the trimmed `slide-builders.md` are in the installed copy. (The `scripts/tests/` dir is dev-only — confirm sync excludes it or that its inclusion is harmless.)

- [ ] **Step 3: Update spec.md**

Mark Token-Efficiency **option 1 as IMPLEMENTED (2026-07-12)** with a one-line result (DRAGON-01 rebuilt from a ~X KB spec vs the 28 KB hand-authored builder; 71 pass / 0 fail). Update the top Status block. Note Goal 2 is now next.

- [ ] **Step 4: Append a Lessons-Learned entry to CLAUDE.md** (only if something non-obvious surfaced during the build — e.g., a template-layout quirk, an EMU-spacing gotcha in the parameterized diagrams, or a sync exclusion issue).

- [ ] **Step 5: Commit + checkpoint**

```bash
rm -f .git/index.lock
git add -A
git commit -m "chore: sync skills + dist; spec.md option 1 IMPLEMENTED; lessons update"
```
Then run the `/checkpoint` skill: update memory (`presentations-deck-skills-project.md`), report every file changed, and emit a self-contained resume prompt (next step: Goal 2 Presenton MCP). Remind Allan to re-upload `dist/*.skill` to the Claude desktop app.

---

## Self-Review

**Spec coverage** (against spec.md Token-Efficiency option 1 + resume prompt):
- "Model emits compact typed JSON referencing claim IDs" → deck-spec.md + `{{claim-id}}` tokens (Tasks 1–2, 8). ✅
- "Single committed build_deck.py consumes spec + claims" → Task 6 orchestrator. ✅
- "Cuts ~28 KB code to ~5–8 KB data" → Task 7 DRAGON spec replaces the 28 KB builder; result recorded in spec.md (Task 9). ✅ (measure actual size in Task 7.)
- "Converges with parsed.json → build_from_parsed.py pattern" → same structured-JSON→builder→qa_crosscheck shape, mirrored helpers (Task 3). ✅
- "Claim extraction stays on strong model; qa_crosscheck gates every build" → builder only mechanizes the build; Task 7 gate; guardrail stated in header. ✅
- "Font floors as hard override" → `pt(size, floor)` clamp + no font field in spec; tests in Tasks 3–5. ✅
- "Every claim assigned before QA" → unreferenced-claim guard (Task 6), resolved-claims file auto-fills `slide`. ✅

**Placeholder scan:** No TBD/TODO/"handle edge cases"; every code step shows complete code; fixtures are concrete. Task 7 Step 1 (authoring the DRAGON spec) is descriptive rather than a full JSON dump — that content is a faithful port of the already-shown `build/build_dragon01.py` payload, produced during execution and gated by qa_crosscheck; acceptable since it is data, not logic, and fully specified by the source it ports.

**Type consistency:** `Ctx(prs, layouts, study, citation, skill_root, images_base)` used identically in tests and orchestrator. Renderer signature `render_x(ctx, s, n)` uniform; `RENDERERS` keys match spec `type` values (`title/bullets/table/figure/study_design/consort`). `Claims` methods (`value/assign/substitute/resolved_json/unreferenced`) consistent across Tasks 1–6. `build_from_files(...)` signature identical in tests (Task 6) and CLI. `pt(size, floor)` and `col(name)` used consistently.
