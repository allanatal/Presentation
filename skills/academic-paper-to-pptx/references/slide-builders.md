# Slide Builder Patterns — python-pptx with Moffitt Template

All patterns use the bundled `template.pptx` which provides the Moffitt slide master.
The master handles: top bar, bottom 4-color bar, Moffitt badge icon, title formatting, bullet styling.

## Setup

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import shutil, os

shutil.copy("/path/to/skill/references/template.pptx", "/home/claude/output.pptx")
prs = Presentation("/home/claude/output.pptx")

master = prs.slide_masters[0]
LAYOUTS = {l.name: l for l in master.slide_layouts}

# Colors for manual elements
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

FONT = "Arial"
```

## Shared Helper: Study Badge + Citation

Call on every content slide (not title slide). These are NOT in the template master — they were added manually per-slide in the model.

```python
def add_badge_and_citation(slide, study_name, citation_text):
    """Add study name badge (top-right) and citation (bottom-right)."""
    # Study name - positioned next to the inherited Moffitt icon
    badge = slide.shapes.add_textbox(Emu(9984826), Emu(157655), Emu(1671098), Emu(369332))
    p = badge.text_frame.paragraphs[0]
    p.text = study_name
    p.font.size = Pt(11)
    p.font.color.rgb = TITLE_BLUE
    p.alignment = PP_ALIGN.RIGHT

    # Citation
    cite = slide.shapes.add_textbox(Emu(9987101), Emu(6463328), Emu(2204899), Emu(307777))
    p = cite.text_frame.paragraphs[0]
    p.text = citation_text
    p.font.size = Pt(10)
    p.font.color.rgb = GRAY
    p.alignment = PP_ALIGN.RIGHT
```

---

## Title Slide

Uses `Title Slide` layout. Override placeholder font for the trial name color.

```python
def build_title_slide(prs, trial_name, author, affiliations, date, logo_path=None):
    slide = prs.slides.add_slide(LAYOUTS['Title Slide'])

    # idx=0: CENTER_TITLE — trial name
    title_ph = slide.placeholders[0]
    title_ph.text = trial_name
    for para in title_ph.text_frame.paragraphs:
        para.font.size = Pt(24)
        para.font.bold = True
        para.font.color.rgb = RED
        para.font.name = FONT

    # idx=1: SUBTITLE — author name
    sub_ph = slide.placeholders[1]
    sub_ph.text = author
    for para in sub_ph.text_frame.paragraphs:
        para.font.size = Pt(20)
        para.font.bold = True
        para.font.color.rgb = BODY
        para.font.name = FONT

    # Affiliations + date as manual textbox (more control)
    lines = affiliations + ["", date]
    aff_box = slide.shapes.add_textbox(Emu(1154187), Emu(4928445), Emu(9340438), Emu(1200329))
    aff_box.text_frame.word_wrap = True
    for i, line in enumerate(lines):
        if i == 0:
            p = aff_box.text_frame.paragraphs[0]
        else:
            p = aff_box.text_frame.add_paragraph()
        p.text = line
        p.font.size = Pt(14) if line != date else Pt(12)
        p.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        p.font.name = FONT
        p.alignment = PP_ALIGN.CENTER

    # Logo (if WMF/PNG path provided)
    if logo_path and os.path.exists(logo_path):
        slide.shapes.add_picture(logo_path, Emu(4790362), Emu(798667), Emu(2266391), Emu(578214))

    return slide
```

---

## Content Slide (Bullets)

For Research Objective, Literature Background, Hypotheses, Statistical Methods, Discussion.

```python
def build_bullet_slide(prs, title, study_name, citation, bullets):
    """
    bullets: list of dicts: [
        {"text": "Main point", "bold": False, "sub": ["Sub-point 1", "Sub-point 2"]},
        {"text": "Another point", "bold": True},
    ]
    """
    slide = prs.slides.add_slide(LAYOUTS['Title and Content'])

    # Title (placeholder idx=0) — formatting inherited from master
    slide.placeholders[0].text = title

    # Content (placeholder idx=1) — use bullet levels
    tf = slide.placeholders[1].text_frame
    tf.clear()

    first = True
    for b in bullets:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.text = b["text"]
        p.level = 0
        p.font.name = FONT
        p.font.size = Pt(18)  # MINIMUM 18pt for body text on content slides
        if b.get("bold"):
            p.font.bold = True

        for sub_text in b.get("sub", []):
            ps = tf.add_paragraph()
            ps.text = sub_text
            ps.level = 1
            ps.font.name = FONT
            ps.font.size = Pt(16)  # Sub-bullets: 16pt minimum

    add_badge_and_citation(slide, study_name, citation)
    return slide
```

---

## Study Design Diagram

Build on `Title Only` layout with python-pptx shapes. Adapt for each study.

```python
def build_study_design(prs, study_name, citation, design):
    """
    design = {
        "description": "Global phase 3 trial of ...",
        "eligibility": ["Age ≥18 years", ...],
        "stratification": ["Geographic region", ...],
        "sample_size": 914,
        "randomization": "1:1:1",
        "arms": [
            {"name": "Arm A", "label": "Trastuzumab + CT", "color": CONTROL_GRAY, "details": ""},
            {"name": "Arm B", "label": "Zanidatamab + CT", "color": ARM_BLUE, "details": "1800 mg..."},
            {"name": "Arm C", "label": "Zani + tisle + CT", "color": ARM_TEAL, "details": "..."},
        ],
        "primary_endpoints": ["PFS (per BICR)", "OS"],
        "secondary_endpoints": ["cORR", "AE frequency"],
        "nct": "NCT05152147",
        "duration_note": "Until progression/death/toxicity",
    }
    """
    slide = prs.slides.add_slide(LAYOUTS['Title Only'])
    slide.placeholders[0].text = "Study Design"

    # Description (italic subtitle)
    desc = slide.shapes.add_textbox(Emu(838200), Emu(1500000), Emu(10515600), Emu(500000))
    desc.text_frame.word_wrap = True
    p = desc.text_frame.paragraphs[0]
    p.text = design["description"]
    p.font.size = Pt(14)  # Description subtitle
    p.font.italic = True
    p.font.color.rgb = TITLE_BLUE
    p.font.name = FONT

    content_y = Emu(2100000)  # ~0.83" below description

    # ── Eligibility Box (left) ──
    elig_x, elig_w, elig_h = Emu(300000), Emu(3200000), Emu(3000000)
    elig_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, elig_x, content_y, elig_w, elig_h)
    elig_shape.fill.solid()
    elig_shape.fill.fore_color.rgb = LIGHT_BG
    elig_shape.line.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
    elig_shape.line.width = Pt(0.5)

    tf = elig_shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Key Eligibility Criteria"
    p.font.size = Pt(13)  # Box headings: 13pt bold
    p.font.bold = True
    p.font.color.rgb = BODY
    p.font.name = FONT
    for criterion in design["eligibility"]:
        pe = tf.add_paragraph()
        pe.text = f"• {criterion}"
        pe.font.size = Pt(12)  # MINIMUM 12pt in study design shapes
        pe.font.color.rgb = BODY
        pe.font.name = FONT

    # Stratification below
    ps = tf.add_paragraph()
    ps.text = ""
    ph = tf.add_paragraph()
    ph.text = "Stratification Factors"
    ph.font.size = Pt(12)  # Sub-headings: 12pt bold
    ph.font.bold = True
    ph.font.color.rgb = BODY
    ph.font.name = FONT
    for sf in design["stratification"]:
        pf = tf.add_paragraph()
        pf.text = f"• {sf}"
        pf.font.size = Pt(12)  # MINIMUM 12pt in study design shapes
        pf.font.color.rgb = BODY
        pf.font.name = FONT

    # ── Randomization Circle ──
    circ_x, circ_y, circ_d = Emu(3800000), content_y + Emu(800000), Emu(900000)
    circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, circ_x, circ_y, circ_d, circ_d)
    circ.fill.solid()
    circ.fill.fore_color.rgb = RGBColor(0x4A, 0x6A, 0x8A)
    circ.line.fill.background()

    circ_tf = circ.text_frame
    circ_tf.paragraphs[0].text = "R"
    circ_tf.paragraphs[0].font.size = Pt(16)
    circ_tf.paragraphs[0].font.bold = True
    circ_tf.paragraphs[0].font.color.rgb = WHITE
    circ_tf.paragraphs[0].font.name = FONT
    circ_tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    p_ratio = circ_tf.add_paragraph()
    p_ratio.text = design["randomization"]
    p_ratio.font.size = Pt(12)  # MINIMUM 12pt
    p_ratio.font.color.rgb = WHITE
    p_ratio.font.name = FONT
    p_ratio.alignment = PP_ALIGN.CENTER

    # N label below circle
    n_box = slide.shapes.add_textbox(circ_x - Emu(100000), circ_y + circ_d + Emu(50000),
                                      circ_d + Emu(200000), Emu(250000))
    pn = n_box.text_frame.paragraphs[0]
    pn.text = f"N = {design['sample_size']}"
    pn.font.size = Pt(12)  # MINIMUM 12pt
    pn.font.bold = True
    pn.font.color.rgb = BODY
    pn.font.name = FONT
    pn.alignment = PP_ALIGN.CENTER

    # ── Treatment Arm Boxes ──
    arm_x, arm_w = Emu(5200000), Emu(3200000)
    n_arms = len(design["arms"])
    arm_spacing = min(Emu(900000), Emu(2800000) // n_arms)

    for i, arm in enumerate(design["arms"]):
        arm_y = content_y + Emu(200000) + arm_spacing * i
        arm_h = arm_spacing - Emu(80000)

        arm_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, arm_x, arm_y, arm_w, arm_h)
        arm_shape.fill.solid()
        arm_shape.fill.fore_color.rgb = arm["color"]
        arm_shape.line.fill.background()

        atf = arm_shape.text_frame
        atf.word_wrap = True
        pa = atf.paragraphs[0]
        pa.text = f"{arm['name']}: {arm['label']}"
        pa.font.size = Pt(12)  # MINIMUM 12pt
        pa.font.bold = True
        pa.font.color.rgb = WHITE
        pa.font.name = FONT
        if arm.get("details"):
            pd = atf.add_paragraph()
            pd.text = arm["details"]
            pd.font.size = Pt(12)  # MINIMUM 12pt (was 8pt)
            pd.font.color.rgb = WHITE
            pd.font.name = FONT

    # ── Endpoints Box (far right) ──
    ep_x, ep_w, ep_h = Emu(8700000), Emu(2800000), Emu(2000000)
    ep_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, ep_x, content_y + Emu(200000), ep_w, ep_h)
    ep_shape.fill.solid()
    ep_shape.fill.fore_color.rgb = LIGHT_BG
    ep_shape.line.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
    ep_shape.line.width = Pt(0.5)

    etf = ep_shape.text_frame
    etf.word_wrap = True
    pe = etf.paragraphs[0]
    pe.text = "Primary Endpoints"
    pe.font.size = Pt(13)  # Box headings: 13pt bold
    pe.font.bold = True
    pe.font.color.rgb = BODY
    pe.font.name = FONT
    for ep in design["primary_endpoints"]:
        p = etf.add_paragraph()
        p.text = f"• {ep}"
        p.font.size = Pt(12)  # MINIMUM 12pt
        p.font.color.rgb = BODY
        p.font.name = FONT

    ps = etf.add_paragraph()
    ps.text = "Secondary Endpoints"
    ps.font.size = Pt(13)  # Box headings: 13pt bold
    ps.font.bold = True
    ps.font.color.rgb = TITLE_BLUE
    ps.font.name = FONT
    for ep in design["secondary_endpoints"]:
        p = etf.add_paragraph()
        p.text = f"• {ep}"
        p.font.size = Pt(12)  # MINIMUM 12pt
        p.font.color.rgb = BODY
        p.font.name = FONT

    if design.get("nct"):
        pnct = etf.add_paragraph()
        pnct.text = f"ClinicalTrials.gov: {design['nct']}"
        pnct.font.size = Pt(12)  # MINIMUM 12pt (was 8pt)
        pnct.font.color.rgb = GRAY
        pnct.font.name = FONT

    # ── Duration arrow (bottom) ──
    if design.get("duration_note"):
        dur = slide.shapes.add_textbox(Emu(3500000), Emu(5500000), Emu(7000000), Emu(400000))
        dur.text_frame.word_wrap = True
        pd = dur.text_frame.paragraphs[0]
        pd.text = design["duration_note"]
        pd.font.size = Pt(12)  # MINIMUM 12pt
        pd.font.color.rgb = BODY
        pd.font.name = FONT
        pd.alignment = PP_ALIGN.CENTER

    add_badge_and_citation(slide, study_name, citation)
    return slide
```

---

## Demographics Table

```python
def build_demographics(prs, study_name, citation, table_data):
    """
    table_data = {
        "key_message": "Demographics were balanced across all arms",
        "arms": [
            {"label": "Zani+Tisle+CT\n(n=302)", "color": ARM_TEAL},
            {"label": "Zani+CT\n(n=304)", "color": ARM_BLUE},
            {"label": "Tras+CT\n(n=308)", "color": CONTROL_GRAY},
        ],
        "rows": [
            {"category": "Age, median (range)", "values": ["63.0 (22-81)", "62.5 (25-87)", "64.0 (21-84)"], "bold": True},
            {"category": "Male sex", "values": ["244 (80.8)", "244 (80.3)", "238 (77.3)"]},
            {"category": "Geographic region", "values": ["", "", ""], "is_header": True},
            {"category": "  Asia", "values": ["159 (52.6)", "163 (53.6)", "165 (53.6)"]},
        ],
    }
    """
    slide = prs.slides.add_slide(LAYOUTS['Title Only'])
    slide.placeholders[0].text = "Demographics"

    # Key message
    if table_data.get("key_message"):
        msg = slide.shapes.add_textbox(Emu(838200), Emu(1400000), Emu(10515600), Emu(400000))
        p = msg.text_frame.paragraphs[0]
        p.text = table_data["key_message"]
        p.font.size = Pt(14)  # Key message: 14pt minimum
        p.font.italic = True
        p.font.color.rgb = TITLE_BLUE
        p.font.name = FONT

    n_cols = 1 + len(table_data["arms"])
    n_rows = 1 + len(table_data["rows"])

    # Calculate column widths
    label_w = Inches(2.5)
    data_w = Inches((10.5 - 2.5) / len(table_data["arms"]))
    col_widths = [label_w] + [data_w] * len(table_data["arms"])

    # Center table horizontally on slide (slide width = 12192000 EMU)
    total_table_w = sum(col_widths)
    table_x = (Emu(12192000) - total_table_w) // 2

    table_shape = slide.shapes.add_table(
        n_rows, n_cols,
        table_x, Emu(1900000),  # centered horizontally
        total_table_w, Emu(n_rows * 320000)
    )
    tbl = table_shape.table

    # Set column widths
    for i, w in enumerate(col_widths):
        tbl.columns[i].width = int(w)

    # Header row
    tbl.cell(0, 0).text = ""
    for j, arm in enumerate(table_data["arms"]):
        cell = tbl.cell(0, j + 1)
        cell.text = arm["label"]
        cell.fill.solid()
        cell.fill.fore_color.rgb = arm["color"]
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(14)  # Table headers: 14pt minimum
            p.font.bold = True
            p.font.color.rgb = WHITE
            p.font.name = FONT
            p.alignment = PP_ALIGN.CENTER

    # Data rows
    for i, row in enumerate(table_data["rows"]):
        # Label cell
        cell = tbl.cell(i + 1, 0)
        cell.text = row["category"]
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(14)  # MINIMUM 14pt in tables (16pt preferred)
            p.font.bold = row.get("bold", False) or row.get("is_header", False)
            p.font.color.rgb = BODY
            p.font.name = FONT

        # Value cells
        for j, val in enumerate(row["values"]):
            cell = tbl.cell(i + 1, j + 1)
            cell.text = val
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(14)  # MINIMUM 14pt in tables (16pt preferred)
                p.font.color.rgb = BODY
                p.font.name = FONT
                p.alignment = PP_ALIGN.CENTER

    add_badge_and_citation(slide, study_name, citation)
    return slide
```

---

## Figure Embedding

```python
def build_figure_slide(prs, title, study_name, citation, image_path, key_message=None):
    slide = prs.slides.add_slide(LAYOUTS['Title Only'])
    slide.placeholders[0].text = title

    if key_message:
        msg = slide.shapes.add_textbox(Emu(838200), Emu(1400000), Emu(10515600), Emu(400000))
        p = msg.text_frame.paragraphs[0]
        p.text = key_message
        p.font.size = Pt(14)  # Key message: 14pt minimum
        p.font.italic = True
        p.font.color.rgb = TITLE_BLUE
        p.font.name = FONT

    # Calculate aspect ratio
    from PIL import Image
    img = Image.open(image_path)
    img_w, img_h = img.size
    ratio = img_w / img_h

    max_w = Emu(10500000)  # ~11.5 inches
    max_h = Emu(4500000)   # ~4.9 inches
    fig_y = Emu(1900000)

    w = max_w
    h = int(w / ratio)
    if h > max_h:
        h = max_h
        w = int(h * ratio)

    fig_x = Emu(838200) + (max_w - w) // 2  # center horizontally

    slide.shapes.add_picture(image_path, fig_x, fig_y, w, h)

    add_badge_and_citation(slide, study_name, citation)
    return slide
```

---

## Safety / AE Table

Same pattern as demographics table — use `add_table()` with arm-colored headers, sorted AE rows, and **the same font size minimums** (14pt minimum for all cell text, centered horizontally on slide).

---

## Conclusion Slide

```python
def build_conclusion(prs, study_name, citation, key_finding, bullets):
    slide = prs.slides.add_slide(LAYOUTS['Title and Content'])
    slide.placeholders[0].text = "Conclusion"

    tf = slide.placeholders[1].text_frame
    tf.clear()

    # Key finding as first bold bullet
    p = tf.paragraphs[0]
    p.text = key_finding
    p.level = 0
    p.font.bold = True
    p.font.name = FONT
    p.font.size = Pt(18)  # MINIMUM 18pt for body text

    # Remaining bullets
    for b in bullets:
        pb = tf.add_paragraph()
        pb.text = b["text"]
        pb.level = 0
        pb.font.name = FONT
        pb.font.size = Pt(18)  # MINIMUM 18pt for body text
        if b.get("bold"):
            pb.font.bold = True
        for sub in b.get("sub", []):
            ps = tf.add_paragraph()
            ps.text = sub
            ps.level = 1
            ps.font.name = FONT
            ps.font.size = Pt(14)  # Sub-bullets: 14pt minimum

    # Optional: highlight box behind key finding (like model)
    # Add a filled rectangle behind the first bullet area
    highlight = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Emu(838200), Emu(1700000), Emu(10515600), Emu(600000)
    )
    highlight.fill.solid()
    highlight.fill.fore_color.rgb = NAVY
    highlight.line.fill.background()
    # Move to back so text shows on top
    # python-pptx: move shape element to front of spTree (renders behind later shapes)
    sp_tree = slide.shapes._spTree
    sp_tree.remove(highlight._element)
    sp_tree.insert(2, highlight._element)  # insert near front (behind placeholders)

    # Overlay text on highlight
    ht = slide.shapes.add_textbox(Emu(900000), Emu(1720000), Emu(10400000), Emu(560000))
    hp = ht.text_frame.paragraphs[0]
    hp.text = key_finding
    hp.font.size = Pt(13)
    hp.font.color.rgb = WHITE
    hp.font.name = FONT
    hp.font.bold = False

    add_badge_and_citation(slide, study_name, citation)
    return slide
```

---

## Save

```python
prs.save("/home/claude/output.pptx")
```

Then run Phase 4 QA from SKILL.md.
