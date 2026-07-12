# spec.md — Deck-Generation Skill Enhancements

**Owner:** Allan (medical oncologist, GI malignancies — Moffitt Cancer Center)
**Working env:** Claude Code / VS Code + Claude Code
**Scope:** Three enhancements to my existing PowerPoint-generation skills, sequenced by value-to-effort.
**Status:** Draft — carry into Claude Code as starting context.

---

## 0. Context & Existing System

I already maintain two custom Claude skills that convert scientific content into
Moffitt-branded PowerPoint decks. This spec **extends** them; it does not replace them.
The existing architecture is good and stays intact.

### Skill 1 — `academic-paper-to-pptx`
- **Input:** research paper (PDF / DOCX / MD).
- **Output:** editable Moffitt-styled `.pptx`.
- **Approach:** template-based. A bundled `references/template.pptx` holds the Moffitt
  slide master (branding baked in). `python-pptx` opens it and adds slides via the
  master's layouts.
- **Four phases:** Ingest → Extract figures → Build → QA.
- **Data fidelity rule:** *never invent data*; every number comes from the paper;
  uncertain values get `[CHECK]` placeholders flagged to me.

### Skill 2 — `pptx-to-pptx`
- **Input:** an existing `.pptx` ("source").
- **Output:** the same content restyled into the Moffitt template.
- **Approach:** parse source **element-by-element**, classify each element's type/role,
  rebuild each as an editable object.
- **Editability Principle (mandatory):** never rasterize a whole slide. Only rasterize
  the *specific shape* that can't be recreated (chart objects `shape.has_chart`; complex
  group shapes >6 children / freeform / curved connectors). Everything else — titles,
  tables, textboxes, simple shapes — is rebuilt editable.

### Shared resources (both skills depend on these)
- `academic-paper-to-pptx/references/template.pptx` — Moffitt slide master.
- `academic-paper-to-pptx/references/slide-builders.md` — python-pptx code patterns.
- `academic-paper-to-pptx/references/style-spec.md` — colors, positions, font rules.

### Non-negotiable constraints (apply to everything below)
- **Editable PPTX output** — true shapes/text/tables, not slide images.
- **Font-size minimums** (hard floors; split slides rather than shrink):
  body 18pt · sub-bullets 16pt · table cells 14pt · diagram text 12pt ·
  key-message 14pt · citation 10pt · badge 11pt.
- **Scientific accuracy** — HRs, CIs, p-values, medians, ORR/DOR, AE %s, trial-arm
  labels, NCT numbers, citations must match the source exactly.
- **Privacy / PHI** — sensitive or patient-adjacent files are processed **locally only**.
  Cloud models/services are for de-identified or already-public content only
  (published abstracts, ASCO data, guideline summaries). Anything with real patient
  data goes through Moffitt-approved tooling with IT/privacy involvement. This
  constraint governs tool selection in Goals 2 and 3.

---

## Goal 1 — QA Checklist (DO THIS FIRST)

**Why first:** smallest lift, highest recurring value. Every deck I produce benefits
immediately. This is the "trust-but-verify" step that catches wrong hazard ratios,
misattributed endpoints, swapped treatment arms, bad p-values, and hallucinated
citations before a deck is considered final.

### Deliverable
A reusable QA-checklist artifact generated **alongside every deck**, plus a tightened
QA phase in both skills.

1. **`references/qa-checklist.md`** — a shared reference file (place it in
   `academic-paper-to-pptx/references/` so both skills can use it, mirroring how
   `slide-builders.md` and `style-spec.md` are already shared).
2. **Per-deck QA output** — when a deck is built, also emit a filled-in checklist
   (e.g. `<deckname>_QA.md`) that enumerates each slide and the specific claims to
   verify against the source. Optionally also write the checklist into PowerPoint
   **speaker notes** on a final "QA" slide or per-slide, so the verification travels
   with the file.

### Checklist content (organize into these sections)
- **Data accuracy (highest priority)**
  - Every HR / OR / RR with its CI and p-value matches source.
  - Median PFS/OS and other time-to-event numbers match.
  - ORR, DOR, DCR, and other response metrics match.
  - Table 1 demographics: each number matches exactly.
  - Safety/AE percentages match, per arm.
  - No number appears that isn't in the source (`[CHECK]` placeholders resolved).
- **Structural / labeling accuracy**
  - Treatment-arm labels correct and not swapped.
  - Endpoints correctly tagged primary vs secondary.
  - NCT number correct.
  - Biomarker cutoffs/thresholds correct (e.g., CLDN18.2, HER2, MSI-H, ctDNA, PD-L1 CPS/TPS).
  - Drug names, doses, schedules correct.
- **Citation integrity**
  - Every citation traces to a real source in the paper; no fabricated references.
  - Citation present on each content slide.
- **Editability (esp. for `pptx-to-pptx`)**
  - Titles, tables, textboxes, simple shapes are editable objects, not images.
  - Only true chart objects / complex groups are rasterized.
- **Legibility / layout**
  - No text below font-size minimums.
  - No overflow / cut-off text in tables or bullets.
  - Tables horizontally centered.
  - Figures embedded at correct aspect ratio.
- **Branding**
  - Moffitt branding present on every slide (inherited from template).
  - Study badge + citation on content slides.

### Design notes
- Make the checklist **actionable, not generic**: it should reference the *actual*
  extracted values so I can eyeball slide-vs-source quickly. Where the build step already
  extracted HR/CI/median/etc. into structured data, reuse that data to auto-populate the
  checklist lines.
- Keep a short **"reviewer sign-off"** block at the top (deck name, source, date,
  reviewer) so it doubles as an audit trail.

### Acceptance criteria
- Both skills read `qa-checklist.md` during their QA phase.
- Building a deck also produces a per-deck filled checklist file.
- The checklist explicitly lists the numeric claims to verify, pulled from the deck's
  own extracted data where available.

---

## Goal 2 — Presenton MCP Integration (DO THIS SECOND)

**Why:** Presenton is an open-source (Apache-2.0), self-hostable AI presentation
generator that exports **editable PPTX**, can build a reusable template **from my own
PPTX**, runs locally (Docker / desktop), works with my Anthropic key **or** a local
Ollama model, and ships a **built-in MCP server**. It's the best external match to my
needs and slots into the Claude Code / MCP workflow directly.

**Role in my workflow:** a *fast first-draft engine for generic / public-data decks* —
not a replacement for my accuracy-critical skill path. Draft in Presenton, then finish
and QA in my own pipeline.

### Tasks
1. **Stand up Presenton locally.**
   - Docker: `docker run -it --name presenton -p 5000:80 -v "./app_data:/app_data" ghcr.io/presenton/presenton:latest`
     (or the desktop app). Confirm the UI at `http://localhost:5000`.
   - Configure provider via env: for de-identified/public work, `LLM=anthropic` +
     `ANTHROPIC_API_KEY`; for maximum privacy, `LLM=ollama` with a local model
     (`OLLAMA_MODEL`, `START_OLLAMA`). Set `IMAGE_PROVIDER` appropriately (e.g. `pexels`
     with a key, or disable image generation).
   - Note the single-admin auth model (`AUTH_USERNAME` / `AUTH_PASSWORD`) — needed for API/MCP.
2. **Wire Presenton's MCP server into Claude Code.**
   - Add the Presenton MCP server to my Claude Code MCP config.
   - Verify Claude Code can trigger a generation and retrieve the resulting PPTX path.
3. **Prove the round-trip.**
   - Generate a short deck from a prompt / outline / markdown via the API or MCP:
     endpoint `POST /api/v1/ppt/presentation/generate` with `export_as: "pptx"`.
   - Confirm output opens as an **editable** PPTX.
4. **Define the handoff into my pipeline.**
   - Take a Presenton-drafted PPTX and run it through `pptx-to-pptx` to reskin into the
     Moffitt template — validating that the two systems compose cleanly.
   - Document which content types Presenton drafts well vs. what still needs my
     python-pptx path (study-design diagrams, exact-data tables).

### Guardrails
- **Privacy:** only feed Presenton de-identified/public content when using cloud models
  (Anthropic key). Reserve the Ollama-local configuration for anything more sensitive,
  and still keep true PHI out of it per the constraint above.
- Presenton supports charts/tables/images but they must be **specified during template
  creation**, not generated freely per slide — factor this into what I expect from it.

### Acceptance criteria
- Presenton runs locally and is reachable from Claude Code via MCP.
- I can generate a draft PPTX end-to-end from Claude Code.
- A Presenton draft successfully passes through `pptx-to-pptx` into the Moffitt template.
- A short `docs/presenton-workflow.md` records setup, env vars used, and the handoff steps.

---

## Goal 3 — Template-Extraction Enhancement (DO THIS THIRD)

**Why:** borrow Presenton's *AI-template-from-PPTX* idea and fold a version of it into my
own skill, so I'm not locked to the single bundled `template.pptx`. Goal: point the skill
at *any* branded PPTX (a new Moffitt template variant, a co-author's institutional
template, a journal/congress template) and have it extract a reusable style definition —
colors, typography, spacing, layout, logo/branding positions — that the builders honor.

**This is an enhancement to `academic-paper-to-pptx` (and, by sharing, `pptx-to-pptx`),
keeping the true-editability + local-execution advantages my current setup already has.**

### Tasks
1. **Template-extraction step.**
   - Given a source template `.pptx`, parse the slide master(s) and representative
     layouts to extract a **structured style spec**: theme colors, fonts + sizes per
     role, title formatting, bullet indents, footer/badge/citation positions, brand-bar
     geometry, logo placement.
   - Emit this as a machine-readable spec (e.g. `style-spec.json`) that generalizes the
     existing hand-authored `style-spec.md` constants.
2. **Make builders parameter-driven.**
   - Refactor `slide-builders.md` patterns (and any build scripts) to read positions,
     colors, and fonts from the extracted spec rather than hard-coded constants — while
     preserving the mandatory font-size floors as a hard override even if a template
     specifies smaller.
3. **Template registry.**
   - Support a small set of named templates (e.g. `moffitt-default`, plus any new ones)
     selectable at build time, each backed by its own extracted spec + template.pptx.
4. **Validation.**
   - Round-trip: extract a spec from the current Moffitt `template.pptx`, rebuild a known
     deck using the extracted spec, and confirm it matches the current hand-tuned output
     (visual + editability parity).

### Guardrails
- Font-size minimums remain **hard floors** regardless of what a template specifies.
- Extraction runs **locally**; templates may be institutional/unpublished.
- Don't regress the current Moffitt output — it's the reference standard.

### Acceptance criteria
- Skill can ingest an arbitrary branded PPTX and produce a reusable `style-spec.json`.
- Builders consume the extracted spec; the Moffitt round-trip matches current output.
- A new template can be added without editing builder code.

---

## Suggested Working Method in Claude Code
- Dedicated worktree/branch for this project; `/clear` between phases with this spec as
  the persistent external memory.
- Sequence strictly: **Goal 1 → Goal 2 → Goal 3.** Land and use Goal 1 before starting Goal 2.
- Keep each goal's changes reviewable in isolation.
- After each goal: run the QA checklist (Goal 1's own deliverable) on a real deck as the
  regression gate.

## Out of Scope (for now)
- Replacing python-pptx.
- Cloud SaaS deck tools (e.g. ChatSlide) for anything beyond de-identified/public drafts.
- Animations/transitions (output stays clean/static, per `pptx-to-pptx`).
- Any workflow that would route PHI through cloud models/services.
