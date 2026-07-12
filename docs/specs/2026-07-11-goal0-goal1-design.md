# Presentations Project — Goal 0 (Consolidation) + Goal 1 (QA Checklist)

## Context

Allan maintains two custom Claude skills that build Moffitt-branded, fully editable PPTX decks:
`academic-paper-to-pptx` (paper → deck) and `pptx-to-pptx` (restyle existing deck). `spec.md` in the
Presentations folder defines three enhancements (QA checklist → Presenton MCP → template extraction).

Exploration findings that motivate a new **Goal 0** before Goal 1:

- The skills are installed **only in the Claude desktop app** (session cache under
  `~/Library/Application Support/Claude/local-agent-mode-sessions/skills-plugin/b47439e7-*/21d66937-*/skills/`),
  **not** in Claude Code (`~/.claude/skills/`). Their SKILL.md files reference desktop-sandbox paths
  (`/mnt/skills/public/pptx/...`) that don't exist locally.
- Sources have diverged: Dropbox has only an **older** `academic-paper-to-pptx.skill` (Apr 26, missing
  `template.pptx` + logo); `pptx-to-pptx` source exists **only as zips in ~/Downloads**; the installed
  desktop copies (Apr 29–30) appear newest. No git anywhere.
- The desktop-app skills dir is an internal cache — the durable update path is re-uploading a packaged
  `.skill` zip through the app UI, not writing into Application Support.

Decisions made with Allan during brainstorming:

1. **Canonical home**: this Presentations folder becomes a git repo → remote
   `https://github.com/allanatal/Presentation.git`; skills sync to **both** runtimes
   (live copy in `~/.claude/skills/`, packaged `dist/*.skill` zips for desktop-app upload).
2. **Scope now**: Goal 0 + Goal 1 in full detail. Goals 2/3 unchanged in spec. New **Goal 2.5**
   (Nano Banana Pro conceptual figures) is added to spec.md as design-only (requirements below).
3. **QA depth**: filled per-deck checklist **plus** automated numeric cross-check script.
4. **Architecture**: Approach A — monorepo + sync script + one shared QA module serving both skills.

## Target repo layout

```
Presentations/                     (git repo → github.com/allanatal/Presentation.git)
├── spec.md                        (updated: add Goal 0 status + Goal 2.5; mark decisions)
├── CLAUDE.md                      (project instructions + "## Lessons Learned / Do Not Repeat Log")
├── .gitignore                     (dist/, ~$*, .DS_Store, venv/, output decks)
├── skills/
│   ├── academic-paper-to-pptx/
│   │   ├── SKILL.md               (ported paths + tightened QA phase)
│   │   ├── references/
│   │   │   ├── template.pptx      (Moffitt master — from newest installed copy)
│   │   │   ├── slide-builders.md
│   │   │   ├── style-spec.md
│   │   │   ├── qa-checklist.md    (NEW — shared checklist spec, Goal 1)
│   │   │   └── Picture_3.x-wmf    (Moffitt logo)
│   │   └── scripts/
│   │       ├── extract_figures.py
│   │       └── qa_crosscheck.py   (NEW — shared by both skills, Goal 1)
│   └── pptx-to-pptx/
│       ├── SKILL.md               (ported paths + tightened QA phase)
│       └── scripts/ (parse_pptx.py, build_from_parsed.py, convert_pptx.py)
├── tools/
│   └── sync-skills.sh             (NEW — copy → ~/.claude/skills/ + build dist zips)
├── dist/                          (generated .skill zips, gitignored)
└── docs/
    └── specs/2026-07-11-goal0-goal1-design.md   (design doc, written at implementation start)
```

## Goal 0 — Consolidate & port (do first)

1. **Seed canonical sources with reconciliation.** Copy the newest installed desktop copies as
   baseline, then diff every text file against the Dropbox `.skill` (Apr 26) and the Downloads zips
   (`academic-paper-to-pptx.zip` Apr 29, `files_2/academic-paper-to-pptx-v2.zip`, `pptx-to-pptx.zip`
   Apr 30, `pptx-to-pptx_old.zip`). Sizes fluctuate non-monotonically (e.g. Dropbox slide-builders.md
   is larger than newer copies), so produce a short **diff report** and surface anything that looks
   like lost content to Allan before finalizing. Never assume newest == most complete.
2. **Port SKILL.md files and scripts to be environment-agnostic.** (Full inventory from reading
   every file — these are ALL the sandbox-isms:)
   - Both SKILL.md files: `/mnt/skills/public/pptx/...` (QA helpers, thumbnail.py, soffice.py) and
     `/mnt/skills/user/academic-paper-to-pptx/references/...` (shared refs in pptx-to-pptx), plus
     `/home/claude/...` output paths in code examples. Replace with "relative to this skill's root"
     phrasing + a short environment table (Claude Code: `~/.claude/skills/<name>/`; desktop app:
     `/mnt/skills/user/<name>/`), and cwd-relative output paths.
   - `pptx-to-pptx/scripts/convert_pptx.py`: hardcoded `SOFFICE = /mnt/skills/public/.../soffice.py`;
     defaults `--output /home/claude/output.pptx`, `--images-dir /home/claude/source_images`.
     → call `soffice` from PATH (fallback to the sandbox helper if it exists), cwd-relative defaults.
   - `pptx-to-pptx/scripts/build_from_parsed.py`: hardcoded `TEMPLATE` and `LOGO_PATH` under
     `/mnt/skills/user/...` → resolve relative to the script's own location
     (`../../academic-paper-to-pptx/references/`), overridable via env var or CLI flag.
   - `pptx-to-pptx/scripts/parse_pptx.py`: `--images-dir` default `/home/claude/source_images` → cwd.
   - `academic-paper-to-pptx/scripts/extract_figures.py`: shells out to ImageMagick `identify` for
     dimensions → switch to Pillow (already a required dep) so ImageMagick isn't needed locally.
   - QA rendering in both SKILL.md files: plain `soffice --headless --convert-to pdf` + `pdftoppm`;
     note the LibreOffice/poppler dependency.
3. **Local dependency check** (report, don't silently install): `python3` + `python-pptx`,
   `pdfplumber`, `markitdown`, `Pillow`; `soffice` (LibreOffice) and `pdftoppm` (poppler) on PATH.
   Install what's missing with Allan's confirmation (brew / pip).
4. **`tools/sync-skills.sh`**: rsync both skill dirs → `~/.claude/skills/`; zip each into
   `dist/<name>.skill`; print installed paths + zip paths. Also **archive** (don't delete) the stale
   Dropbox `academic-paper-to-pptx.skill` by replacing it with a pointer note or moving it into the
   repo's `docs/legacy/` — the repo becomes the only editable source.
5. **git init + first commits + push** to `https://github.com/allanatal/Presentation.git`
   (`main`; commit per logical step: seed, port, sync tooling).

## Goal 1 — QA checklist + auto cross-check

### Claims data (the backbone)

Both skills persist what they extracted into a common JSON next to the output deck:

- `academic-paper-to-pptx`: Phase 1c now **also writes** `<deckname>_claims.json` — every number
  destined for slides: `{id, category (efficacy|safety|demographics|design|citation|other),
  label ("mPFS FOLFOX arm"), value ("6.9 months"), numeric_tokens, source_hint (page/table),
  slide (filled in Phase 3)}` + deck metadata (source file, NCT, title, date).
- `pptx-to-pptx`: `parse_pptx.py`'s existing parsed-source JSON serves as the claims source
  (source deck text is ground truth); same cross-check runs in "pptx mode". (Verified by reading the
  script: it already emits per-slide `full_text`, table `rows`, notes, and positions — exactly what
  the cross-check needs; no rework required, only the path defaults.)

### `scripts/qa_crosscheck.py` (shared, deterministic — no LLM judgment)

Inputs: built `.pptx` + claims JSON (+ mode flag). Outputs: `<deckname>_QA.md`, non-zero exit if any ❌.

1. Extract all text from the built deck via python-pptx (incl. tables, group shapes, notes).
2. Tokenize numeric patterns: HR/OR/RR + CI, p-values, medians + units, percentages, N, NCT ids,
   doses/schedules.
3. **Two-way diff**: (a) each claim's value found on its assigned slide → ✅/❌;
   (b) every numeric token in the deck traced back to a claim → orphans flagged ❌
   ("number in deck but not in source"); unresolved `[CHECK]` placeholders flagged ❌.
4. **Font-floor audit**: walk every run; flag below floors (body 18 / sub-bullet 16 / table 14 /
   diagram 12 / key-message 14; citation 10 and badge 11 as exceptions) → ⚠️/❌ list.
5. **Editability audit (pptx-to-pptx mode)**: per slide, count editable shapes vs pictures; ❌ any
   slide that is a single full-slide image; list each rasterized shape with its justification
   (chart / complex group).
6. Emit `<deckname>_QA.md`: sign-off block (deck, source, date, reviewer) then sections mirroring
   spec.md's checklist (Data accuracy, Structural/labeling, Citation integrity, Editability,
   Legibility/layout, Branding), each line auto-populated with the actual claim + verdict. Items the
   script can't verify mechanically (arm labels not swapped, endpoint primary/secondary tagging,
   citation authenticity, branding presence) are emitted as explicit **manual-review lines** for
   Allan's eyeball pass.
7. Optional `--notes` flag writes per-slide QA lines into speaker notes (default **off**; the QA.md
   file is the artifact of record).

### `references/qa-checklist.md` (shared reference)

Defines the checklist template/sections, the claims-JSON schema, how builders assign slide numbers,
and the rule: **a deck is not final until its `_QA.md` has zero ❌ and all manual lines are checked.**

### SKILL.md QA-phase updates (both skills)

Phase 4 becomes: render-and-inspect (existing) → write claims JSON (paper skill) → run
`qa_crosscheck.py` → fix ❌ and re-run → deliver deck + `_QA.md` together. Both files instruct
reading `qa-checklist.md` during QA.

## Spec.md updates (same commit series)

- Insert **Goal 0** (consolidation/port) as completed groundwork; record the four brainstorm decisions.
- Insert **Goal 2.5 — Optional AI figure generation (Nano Banana Pro / gemini-3-pro-image-preview)**,
  after Goal 2, before Goal 3, capturing Allan's requirements verbatim: conceptual/illustrative
  figures only (MoA cartoons, pathway schematics, anatomical illustrations, visual abstracts),
  embedded editable via existing `add_picture()` path; **hard guardrail: never generate data figures**
  (no KM/forest/waterfall/swimmer/any chart with real numbers — data figures come only from source
  extraction or programmatic build); every AI image tagged as AI-generated and flagged in the Goal 1
  checklist with an explicit "contains no data — decorative/conceptual only" line; access via Gemini
  API `GEMINI_API_KEY` (paid tier ~$0.13/image at 1–2K so outputs carry only invisible SynthID, not
  the visible watermark; free Pro tier ≈3 low-res/day is insufficient); cloud service ⇒ no PHI, no
  unpublished patient data in prompts; helper `scripts/generate_figure.py` (prompt → PNG in figures
  dir → path), strictly **opt-in per figure**, never auto-generates. Design-only for now.

## Verification (end-to-end)

1. `tools/sync-skills.sh` runs clean; both skills visible to Claude Code from `~/.claude/skills/`.
2. **Paper round-trip**: build a deck from a public open-access oncology RCT PDF end-to-end locally;
   confirm editable PPTX opens; `_claims.json` + `_QA.md` produced and populated with real values.
3. **Negative control**: deliberately corrupt one HR in a test build → cross-check must flag ❌
   (mismatch) and an injected orphan number → ❌; a sub-floor font run → flagged.
4. **pptx-to-pptx round-trip**: restyle a sample deck; editability audit lists shapes correctly and
   rejects a full-slide-image slide in a synthetic test.
5. Commit + push each landed step;
   `rm -f .git/index.lock && git add -A && git commit -m "…" && git push` pattern (Dropbox lock issue).
6. Per global rules: run every script end-to-end before reporting completion; print all output paths.

## Out of scope now

Goals 2 (Presenton), 2.5 (implementation), 3 (template extraction); plugin repackaging; any PHI-bearing
test content (public papers only).
