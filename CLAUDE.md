# Presentations — Moffitt Deck-Generation Skills

Canonical source for the `academic-paper-to-pptx` and `pptx-to-pptx` Claude skills and
their shared QA tooling. Remote: https://github.com/allanatal/Presentation.git

## Ground rules

- **This repo is the ONLY editable source.** Installed copies (`~/.claude/skills/`, Claude
  desktop app) are downstream. After ANY change under `skills/`, run `tools/sync-skills.sh`
  and remind Allan to re-upload `dist/*.skill` to the desktop app.
- `spec.md` is the project spec and decision log — keep its status lines current.
- Scientific integrity is the point of this project: never weaken the data-fidelity rule
  (no invented numbers, `[CHECK]` placeholders), the font-size floors, the editability
  principle, or the QA zero-❌ rule. They are hard constraints, not preferences.
- PHI: patient-adjacent files are processed locally only; cloud services (incl. any future
  Nano Banana Pro figures, Presenton with Anthropic key) get de-identified/public content only.
- Test decks use public/open-access papers only.

## Layout

- `skills/academic-paper-to-pptx/` — paper → Moffitt deck (template.pptx + builders + QA script)
- `skills/pptx-to-pptx/` — restyle existing deck (shares the sibling skill's references)
- `tools/sync-skills.sh` — install to `~/.claude/skills/` + package `dist/*.skill`
- `docs/specs/` — design docs; `docs/legacy/` — superseded lineages + reconciliation report
- QA workflow: `skills/academic-paper-to-pptx/references/qa-checklist.md` (claims schema,
  cross-check usage, zero-❌ rule)

## Verification habits

- Run `qa_crosscheck.py` self-test before trusting changes to it: clean synthetic deck must
  exit 0; a deck with an injected wrong value / orphan number / `[CHECK]` / sub-floor font /
  full-slide image must flag all five.
- Render QA via `soffice --headless --convert-to pdf` + `pdftoppm` (installed via brew).

## Lessons Learned / Do Not Repeat Log

- 2026-07-11 — Font audit initially read only `run.font.size` and silently missed
  paragraph-level sizes (`p.font.size`), which is how slide-builders.md actually sets fonts.
  Caught only because the negative-control test injected a sub-floor font. Always check
  effective formatting (run → paragraph fallback) in python-pptx, and always run negative
  controls on QA tooling before trusting it.
- 2026-07-11 — The skills had THREE diverged copy locations (desktop-app install, Dropbox
  archive, Downloads zips) with non-monotonic file sizes; the Dropbox copy was a different
  architecture lineage (PptxGenJS), not an older version of the same files. Never assume
  newest timestamp == same lineage; diff before consolidating. The CONSORT builder was lost
  in the PptxGenJS→python-pptx rewrite (preserved in docs/legacy, candidate future port).
- 2026-07-11 — Pre-existing bug found by the new QA cross-check on its first real run:
  `build_from_parsed.py` used `1 in slide.placeholders`, which is ALWAYS False in python-pptx
  (membership iterates shape objects, not idx ints), silently dropping every content slide's
  bullets and every title-slide subtitle. Fixed with `placeholder_by_idx()`; the same trap was
  in SKILL.md example code. Never use `idx in slide.placeholders`; and never trust a
  conversion pipeline that hasn't been checked by content diffing.
- 2026-07-12 — Deck-spec→builder refactor landed. Two environment/tooling notes: (1) `build/`
  is intentionally gitignored ("repo is about skills, not decks") — the DRAGON regression proof
  (`DRAGON01_deck_spec.json`, built pptx, QA) lives there as a LOCAL artifact; don't `git add`
  it. Durable record goes in spec.md + the tracked `tests/fixtures/mini_spec.json`. (2) This
  Mac has no `python` on PATH, only `python3`; and `sync-skills.sh`'s zip exclude `__pycache__/*`
  did NOT match nested `scripts/__pycache__/` (zip -x matches the full member path) — fixed to
  `*__pycache__*` / `*.pyc` / `scripts/tests/*` so dev artifacts don't ship in `dist/*.skill`.
- 2026-07-12 — Design pattern worth keeping: making font floors a STRUCTURAL guarantee. The
  deck spec has no font-size field at all, so a sub-floor font is unrepresentable; sizes live
  only in `build_deck.py` (all clamped via `pt(size, floor)`). Prefer designing a whole bug
  class out of the input format over re-checking for it downstream.
