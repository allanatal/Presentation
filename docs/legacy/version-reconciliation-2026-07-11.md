# Skill Version Reconciliation — 2026-07-11

Before seeding this repo, five pre-existing copies/archives of the two skills were compared.
The repo's `skills/` tree was seeded from the **installed Claude desktop-app copies**
(`~/Library/Application Support/Claude/local-agent-mode-sessions/skills-plugin/…/skills/`,
dated 2026-04-29/30), which byte-for-byte match the newest packaged zips.

## Copies compared

| Copy | Date | Verdict |
|------|------|---------|
| Installed desktop-app copies (both skills) | Apr 29–30 | **Baseline — newest, used to seed repo** |
| `~/Downloads/academic-paper-to-pptx.zip` | Apr 29 | Identical to baseline (all 6 files) |
| `~/Downloads/pptx-to-pptx.zip` | Apr 30 | Identical to baseline (all 4 files) |
| `~/Downloads/files_2/academic-paper-to-pptx-v2.zip` | Apr 26–27 | Older python-pptx version; baseline adds font-size minimums table + table-centering rules. Nothing unique. |
| `~/Downloads/pptx-to-pptx_old.zip` | Apr 29 | Older SKILL.md (657 changed lines), identical scripts. Nothing unique. |
| Dropbox `Claude_projects/academic-paper-to-pptx.skill` | Apr 26 | **Different lineage**: original PptxGenJS-based architecture (no template.pptx; branding drawn in code). Superseded by the python-pptx + template approach. |

## Content flagged as potentially lost

- **CONSORT diagram builder**: the PptxGenJS-era `slide-builders.md` (Dropbox `.skill`) contains a
  full "CONSORT Diagram" builder section. The current `slide-builders.md` has **no CONSORT
  pattern**, even though the current SKILL.md presentation structure still lists
  "Slide 9 — CONSORT Diagram (if patient flow data available)". If a CONSORT slide is ever needed,
  the old PptxGenJS code (preserved in `docs/legacy/academic-paper-to-pptx-pptxgenjs.skill`) can be
  translated to python-pptx. Logged as a candidate future enhancement — not ported in Goal 0/1.

Nothing else unique to the older copies was found; all other differences are strictly superseded
content.

## Disposition

- Repo `skills/` = canonical source going forward (single source of truth, git-versioned).
- The stale Dropbox `.skill` is preserved at `docs/legacy/academic-paper-to-pptx-pptxgenjs.skill`
  and removed from `Claude_projects/` root to prevent accidental editing.
- `Claude_projects/academic-paper-to-pptx.zip` was a **byte-identical duplicate** of that `.skill`
  file (confirmed with `cmp`) and was deleted.
- Downloads zips left untouched (user's own files) but are now redundant.
