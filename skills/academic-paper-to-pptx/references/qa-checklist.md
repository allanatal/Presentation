# Deck QA Checklist — Shared Reference

Used by **both** `academic-paper-to-pptx` and `pptx-to-pptx` during their QA phase.
The rule: **a deck is not final until its `<deckname>_QA.md` shows zero ❌ and every
manual-review line has been checked by the reviewer.** The deck and its QA file are
delivered together.

## How QA works

1. During the build, the skill writes a **claims file** next to the deck (schema below).
2. After the build, run the shared cross-check script:

```bash
python "<academic-paper-to-pptx skill root>/scripts/qa_crosscheck.py" deck.pptx \
    --claims deck_claims.json --mode paper --source-text paper_text.md \
    --output deck_QA.md
# pptx-to-pptx restyles: --mode pptx --claims parsed.json (output of parse_pptx.py)
```

3. The script is **deterministic** (no model judgment). It verifies mechanically what it
   can and emits explicit `MANUAL` lines for what it cannot. Fix every ❌, re-run, then
   the reviewer works through the MANUAL lines against the source.

## Claims file schema (`<deckname>_claims.json`, paper mode)

Written during Phase 1 analysis and updated with slide numbers during Phase 3 build.
**Every number destined for a slide must be a claim.** Never invent values — claims are
extracted verbatim from the source paper.

```json
{
  "deck": {
    "title": "Deck title",
    "source": "paper.pdf",
    "nct": "NCT01234567",
    "date": "2026-07-11"
  },
  "claims": [
    {
      "id": "eff-01",
      "category": "efficacy",
      "label": "Median OS, experimental arm",
      "value": "14.1 months (95% CI 12.8-15.9)",
      "slide": 11,
      "source_hint": "Table 2 / p. 8"
    }
  ]
}
```

- `category`: `efficacy` | `safety` | `demographics` | `design` | `citation` | `other`
- `value`: the claim text exactly as it appears in the source (numbers verbatim)
- `slide`: 1-based slide number where the claim appears (fill during build; `null` until placed)
- `source_hint`: where in the paper the reviewer can verify it

In **pptx mode**, the claims source is `parse_pptx.py`'s JSON of the source deck — the
cross-check compares numeric content slide-by-slide (source deck is ground truth).

## What the script verifies automatically

| Check | Verdict basis |
|---|---|
| Every claim's numbers appear on its assigned slide | ❌ if any numeric token missing |
| No orphan numbers (in deck but not in source text/claims) | ❌ per orphan token |
| No unresolved `[CHECK]` placeholders | ❌ per occurrence |
| Font-size floors (body 18 / sub-bullet 16 / table 14 / diagram 12; citation 10 & badge 11 exempt) | ❌ per violating text run |
| Editability: no full-slide screenshot; pictures listed per slide | ❌ full-slide image; pictures listed for review |
| Citation textbox present on content slides | ⚠️ per slide missing one |

## What stays MANUAL (script lists these explicitly)

- Treatment-arm labels not swapped (numbers can match while labels are crossed).
- Endpoints tagged correctly primary vs secondary.
- Biomarker cutoffs/thresholds semantics (e.g., CLDN18.2, HER2, MSI-H, PD-L1 CPS/TPS).
- Drug names, doses, schedules read correctly in context.
- Citations trace to real references in the paper (no fabrication).
- Figures at correct aspect ratio; no visual overflow (also eyeball rendered slides).
- Moffitt branding present on every slide (template inheritance).

## Reviewer sign-off block

Each `_QA.md` starts with:

```
Deck: <file> | Source: <paper/deck> | Built: <date>
Reviewer: ______________  Date: ______________  Verdict: ______________
```

The QA file doubles as the audit trail — keep it with the deck.
