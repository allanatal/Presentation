# Deck-Spec Schema — `build_deck.py`

The model authors a compact **deck spec** (JSON); the committed builder
`scripts/build_deck.py` turns spec + claims into the deck. **Do not hand-author
python-pptx.** The spec carries structure + prose; numbers come from the claims file.

```bash
python "<skill-root>/scripts/build_deck.py" <deck>_deck_spec.json \
    --claims <deck>_claims.json --out <deck>.pptx \
    --resolved-claims <deck>_claims_resolved.json
```

The builder writes `<deck>_claims_resolved.json` with each claim's `slide` field
**auto-filled** from where it was placed. Phase 4 QA runs against that resolved file —
**you never maintain the `slide` field by hand.**

## Two integrity properties this design gives you for free

1. **No sub-floor fonts, ever.** The spec has **no font-size field**. Sizes live only in
   `build_deck.py` and are all ≥ the style-spec floors (body 18 / sub-bullet 16 / table 14 /
   diagram 12). The whole font-floor bug class is designed out.
2. **Numbers are never hand-typed into code.** Prose uses `{{claim-id}}` tokens replaced
   with the claim's exact value; tables inline values but **declare their claim IDs**, and
   `qa_crosscheck.py` verifies the claim's numbers actually landed on the slide. A mistyped
   table cell fails QA.

## Claim references — two ways to place a claim on a slide

- **`{{claim-id}}` token** inside any string → replaced by that claim's `value`, and the
  claim is assigned to that slide. Best for key messages and bullets:
  `"key_message": "Median OS {{eff-01}} vs {{eff-02}}; {{eff-03}}"`.
- **slide-level `"claims": ["dem-01", ...]` list** → assigns those claims to the slide; the
  slide's own inline text must contain their numbers (QA checks). Best for tables, CONSORT,
  study-design, and figure captions where numbers are woven into prose or cells.

**Contract:** every claim in the claims file must be referenced exactly once (token or list).
`build_deck.py` **errors** if any claim is never placed (override with `--allow-unreferenced`,
but the default is a hard stop — it is the "every claim assigned before QA" rule enforced).

## Top-level shape

```json
{
  "deck": {
    "study": "DRAGON-01",                     // top-right badge on every content slide
    "citation": "Yan C, et al. JAMA Oncol. 2026",  // bottom-right on every content slide
    "title": "...", "subtitle": "...",        // title slide (also usable per-slide)
    "authors": "...", "affiliation": "...", "date": "..."
  },
  "slides": [ { "type": "...", ... }, ... ]
}
```

## Slide types

Colors are palette names: `teal` (experimental), `blue` (2nd experimental), `gray`/`control`,
`navy`, `lightbg`. Unknown names fall back to teal.

### `title`
Usually just `{"type": "title"}` — inherits `title/subtitle/authors/affiliation/date` from
`deck`. Optional per-slide overrides of any of those, plus `title_size` (default 26).

### `bullets`
```json
{"type": "bullets", "title": "Clinical Background",
 "key_message": "optional italic subtitle under the title",
 "bullets": [
   {"text": "Level-0 point", "bold": true, "sub": ["level-1 sub-bullet", "..."]},
   {"text": "Another point"}
 ]}
```
Body = 18pt, sub-bullets = 16pt (fixed). Add `"claims": [...]` if a bullet carries claim
numbers (inline), or use `{{id}}` tokens inside `text`.

### `table`
```json
{"type": "table", "title": "Baseline Characteristics (mITT)",
 "key_message": "optional italic line above the table",
 "label_header": "Characteristic",
 "font": 15, "row_h": 352000, "top": 1900000,   // all optional
 "arms": [{"label": "IP (n=148)", "color": "teal"}, {"label": "PS (n=74)", "color": "gray"}],
 "rows": [
   {"label": "Age, median (range), y", "vals": ["60 (24-70)", "56 (23-74)"], "bold": true},
   {"label": "Sex", "hdr": true},                      // section header row (shaded, no vals)
   {"label": "  Female", "vals": ["68 (45.9)", "36 (48.6)"]}
 ],
 "note": "optional italic note textbox below the table",
 "claims": ["dem-00", "dem-01", "..."]}
```
2–3 arms. Table is auto-centered. Cell text ≥ 14pt (fixed). One value per arm, in arm order.
If a table would overflow, split across two `table` slides — never shrink below the floor.

### `figure`
```json
{"type": "figure", "title": "Primary Endpoint: Overall Survival",
 "image": "figures/km_os.png",                 // relative to --images-base (default: spec dir)
 "key_message": "Median OS {{eff-01}} vs {{eff-02}}; {{eff-03}}",
 "caption": ["line 1 of caption", "line 2 ..."],
 "max_h": 3150000, "fig_top": 1880000,          // optional
 "claims": ["eff-04", "eff-05", "eff-06"]}      // claims whose numbers are in the caption
```
Embeds the image with preserved aspect ratio. **Only source/extracted figures or
programmatic charts** — never a full-slide screenshot (QA flags it).

### `study_design`
```json
{"type": "study_design",
 "description": "one-line italic description",
 "eligibility": ["Age 18-75 y", "..."],
 "eligibility_notes": ["No stratification factors prespecified"],   // optional, gray italic
 "randomization": "2:1", "n": "N = 222",
 "arms": [
   {"name": "IP group (n=148)", "detail": "regimen text ...", "color": "teal"},
   {"name": "PS group (n=74)",  "detail": "regimen text ...", "color": "gray"}],
 "primary_endpoints": ["Overall survival"],
 "secondary_endpoints": ["Progression-free survival", "Safety"],
 "registration": "ChiCTR-IIR-16009802",         // or NCT number
 "footer": "duration / enrollment / cutoff line",
 "claims": ["des-01", "des-02", "..."]}
```
Draws eligibility box → R circle → arm boxes → endpoints box, with connectors. Text ≥ 12pt.
Supports 2–3 arms (spacing auto-computed).

### `consort`
```json
{"type": "consort",
 "assessed": "246 Assessed for eligibility",
 "excluded": ["8 Excluded", "6 Did not meet criteria; 2 declined"],
 "randomized": "238 Randomized (2:1)",
 "arms": [
   {"color": "teal",
    "allocated": ["158 Allocated to IP group", "IP + IV paclitaxel + S-1"],
    "not_received": ["10 Did not receive intervention", "declined IP port placement"],
    "received": "148 Received intervention (mITT)",
    "outcomes": ["Conversion surgery: 75 (63 R0, 12 R2)", "132 died", "..."]},
   {"color": "gray", "allocated": ["..."], "not_received": ["..."],
    "received": "74 Received intervention (mITT)", "outcomes": ["..."]}],
 "claims": ["con-01", "con-02", "..."]}
```
Standard CONSORT topology: assessed → (excluded side-box) → randomized → per-arm
allocated → not-received → received → outcomes, with connectors. 2-arm layout is exact;
≥3 arms spread evenly. For a non-standard flow, render it as a `figure` from a pre-made image.

## Worked minimal example

See `scripts/tests/fixtures/mini_spec.json` (+ `mini_claims.json`) for a complete 4-slide
deck exercising `title`, `bullets`, `table` (inline vals + `claims`), and `bullets` with
`{{id}}` tokens. Build it with:

```bash
python scripts/build_deck.py scripts/tests/fixtures/mini_spec.json \
    --claims scripts/tests/fixtures/mini_claims.json --out /tmp/mini.pptx \
    --resolved-claims /tmp/mini_resolved.json
```
