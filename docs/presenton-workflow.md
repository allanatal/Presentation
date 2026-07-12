# Presenton Workflow (Goal 2)

Presenton is a self-hosted, Apache-2.0 AI presentation generator that exports **editable
PPTX** and ships a **built-in MCP server**. In this project it is a *fast first-draft engine
for generic / public-data decks only* — never the accuracy-critical path. Draft in Presenton,
then finish and QA in the Moffitt pipeline (`pptx-to-pptx` → `qa_crosscheck.py`).

**Status:** stood up and round-trip-proven **2026-07-12** (Docker Desktop + host Ollama
`qwen2.5:32b`). Acceptance criteria in `spec.md` §Goal 2 all met.

> **Privacy (per spec §0):** cloud providers (Anthropic/OpenAI) get **de-identified / public
> content only**. True PHI never goes to Presenton at all; use the Ollama-local provider for
> anything more sensitive, and still keep real patient data out.

---

## 1. Prerequisites (one-time, this Mac)

- **Docker Desktop** — installed via `brew install --cask docker-desktop` (the cask needs a
  sudo password for CLI symlinks in `/usr/local/bin`; approve the prompt). Launch with
  `open -a Docker` and wait for the daemon (`docker info` succeeds).
- **Ollama** — already installed (`/usr/local/bin/ollama`, server on `:11434`); model
  `qwen2.5:32b` pulled. Only needed for the local provider.
- The `docker` CLI lives at `/usr/local/bin/docker`; prepend `export PATH="/usr/local/bin:$PATH"`
  if a shell can't find it.

## 2. Files & layout (outside Dropbox on purpose)

Presenton state lives in `~/presenton/` — **not** under Dropbox, to avoid sync churn on the
SQLite DB and export blobs (mirrors the "build artifacts are local" rule).

```
~/presenton/
  presenton.env          # provider config + admin creds (chmod 600; NOT in git)
  app_data/              # container bind-mount: DB, generated pptx under app_data/exports/
  out/                   # decks copied out for QA + the Moffitt reskin
```

## 3. Provider config — `~/presenton/presenton.env`

`CAN_CHANGE_KEYS=true` lets you switch providers/keys live in the UI (Settings), so all three
options below are available without recreating the container. Pick the active one via `LLM`.

```ini
# ── Active provider (choose one) ──────────────────────────────
LLM=ollama                                    # ollama | anthropic | openai
OLLAMA_URL=http://host.docker.internal:11434  # host Ollama, reachable from the container
OLLAMA_MODEL=qwen2.5:32b

# ── Anthropic (de-identified / public content only) ───────────
# LLM=anthropic
# ANTHROPIC_API_KEY=...        # export in terminal / paste in UI; never commit
# ANTHROPIC_MODEL=claude-sonnet-5

# ── OpenAI / ChatGPT key (de-identified / public content only) ─
# LLM=openai
# OPENAI_API_KEY=...
# OPENAI_MODEL=gpt-5

# ── Images & auth ─────────────────────────────────────────────
DISABLE_IMAGE_GENERATION=true   # Presenton still inserts decorative placeholder graphics
CAN_CHANGE_KEYS=true            # switch provider/keys in the UI
AUTH_USERNAME=allan
AUTH_PASSWORD=<generated; stored only in this file>
```

`host.docker.internal` is how the container reaches the Mac's Ollama; verified with
`docker exec presenton curl -s http://host.docker.internal:11434/api/tags`.

> **Port note:** the upstream docs use `-p 5000:80`, but macOS **AirPlay Receiver / Control
> Center already listens on :5000**. This project uses **5001** everywhere. (Free 5000 by
> turning off AirPlay Receiver in System Settings → General → AirDrop & Handoff if you prefer.)

## 4. Run the container

```bash
export PATH="/usr/local/bin:$PATH"
docker run -d --name presenton --restart unless-stopped \
  -p 5001:80 \
  --env-file ~/presenton/presenton.env \
  -v "$HOME/presenton/app_data:/app_data" \
  ghcr.io/presenton/presenton:latest
```

UI at <http://localhost:5001>. `--restart unless-stopped` brings it back after reboot /
Docker restart. Manage with `docker {logs,stop,start,rm -f} presenton`.

## 5. Wire the MCP server into Claude Code

Presenton exposes a streamable-HTTP MCP server at **`/mcp`**, bearer-auth. Get a token, then
register it:

```bash
# 1) log in (creates/uses the single admin from presenton.env)
source ~/presenton/presenton.env
curl -s -X POST http://localhost:5001/api/v1/auth/login \
  -H "Content-Type: application/json" \
  -d "{\"username\":\"$AUTH_USERNAME\",\"password\":\"$AUTH_PASSWORD\"}" \
  -o ~/presenton/.login_response.json
TOKEN=$(python3 -c "import json;print(json.load(open('$HOME/presenton/.login_response.json'))['access_token'])")

# 2) register with Claude Code (project scope)
claude mcp add --transport http presenton http://localhost:5001/mcp \
  --header "Authorization: Bearer $TOKEN"

claude mcp list        # -> presenton: http://localhost:5001/mcp (HTTP) - ✔ Connected
```

Tools advertised by the server (confirmed via `tools/list`):

| MCP tool | Purpose |
|---|---|
| `generate_presentation` | Generate a deck; returns the base URL of the PDF/PPTX. |
| `templates_list` | List built-in + custom templates. |

Notes:
- The bearer token is a JWT; if the MCP server later shows **auth errors**, re-run step 1–2 to
  refresh it (`claude mcp remove presenton` first).
- New MCP servers are **not hot-loaded** into a running Claude Code session — the tools appear
  after the session reloads. The REST API (below) works immediately and is the more reliable
  automation surface.

## 6. Generate a draft (REST — the proven path)

```bash
source ~/presenton/presenton.env
curl -u "$AUTH_USERNAME:$AUTH_PASSWORD" \
  -X POST http://localhost:5001/api/v1/ppt/presentation/generate \
  -H "Content-Type: application/json" \
  -d '{"content":"<public / de-identified brief>","n_slides":5,
       "template":"general","language":"English","export_as":"pptx"}' \
  -o ~/presenton/.gen_response.json
```

Response:

```json
{ "presentation_id": "…",
  "path": "/app_data/exports/<name>_<uuid>.pptx",
  "edit_path": "/presentation?id=…" }
```

`path` is a **container** path; because `app_data` is bind-mounted, the same file is on the host
at `~/presenton/app_data/exports/<name>_<uuid>.pptx` (or `docker cp "presenton:$path" .`).

**Timing:** ~6.5 min for a 5-slide deck on local `qwen2.5:32b` (Apple GPU). Cloud providers are
much faster. Presenton pulls a few HuggingFace layout models on first run.

## 7. Handoff into the Moffitt pipeline

The `pptx-to-pptx` skill reskins the Presenton draft into the Moffitt template. For Presenton
specifically:

1. **Parse** — `scripts/parse_pptx.py <draft>.pptx --output parsed.json`.
2. **Rebuild** — Presenton lays a deck out *spatially* (title on the right, headings/details in
   columns) and the generic heuristic role-classifier mislabels that. Reconstruct semantically:
   - `title` = the largest-font text on each slide;
   - `key message` = the sentence Presenton rendered largest (→ italic blue lead line);
   - `heading` = short labels with no terminal period; `detail` = the sentences;
   - pair each heading to its **nearest detail by position** → bold bullet + sub-bullet.
   Detect headings by **shape, not absolute font size** — Presenton sizes text inconsistently
   slide-to-slide (e.g. 15 pt headings on some slides, 13.5 pt on others), so a fixed pt
   threshold silently drops structure on the odd slide.
3. **Strip all images** — with `DISABLE_IMAGE_GENERATION=true`, Presenton still fills image slots
   with decorative placeholder graphics (dark striped rectangles + tiny icon glyphs). None carry
   information; drop them.
4. **QA** — `scripts/qa_crosscheck.py <out>.pptx --claims parsed.json --mode pptx`. Must be
   **0 ❌**. Numbers *added* by the reskin are hard fails; numbers *dropped* by trimming are
   warnings to review.

Reference build script for this handoff: `~/presenton/out/build_moffitt_from_presenton.py`
(local artifact, kept beside the decks). **2026-07-12 round-trip result:** 5-slide draft →
Moffitt reskin, 18 editable text shapes, 0 images, `qa_crosscheck` **9 ✅ / 0 ❌ / 1 ⚠**
(the ⚠ is the intentionally-dropped `[Your Name]` / date placeholder on the title slide).

## 8. What Presenton drafts well vs. what still needs the python-pptx path

**Good fit (draft in Presenton):**
- Narrative / educational / overview decks from a prompt or outline.
- Generic structure: title + key message + a few heading/detail bullet groups per slide.
- Public or de-identified content where wording matters and exact numbers don't.

**Do NOT use Presenton for (stay on the Moffitt `build_deck.py` / python-pptx path):**
- **Exact-data tables** (Table 1, efficacy/safety) — Presenton can't be trusted with precise
  HR/CI/median/AE values; those come only from the source paper.
- **Study-design diagrams & CONSORT** — the parameterized renderers in `build_deck.py`.
- **Any data figure** — KM curves, forest/waterfall/swimmer plots (see Goal 2.5 guardrail).
- **Content completeness** — the local model dropped "Phase IV" when asked for Phases 0–IV;
  Presenton drafts approximately, so always reconcile against the intended outline.

**Draft-quality caveats observed (local `qwen2.5:32b`):** occasional missing requested item,
inconsistent per-slide font sizing, and heading↔detail pairing that is only as good as
Presenton's spatial grouping (no semantic links in the file). All are handled/flagged in the
handoff above; cloud providers should reduce the content-omission issues.

## Quick reference

```bash
export PATH="/usr/local/bin:$PATH"
open -a Docker                                   # start daemon
docker start presenton                           # start Presenton (auto-restarts otherwise)
open http://localhost:5001                        # UI
docker logs --tail 20 presenton                  # troubleshoot
docker exec presenton curl -s http://host.docker.internal:11434/api/tags   # host-Ollama link
```
