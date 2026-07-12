#!/bin/bash
# Sync the canonical skill sources (this repo) to both runtimes:
#   1. ~/.claude/skills/            — Claude Code reads these directly
#   2. dist/<name>.skill zip        — upload manually to the Claude desktop app
#
# Run after any change under skills/. The repo is the single source of truth;
# never edit the installed copies directly.
set -euo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SKILLS_SRC="$REPO_DIR/skills"
CLAUDE_SKILLS="$HOME/.claude/skills"
DIST="$REPO_DIR/dist"

mkdir -p "$CLAUDE_SKILLS" "$DIST"

for skill_dir in "$SKILLS_SRC"/*/; do
  name="$(basename "$skill_dir")"
  echo "→ $name"
  rsync -a --delete "$skill_dir" "$CLAUDE_SKILLS/$name/"
  echo "   installed  $CLAUDE_SKILLS/$name/"

  # Desktop-app package: SKILL.md at the archive root (same layout as .skill exports)
  zip_path="$DIST/$name.skill"
  rm -f "$zip_path"
  (cd "$skill_dir" && zip -qr "$zip_path" . -x ".*" -x "__pycache__/*")
  echo "   packaged   $zip_path"
done

echo
echo "Done. To update the Claude desktop app, re-upload the dist/*.skill files via the app UI."
