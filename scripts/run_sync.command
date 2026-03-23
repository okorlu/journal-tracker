#!/bin/zsh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"

cd "$REPO_DIR"

echo "Running journal sync..."
.venv/bin/journal-tracker-sync \
  --profile config/profiles/turkish-politics-starter.json \
  --workbook data/turkish_politics_articles_database.xlsx

echo
echo "Done. Press Enter to close."
read -r
