#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(CDPATH= cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
SOURCE_DIR="$REPO_ROOT/docx-spec-html"
SKILLS_DIR="${JOYCODE_SKILLS_DIR:-$HOME/.joycode/skills}"
DEST_DIR="${1:-$SKILLS_DIR/docx-spec-html}"

if [[ ! -f "$SOURCE_DIR/SKILL.md" ]]; then
  printf 'ERROR: complete Skill source not found at %s\n' "$SOURCE_DIR" >&2
  exit 1
fi

for command_name in curl rsync; do
  if ! command -v "$command_name" >/dev/null 2>&1; then
    printf 'ERROR: %s is required to install the Skill.\n' "$command_name" >&2
    exit 1
  fi
done

mkdir -p "$DEST_DIR"
rsync -a --delete \
  --exclude '.venv/' \
  --exclude '__pycache__/' \
  --exclude '*.pyc' \
  "$SOURCE_DIR/" "$DEST_DIR/"

UV_BIN="$(command -v uv || true)"
if [[ -z "$UV_BIN" && -x "$HOME/.local/bin/uv" ]]; then
  UV_BIN="$HOME/.local/bin/uv"
fi
if [[ -z "$UV_BIN" ]]; then
  printf 'Installing uv and managed Python runtime...\n'
  curl -LsSf https://astral.sh/uv/install.sh | env UV_INSTALL_DIR="$HOME/.local/bin" sh
  UV_BIN="$(command -v uv || true)"
  if [[ -z "$UV_BIN" && -x "$HOME/.local/bin/uv" ]]; then
    UV_BIN="$HOME/.local/bin/uv"
  fi
fi
if [[ -z "$UV_BIN" ]]; then
  printf 'ERROR: uv installation completed but the executable was not found.\n' >&2
  exit 1
fi

"$UV_BIN" venv --clear --python 3.12 "$DEST_DIR/.venv"
SKILL_PYTHON="$DEST_DIR/.venv/bin/python"
"$UV_BIN" pip install --python "$SKILL_PYTHON" -r "$DEST_DIR/requirements.txt"

OFFICECLI_BIN="$(command -v officecli || true)"
if [[ -z "$OFFICECLI_BIN" && -x "$HOME/.local/bin/officecli" ]]; then
  OFFICECLI_BIN="$HOME/.local/bin/officecli"
fi
if [[ -z "$OFFICECLI_BIN" ]]; then
  printf 'Installing OfficeCLI from its official distribution...\n'
  curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
  OFFICECLI_BIN="$(command -v officecli || true)"
  if [[ -z "$OFFICECLI_BIN" && -x "$HOME/.local/bin/officecli" ]]; then
    OFFICECLI_BIN="$HOME/.local/bin/officecli"
  fi
fi
if [[ -z "$OFFICECLI_BIN" ]]; then
  printf 'ERROR: OfficeCLI installation completed but the executable was not found.\n' >&2
  exit 1
fi

OFFICECLI_BIN="$OFFICECLI_BIN" "$SKILL_PYTHON" "$DEST_DIR/scripts/preflight.py"
printf '\nInstalled docx-spec-html at %s\n' "$DEST_DIR"
printf 'Restart JoyCode so it reloads the Skill.\n'
