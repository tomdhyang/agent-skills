#!/usr/bin/env bash
set -euo pipefail

# Link the skills I actually use into the directories each agent reads:
#   ~/.claude/skills  (Claude Code)
#   ~/.agents/skills  (Codex and other Agent Skills harnesses)
#
# BUCKETS lists the directories (relative to the repo) whose immediate
# children are skills. Add a line when vendoring a new source.
# Everything not listed (vendor in-progress/, misc/, deprecated/) stays
# in the repo but is not linked.
#
# Safe to re-run any time: stale links into this repo (or dangling links) are
# removed first, so a skill a vendor deletes disappears here after the next run.
# Real directories (not symlinks) in a destination are never touched.
#
# Run after: a fresh clone on a new machine, `git subtree pull`, or adding a skill.

REPO="$(cd "$(dirname "$0")/.." && pwd)"
BUCKETS=(
  skills/mine
  skills/vendor/mattpocock/skills/engineering
  skills/vendor/mattpocock/skills/productivity
)
DESTS=("$HOME/.claude/skills" "$HOME/.agents/skills")

srcs=()
for bucket in "${BUCKETS[@]}"; do
  for skill_md in "$REPO/$bucket"/*/SKILL.md; do
    [ -f "$skill_md" ] || continue
    srcs+=("$(dirname "$skill_md")")
  done
done

# A user-invoked skill must be user-invoked in both harnesses.
warn=0
for src in "${srcs[@]}"; do
  cc=0; cx=0
  grep -qiE '^disable-model-invocation:\s*(true|yes|on|1)' "$src/SKILL.md" && cc=1
  [ -f "$src/agents/openai.yaml" ] && grep -qiE '^\s*allow_implicit_invocation:\s*false' "$src/agents/openai.yaml" && cx=1
  if [ "$cc" != "$cx" ]; then
    echo "warn   $(basename "$src"): claude=$cc codex=$cx (disable-model-invocation vs agents/openai.yaml allow_implicit_invocation mismatch)" >&2
    warn=1
  fi
done

for DEST in "${DESTS[@]}"; do
  mkdir -p "$DEST"

  # Drop stale symlinks: anything pointing into this repo, or pointing nowhere.
  for entry in "$DEST"/*; do
    [ -L "$entry" ] || continue
    target="$(readlink "$entry")"
    case "$target" in
      /*) abs="$target" ;;
      *)  abs="$(cd "$DEST" && cd "$(dirname "$target")" 2>/dev/null && pwd)/$(basename "$target")" || abs="" ;;
    esac
    if [ ! -e "$entry" ] || [[ "$abs" == "$REPO"/* ]]; then
      rm "$entry"
    fi
  done

  linked=0
  for src in "${srcs[@]}"; do
    name="$(basename "$src")"
    target="$DEST/$name"
    if [ -e "$target" ] && [ ! -L "$target" ]; then
      echo "skip   $name  (real directory already at $target; move it away if you want the repo version)" >&2
      continue
    fi
    ln -sfn "$src" "$target"
    linked=$((linked + 1))
  done
  echo "$DEST: linked $linked skills"
done

[ "$warn" = 0 ] || echo "(see warnings above: fix the mismatch so both harnesses agree)" >&2
