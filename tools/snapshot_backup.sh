#!/usr/bin/env bash
# Mirror the working tree (without .git) into "$ARENA_WORKSPACE"/_backup/repo-snapshot-latest/
set -euo pipefail
SRC=${1:-"$ARENA_WORKSPACE"/repo}
DST=${2:-"$ARENA_WORKSPACE"/_backup/repo-snapshot-latest}
mkdir -p "$DST"
tar -C "$SRC" --exclude=.git --exclude=dist -cf - . | tar -C "$DST" -xf -
echo "snapshot -> $DST ($(du -sh "$DST" | cut -f1))"
