#!/usr/bin/env bash
# Merge the latest upstream (singerla/pptx-automizer) into this fork.
#
#   bash tools/fork/sync-upstream.sh [branch]      (default branch: main)
#
# - AI-INSTRUCTOR.md is generated, so conflicts in it are resolved by
#   regenerating it, never by hand.
# - CHANGELOG.md and ROADMAP.md belong to the fork; conflicts in them keep
#   the fork's version.
# - The fork uses Bun; upstream uses Yarn. Upstream's yarn.lock files are
#   deleted again after the merge, and bun.lock is updated to match
#   upstream's package.json changes. Bun reads Yarn's `resolutions` as-is.
# - Stops on any other conflict (see FORK.md for the few lines the fork
#   touches in upstream files).
# - Runs the typecheck and the test suite at the end. Nothing is pushed.
set -euo pipefail

BRANCH="${1:-main}"
UPSTREAM_URL="https://github.com/singerla/pptx-automizer.git"

cd "$(git rev-parse --show-toplevel)"

if ! git diff --quiet || ! git diff --cached --quiet; then
  echo "Working tree has uncommitted changes - commit or stash them first." >&2
  exit 1
fi

git remote get-url upstream >/dev/null 2>&1 || git remote add upstream "$UPSTREAM_URL"
# Remember how conflicts were resolved, so a repeated conflict resolves itself.
git config rerere.enabled true

# Upstream's tags (v0.0.2–v0.9.3) would collide with the fork's own versions.
git config remote.upstream.tagOpt --no-tags
git fetch --no-tags upstream
git checkout "$BRANCH"

if ! git merge --no-edit upstream/main; then
  if git diff --name-only --diff-filter=U | grep -qx 'AI-INSTRUCTOR.md'; then
    git checkout --theirs AI-INSTRUCTOR.md
    git add AI-INSTRUCTOR.md
  fi
  # The fork keeps its own changelog (starting at 0.1.0) and roadmap.
  for own in CHANGELOG.md ROADMAP.md; do
    if git diff --name-only --diff-filter=U | grep -qx "$own"; then
      git checkout --ours "$own"
      git add "$own"
    fi
  done
  # Deleted in the fork (Bun), changed upstream: keep them deleted.
  for lock in yarn.lock website/yarn.lock; do
    if git diff --name-only --diff-filter=U | grep -qx "$lock"; then
      git rm -q "$lock"
    fi
  done
  remaining="$(git diff --name-only --diff-filter=U)"
  if [ -n "$remaining" ]; then
    echo "Resolve these conflicts, then 'git commit' and run this script again:" >&2
    echo "$remaining" >&2
    exit 1
  fi
  git commit --no-edit
fi

for lock in yarn.lock website/yarn.lock; do
  if [ -e "$lock" ]; then
    git rm -q "$lock"
  fi
done
if ! git diff --cached --quiet; then
  git commit -m "chore: drop upstream yarn.lock (the fork uses Bun)"
fi

bun install
bun install --cwd website
if ! git diff --quiet bun.lock website/bun.lock; then
  git add bun.lock website/bun.lock
  git commit -m "chore: update bun lockfiles after upstream merge"
fi

bun run docs:ai
if ! git diff --quiet AI-INSTRUCTOR.md; then
  git add AI-INSTRUCTOR.md
  git commit -m "docs: regenerate AI-INSTRUCTOR.md after upstream merge"
fi

bunx tsc --noEmit
bun run test

echo "Upstream merged into $BRANCH and tests pass. Push when ready."
