#!/usr/bin/env bash
set -euo pipefail

# Render outside the Work Drive mount to avoid macOS ._* sidecar issues.
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

LOCAL_BUILD_ROOT="${TMPDIR:-/tmp}/oca-cha-render"
mkdir -p "${LOCAL_BUILD_ROOT}"
WORKDIR="$(mktemp -d "${LOCAL_BUILD_ROOT}/build.XXXXXX")"

cleanup() {
  if [[ -n "${WORKDIR:-}" && -d "${WORKDIR}" ]]; then
    rm -rf "${WORKDIR}"
  fi
}
trap cleanup EXIT

echo "Staging repo into local workdir: ${WORKDIR}"
rsync -a --delete \
  --exclude '.git/' \
  --exclude '.cursor/' \
  --exclude '.quarto/' \
  --exclude '.DS_Store' \
  --exclude '._*' \
  "${REPO_ROOT}/" "${WORKDIR}/"

echo "Rendering Quarto book locally..."
(
  cd "${WORKDIR}"
  quarto render --to html
)

mkdir -p "${REPO_ROOT}/docs"
mkdir -p "${REPO_ROOT}/chapters/_generated/objects"

echo "Syncing generated outputs back to repo..."
rsync -a --delete --exclude '.DS_Store' --exclude '._*' "${WORKDIR}/docs/" "${REPO_ROOT}/docs/"
rsync -a --delete --exclude '.DS_Store' --exclude '._*' "${WORKDIR}/chapters/_generated/objects/" "${REPO_ROOT}/chapters/_generated/objects/"

echo "Done. Synced HTML output to ${REPO_ROOT}/docs"
