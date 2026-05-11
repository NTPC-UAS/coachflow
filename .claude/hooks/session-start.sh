#!/bin/bash
set -euo pipefail

# 只在 Claude Code 雲端 session 跑。本機 session 可能有未 commit 的改動，
# 自動 pull 有風險，跳過。
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

cd "${CLAUDE_PROJECT_DIR:-$(git rev-parse --show-toplevel)}"

branch=$(git rev-parse --abbrev-ref HEAD 2>/dev/null || echo "?")
before=$(git rev-parse --short HEAD 2>/dev/null || echo "?")

if ! git fetch --quiet origin main 2>&1; then
  echo "[session-start] git fetch origin main 失敗（branch=${branch}, HEAD=${before}）"
  exit 0
fi

if git pull --ff-only --quiet origin main 2>/dev/null; then
  after=$(git rev-parse --short HEAD)
  if [ "${before}" = "${after}" ]; then
    echo "[session-start] 已是最新（branch=${branch}, HEAD=${after}）"
  else
    echo "[session-start] fast-forward 至 main（branch=${branch}, ${before} → ${after}）"
  fi
else
  echo "[session-start] 無法 fast-forward ${branch} 到 origin/main（可能已分歧或有未 commit 變更）。HEAD=${before}"
fi
