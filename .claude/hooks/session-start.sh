#!/bin/bash
set -euo pipefail

cd "${CLAUDE_PROJECT_DIR:-$(git rev-parse --show-toplevel)}"

branch=$(git rev-parse --abbrev-ref HEAD 2>/dev/null || echo "?")
before=$(git rev-parse --short HEAD 2>/dev/null || echo "?")

if ! git fetch --quiet origin main 2>/dev/null; then
  echo "[session-start] git fetch origin main 失敗（branch=${branch}, HEAD=${before}）"
  exit 0
fi

behind=$(git rev-list --count HEAD..origin/main 2>/dev/null || echo "?")
ahead=$(git rev-list --count origin/main..HEAD 2>/dev/null || echo "?")

dirty=""
if ! git diff --quiet 2>/dev/null || ! git diff --cached --quiet 2>/dev/null; then
  dirty=" / dirty"
fi

is_remote="false"
if [ "${CLAUDE_CODE_REMOTE:-}" = "true" ]; then
  is_remote="true"
fi

# 自動 pull 條件：雲端一律 pull；本機只在 working tree 乾淨時 pull
should_pull="false"
if [ "$is_remote" = "true" ]; then
  should_pull="true"
elif [ -z "$dirty" ]; then
  should_pull="true"
fi

if [ "$should_pull" = "true" ]; then
  if git pull --ff-only --quiet origin main 2>/dev/null; then
    after=$(git rev-parse --short HEAD)
    if [ "${before}" = "${after}" ]; then
      echo "[session-start] 已是最新（branch=${branch}, HEAD=${after}${dirty}）"
    else
      echo "[session-start] fast-forward 至 main（branch=${branch}, ${before} → ${after}${dirty}）"
    fi
    exit 0
  fi
  echo "[session-start] 無法 fast-forward ${branch} 到 origin/main（已分歧）。HEAD=${before}, 落後 ${behind} / 本地多 ${ahead}${dirty}"
  exit 0
fi

# 本機 + dirty：只顯示狀態，不動 working tree
echo "[session-start] 本機 session：branch=${branch}, HEAD=${before}, 落後 main ${behind} / 本地多 ${ahead}${dirty}（未自動 pull）"
