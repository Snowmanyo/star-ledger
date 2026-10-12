#!/bin/sh
# 本機模擬測試：載入真正的 Code.gs，假的試算表、LINE 與 Gemini。用法：sh tests/run.sh（需要 seed-backup.json）
cd "$(dirname "$0")/.."
for t in tests/test*.js; do
  cat tests/harness.js "$t" > "${TMPDIR:-/tmp}/sl-run.js"
  printf '%s ' "$t"; node "${TMPDIR:-/tmp}/sl-run.js" apps-script/Code.gs seed-backup.json 2>&1 | grep -E 'ALL PASSED|Error' | head -1
done
