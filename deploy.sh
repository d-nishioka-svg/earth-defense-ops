#!/bin/bash
set -e
cd "$(dirname "$0")"

DESCRIPTION="deploy $(date +%Y-%m-%d-%H%M)"

echo "▶ GASにpush中..."
clasp push --force

echo "▶ GitHubにpush中..."
git add -A
git commit -m "$DESCRIPTION" || echo "（変更なし、コミットスキップ）"
git push origin main

echo "✓ 完了！（HEAD deploymentが自動更新されました）"
