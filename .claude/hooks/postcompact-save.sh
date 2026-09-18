#!/usr/bin/env bash
# PostCompact hook: when a compaction happens, append the summary Claude Code
# hands us to this project's CLAUDE.md and commit+push it, so the gist of the
# conversation survives even if the session context is later lost (e.g. /clear).
# Committed into the repo itself (project-scoped .claude/settings.json) so it
# travels with every clone/session, regardless of which container runs it.
# Never blocks or fails the session: all errors are swallowed, always exit 0.
set -u

input="$(cat 2>/dev/null || true)"

cwd="$(printf '%s' "$input" | jq -r '.cwd // empty' 2>/dev/null)"
[ -z "$cwd" ] && cwd="$(pwd)"

# Only act inside a git repo that already has a CLAUDE.md - never create one
# from a hook, and never touch a directory that isn't a git checkout.
if [ ! -d "$cwd/.git" ] || [ ! -f "$cwd/CLAUDE.md" ]; then
  exit 0
fi

summary="$(printf '%s' "$input" | jq -r '
  .summary
  // .hookSpecificOutput.summary
  // .compactSummary
  // .custom_instructions
  // empty
' 2>/dev/null)"

ts="$(date '+%Y-%m-%d %H:%M %Z')"

if [ -z "$summary" ] || [ "$summary" = "null" ]; then
  # No summary text was available in the hook payload - still leave a marker
  # so it's visible that a compaction happened and nothing was captured.
  summary="(このコンパクトではPostCompactフックに要約テキストが渡されませんでした。コンパクト前の会話内容はこの保存には反映されていません。)"
fi

{
  echo ""
  echo "## 自動保存（PostCompact: ${ts}）"
  echo ""
  echo "${summary}"
} >> "$cwd/CLAUDE.md" 2>/dev/null

(
  cd "$cwd" || exit 0
  git add CLAUDE.md >/dev/null 2>&1
  if ! git diff --cached --quiet -- CLAUDE.md 2>/dev/null; then
    git commit -m "CLAUDE.mdにPostCompactの自動要約を追記 (${ts})

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>" >/dev/null 2>&1
    branch="$(git rev-parse --abbrev-ref HEAD 2>/dev/null)"
    if [ -n "$branch" ] && [ "$branch" != "HEAD" ]; then
      git push -u origin "$branch" >/dev/null 2>&1
    fi
  fi
) || true

exit 0
