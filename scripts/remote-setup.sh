#!/usr/bin/env bash
# Provision what a Claude Code cloud sandbox lacks for this repo, then export it to the session.
#
# Wired as a SessionStart hook in .claude/settings.json. It is a no-op unless CLAUDE_CODE_REMOTE=true
# (Claude Code on the web, routines, `claude --cloud`, claude-code-action) or --force is passed, so
# local sessions are untouched. Safe to re-run: every step is idempotent and takes seconds once the
# environment cache already holds the toolchain.
#
# Why it exists: the Anthropic-hosted sandbox is Ubuntu 24.04 with OpenJDK 21 and no scala-cli.
# The build compiles with `--release 25` and pins Temurin 25 in .mill-jvm-version. Mill downloads
# that JDK by itself on first use, so `./mill` works with zero setup; this script additionally puts
# the same JDK on PATH (for scala-cli, `java -jar`, `make install-jar`) and installs scala-cli for
# the examples harness and the skill-snippet gate.
#
# Steps:
#   1. Fetch the coursier launcher (`cs`) if missing.
#   2. Provision the JDK named in .mill-jvm-version through coursier (shared cache with Mill).
#   3. Install scala-cli.
#   4. Append JAVA_HOME/PATH to $CLAUDE_ENV_FILE, which Claude Code sources before every Bash call.
#   5. Print a one-paragraph status; SessionStart stdout lands in Claude's context.
#
# Usage: bash scripts/remote-setup.sh [--force]
set -uo pipefail

FORCE=false
[[ "${1:-}" == "--force" ]] && FORCE=true
if [[ "${CLAUDE_CODE_REMOTE:-}" != "true" && "$FORCE" != "true" ]]; then
  exit 0
fi

ROOT="${CLAUDE_PROJECT_DIR:-$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)}"
JVM_SPEC="$(tr -d '[:space:]' < "$ROOT/.mill-jvm-version" 2>/dev/null || true)"
JVM_SPEC="${JVM_SPEC:-temurin:25}"
BIN="${XL_REMOTE_BIN:-$HOME/.local/bin}"
mkdir -p "$BIN"

NOTES=""
note() { NOTES="${NOTES:+$NOTES; }$*"; }
warn() { echo "[remote-setup] WARN: $*" >&2; note "WARN: $*"; }

# ---------- 1. coursier launcher ----------
CS=""
if command -v cs >/dev/null 2>&1; then
  CS="$(command -v cs)"
elif [[ -x "$BIN/cs" ]]; then
  CS="$BIN/cs"
elif [[ "$(uname -s)" == "Linux" ]]; then
  case "$(uname -m)" in
    x86_64 | amd64) ARCH=x86_64 ;;
    aarch64 | arm64) ARCH=aarch64 ;;
    *) ARCH="" ;;
  esac
  URL="https://github.com/coursier/launchers/raw/master/cs-${ARCH}-pc-linux.gz"
  if [[ -n "$ARCH" ]] && curl -fsSL "$URL" | gzip -d > "$BIN/cs.tmp" 2>/dev/null && [[ -s "$BIN/cs.tmp" ]]; then
    chmod +x "$BIN/cs.tmp" && mv "$BIN/cs.tmp" "$BIN/cs" && CS="$BIN/cs"
    note "coursier: installed to $BIN/cs"
  else
    rm -f "$BIN/cs.tmp"
    warn "could not fetch the coursier launcher ($URL); ./mill still self-provisions its JDK, but scala-cli and java-on-PATH are unavailable"
  fi
else
  warn "no coursier launcher found and this is not Linux; install it manually (https://get-coursier.io) to use --force here"
fi

# ---------- 2. JDK from .mill-jvm-version ----------
JAVA_HOME_NEW=""
if [[ -n "$CS" ]]; then
  if JAVA_HOME_NEW="$("$CS" java-home --jvm "$JVM_SPEC" 2>/dev/null)" && [[ -x "$JAVA_HOME_NEW/bin/java" ]]; then
    note "JDK $JVM_SPEC at $JAVA_HOME_NEW"
  else
    JAVA_HOME_NEW=""
    warn "coursier could not provision $JVM_SPEC; ./mill still downloads it itself, but java on PATH stays the sandbox default"
  fi

  # ---------- 3. scala-cli ----------
  if command -v scala-cli >/dev/null 2>&1 || [[ -x "$BIN/scala-cli" ]]; then
    note "scala-cli present"
  elif "$CS" install --install-dir "$BIN" scala-cli >/dev/null 2>&1 && [[ -x "$BIN/scala-cli" ]]; then
    note "scala-cli installed to $BIN"
  else
    warn "scala-cli install failed; scripts/test-examples.sh and scripts/verify-skill-snippets.sh need it"
  fi
fi

# ---------- 4. Export to the session ----------
PATH_NEW="$BIN:$PATH"
[[ -n "$JAVA_HOME_NEW" ]] && PATH_NEW="$JAVA_HOME_NEW/bin:$PATH_NEW"
if [[ -n "${CLAUDE_ENV_FILE:-}" ]]; then
  {
    [[ -n "$JAVA_HOME_NEW" ]] && printf 'export JAVA_HOME=%q\n' "$JAVA_HOME_NEW"
    printf 'export PATH=%q\n' "$PATH_NEW"
  } >> "$CLAUDE_ENV_FILE"
  note "JAVA_HOME/PATH exported for this session"
else
  note "CLAUDE_ENV_FILE unset, so nothing was exported (prefix PATH with $BIN and the JDK bin yourself if needed)"
fi

# ---------- 5. Context for Claude ----------
echo "xl remote setup (scripts/remote-setup.sh): ${NOTES}. Mill pins its own JDK via .mill-jvm-version. In a fresh sandbox the first ./mill build downloads dependencies: give it a 600000 ms Bash timeout. Not available here: gtr, GraalVM (use 'make install-jar', not 'make install'). Details: docs/reference/remote-sessions.md."
exit 0
