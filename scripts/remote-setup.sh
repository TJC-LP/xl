#!/usr/bin/env bash
# Provision what a Claude Code cloud sandbox lacks for this repo, then export it to the session.
#
# Wired as a SessionStart hook in .claude/settings.json. It is a no-op unless CLAUDE_CODE_REMOTE=true
# (Claude Code on the web, routines, `claude --cloud`, claude-code-action) or --force is passed, so
# local sessions are untouched. Safe to re-run: every step is idempotent — the toolchain is reused
# from the environment cache, PATH is built without repeating a prefix it already has, and a line
# is appended to $CLAUDE_ENV_FILE only when the file does not already carry it — so a second
# SessionStart takes seconds and adds nothing.
#
# Why it exists: the Anthropic-hosted sandbox is Ubuntu 24.04 with OpenJDK 21 and no scala-cli.
# The build compiles with `--release 25` and pins Temurin 25 in .mill-jvm-version. Mill downloads
# that JDK by itself on first use, so `./mill` works with zero setup; this script additionally puts
# the same JDK on PATH (for scala-cli, `java -jar`, `make install-jar`) and installs scala-cli for
# the examples harness and the skill-snippet gate.
#
# TLS-intercepting proxies: some sandboxes route HTTPS through a proxy that re-terminates TLS and
# inject its CA into the JVM through JAVA_TOOL_OPTIONS (-Djavax.net.ssl.trustStore=...). GraalVM
# native launchers — Mill's default `<version>-native-*` binary, coursier's `cs`, scala-cli's native
# image — do not read JAVA_TOOL_OPTIONS and die with PKIX errors on their first download. When that
# truststore is present (or XL_REMOTE_JVM_LAUNCHERS=true), every launcher here runs on the JVM
# instead: `./mill` via MILL_VERSION=<.mill-version>-jvm, coursier via its JAR, scala-cli via a
# coursier bootstrap launcher. `curl` is unaffected either way (it honours SSL_CERT_FILE).
#
# Steps:
#   0. Detect an injected JVM truststore; if present, switch ./mill to the JVM launcher.
#   1. Obtain coursier (`cs` native launcher, or the JAR behind a proxy) if missing.
#   2. Provision the JDK named in .mill-jvm-version through coursier (shared cache with Mill).
#   3. Install scala-cli (native launcher, or a JVM bootstrap behind a proxy).
#   4. Append MILL_VERSION/JAVA_HOME/PATH to $CLAUDE_ENV_FILE, which Claude Code sources before
#      every Bash call (each line once; a re-run appends nothing already there).
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
MILL_VER="$(tr -d '[:space:]' < "$ROOT/.mill-version" 2>/dev/null || true)"
BIN="${XL_REMOTE_BIN:-$HOME/.local/bin}"
mkdir -p "$BIN"
# Pinned JVM fallbacks; bump deliberately (both resolve from Maven Central / GitHub releases).
COURSIER_VERSION="${XL_REMOTE_COURSIER_VERSION:-2.1.24}"
SCALA_CLI_VERSION="${XL_REMOTE_SCALA_CLI_VERSION:-1.9.1}"

# Coursier downloads in our private cache must match committed bytes, including cache hits.
# Caller-provided PATH tools and Coursier-managed Scala CLI/JDK installs are separate trust inputs.
. "$ROOT/scripts/verify-download.sh"
NOTES=""
note() { NOTES="${NOTES:+$NOTES; }$*"; }
warn() { echo "[remote-setup] WARN: $*" >&2; note "WARN: $*"; }

# ---------- 0. TLS-intercepting proxy → JVM launchers ----------
JVM_LAUNCHERS=false
if [[ "${JAVA_TOOL_OPTIONS:-}" == *javax.net.ssl.trustStore* || "${XL_REMOTE_JVM_LAUNCHERS:-}" == "true" ]]; then
  JVM_LAUNCHERS=true
fi
MILL_VERSION_EXPORT=""
if [[ "$JVM_LAUNCHERS" == "true" ]]; then
  if ! command -v java >/dev/null 2>&1; then
    warn "a JVM truststore is injected but no java is on PATH; native launchers will fail TLS and nothing here can run on the JVM"
    JVM_LAUNCHERS=false
  elif [[ -n "$MILL_VER" && "$MILL_VER" != *-jvm && "$MILL_VER" != *-native ]]; then
    MILL_VERSION_EXPORT="${MILL_VER}-jvm"
    note "JVM truststore detected: ./mill uses the JVM launcher (MILL_VERSION=$MILL_VERSION_EXPORT)"
  fi
fi

# ---------- 1. coursier ----------
# cs_run abstracts over the native launcher and the JAR so steps 2-3 read the same either way.
CS=""
CS_JAR=""
CS_CACHE="$BIN/.xl-remote"
mkdir -p "$CS_CACHE"
cs_run() {
  if [[ -n "$CS_JAR" ]]; then java -jar "$CS_JAR" "$@"; else "$CS" "$@"; fi
}
# Only execute downloaded/cache-managed tools after checking committed bytes.
# Failure leaves CS empty and lets the remaining session exports complete.
verified_fetch() {
  local key="$1" url="$2" target="$3" expected
  if ! expected=$(xl_expected_sha256 "$ROOT/scripts/launcher-checksums.txt" "$key"); then
    warn "no committed checksum for $key; skipping Coursier provisioning"
    return 1
  fi
  if [[ -s "$target" ]]; then
    if xl_verify_sha256 "$expected" "$target"; then return 0; fi
    warn "discarding mismatched cached $key"
    rm -f "$target"
  fi
  if curl -fsSL "$url" -o "$target.tmp" &&
      xl_verify_sha256 "$expected" "$target.tmp"; then
    mv "$target.tmp" "$target"
  else
    rm -f "$target.tmp"
    warn "could not obtain verified $key; skipping Coursier provisioning"
    return 1
  fi
}
SYSTEM=$(uname -s)
ARCH=$(uname -m)
PATH_CS=$(command -v cs || true)
# Caller tools elsewhere on PATH are caller-trusted. On Linux, legacy $BIN/cs
# may be either user-owned or an old unverified download: preserve it, but use
# the private verified cache instead. Never delete or replace that ambiguous file.
# Only .xl-remote is setup-owned; no unverified executable from it is trusted.
if [[ "$JVM_LAUNCHERS" != "true" && -n "$PATH_CS" &&
      "$PATH_CS" != "$CS_CACHE/"* &&
      ( "$SYSTEM" != "Linux" || "$PATH_CS" != "$BIN/cs" ) ]]; then
  CS="$PATH_CS"
elif [[ "$JVM_LAUNCHERS" == "true" ||
        ( "$SYSTEM" == "Linux" && ( "$ARCH" == "aarch64" || "$ARCH" == "arm64" ) ) ]]; then
  if ! java -version >/dev/null 2>&1; then
    warn "Coursier JVM launcher needs a working Java; skipping provisioning"
  else
    CS_JAR_PATH="$CS_CACHE/coursier.jar"
    URL="https://github.com/coursier/coursier/releases/download/v${COURSIER_VERSION}/coursier"
    if verified_fetch "coursier-$COURSIER_VERSION" "$URL" "$CS_JAR_PATH"; then
      CS_JAR="$CS_JAR_PATH"
    fi
  fi
elif [[ "$SYSTEM" == "Linux" && ( "$ARCH" == "x86_64" || "$ARCH" == "amd64" ) ]]; then
  if ! CS_NATIVE_SHA=$(xl_expected_sha256 "$ROOT/scripts/launcher-checksums.txt" "cs-$COURSIER_VERSION-x86_64-pc-linux"); then
    warn "no committed native Coursier checksum for $COURSIER_VERSION; skipping provisioning"
  elif [[ -x "$CS_CACHE/cs" ]] && xl_verify_sha256 "$CS_NATIVE_SHA" "$CS_CACHE/cs"; then
    CS="$CS_CACHE/cs"
  else
    [[ ! -e "$CS_CACHE/cs" ]] || { warn "discarding mismatched cached cs"; rm -f "$CS_CACHE/cs"; }
    URL="https://github.com/coursier/coursier/releases/download/v${COURSIER_VERSION}/cs-x86_64-pc-linux.gz"
    if verified_fetch "cs-$COURSIER_VERSION-x86_64-pc-linux.gz" "$URL" "$CS_CACHE/cs.gz" &&
        gzip -dc "$CS_CACHE/cs.gz" > "$CS_CACHE/cs.tmp" &&
        xl_verify_sha256 "$CS_NATIVE_SHA" "$CS_CACHE/cs.tmp"; then
      chmod +x "$CS_CACHE/cs.tmp" && mv "$CS_CACHE/cs.tmp" "$CS_CACHE/cs" && CS="$CS_CACHE/cs"
    else
      rm -f "$CS_CACHE/cs.tmp"
      warn "could not provision verified native Coursier"
    fi
    rm -f "$CS_CACHE/cs.gz"
  fi
else
  warn "no trusted Coursier on PATH for $SYSTEM/$ARCH; skipping provisioning"
fi

# ---------- 2. JDK from .mill-jvm-version ----------
# Coursier's default JVM index lives behind a github.com/.../raw redirect that egress proxies tend
# to refuse (403); the canonical raw.githubusercontent.com location is what the redirect resolves to.
JVM_INDEX_OPT=()
[[ "$JVM_LAUNCHERS" == "true" ]] && JVM_INDEX_OPT=(--jvm-index "https://raw.githubusercontent.com/coursier/jvm-index/master/index.json")
JAVA_HOME_NEW=""
if [[ -n "$CS" || -n "$CS_JAR" ]]; then
  if JAVA_HOME_NEW="$(cs_run java-home --jvm "$JVM_SPEC" ${JVM_INDEX_OPT[@]+"${JVM_INDEX_OPT[@]}"} 2>/dev/null)" && [[ -x "$JAVA_HOME_NEW/bin/java" ]]; then
    note "JDK $JVM_SPEC at $JAVA_HOME_NEW"
  else
    JAVA_HOME_NEW=""
    warn "coursier could not provision $JVM_SPEC; ./mill still downloads it itself, but java on PATH stays the sandbox default"
  fi

  # ---------- 3. scala-cli ----------
  if command -v scala-cli >/dev/null 2>&1 || [[ -x "$BIN/scala-cli" ]]; then
    note "scala-cli present"
  elif [[ -n "$CS_JAR" ]]; then
    # A coursier bootstrap is a JVM launcher: it fetches scala-cli's jars through the JVM (and so
    # through the injected truststore) instead of downloading the native image.
    if cs_run bootstrap "org.virtuslab.scala-cli:cli_3:${SCALA_CLI_VERSION}" -M scala.cli.ScalaCli \
        -o "$BIN/scala-cli" -f >/dev/null 2>&1 && [[ -x "$BIN/scala-cli" ]]; then
      note "scala-cli $SCALA_CLI_VERSION installed to $BIN as a JVM launcher"
    else
      warn "scala-cli JVM bootstrap failed; scripts/test-examples.sh and scripts/verify-skill-snippets.sh need it"
    fi
  elif cs_run install --install-dir "$BIN" "scala-cli:$SCALA_CLI_VERSION" >/dev/null 2>&1 && [[ -x "$BIN/scala-cli" ]]; then
    note "scala-cli installed to $BIN"
  else
    warn "scala-cli install failed; scripts/test-examples.sh and scripts/verify-skill-snippets.sh need it"
  fi
fi

# ---------- 4. Export to the session ----------
# The hook runs on every SessionStart, so $CLAUDE_ENV_FILE may already hold these lines and PATH
# may already start with these directories: prepend only what is missing, append only new lines.
prepend_path() { # prepend_path DIR PATH → DIR:PATH, or PATH unchanged when DIR is already on it
  case ":$2:" in *":$1:"*) printf '%s' "$2" ;; *) printf '%s:%s' "$1" "$2" ;; esac
}
PATH_NEW="$(prepend_path "$BIN" "$PATH")"
[[ -n "$JAVA_HOME_NEW" ]] && PATH_NEW="$(prepend_path "$JAVA_HOME_NEW/bin" "$PATH_NEW")"
if [[ -n "${CLAUDE_ENV_FILE:-}" ]]; then
  append_once() { # append_once LINE: add LINE to $CLAUDE_ENV_FILE unless the file already has it
    [[ -f "$CLAUDE_ENV_FILE" ]] && grep -qxF -- "$1" "$CLAUDE_ENV_FILE" || printf '%s\n' "$1" >> "$CLAUDE_ENV_FILE"
  }
  [[ -n "$MILL_VERSION_EXPORT" ]] && append_once "$(printf 'export MILL_VERSION=%q' "$MILL_VERSION_EXPORT")"
  [[ -n "$JAVA_HOME_NEW" ]] && append_once "$(printf 'export JAVA_HOME=%q' "$JAVA_HOME_NEW")"
  append_once "$(printf 'export PATH=%q' "$PATH_NEW")"
  note "${MILL_VERSION_EXPORT:+MILL_VERSION/}JAVA_HOME/PATH exported for this session"
else
  note "CLAUDE_ENV_FILE unset, so nothing was exported (prefix PATH with $BIN and the JDK bin yourself${MILL_VERSION_EXPORT:+, and export MILL_VERSION=$MILL_VERSION_EXPORT} if needed)"
fi

# ---------- 5. Context for Claude ----------
echo "xl remote setup (scripts/remote-setup.sh): ${NOTES}. Mill pins its own JDK via .mill-jvm-version. In a fresh sandbox the first ./mill build downloads dependencies: give it a 600000 ms Bash timeout. Not available here: gtr, GraalVM (use 'make install-jar', not 'make install'). Details: docs/reference/remote-sessions.md."
exit 0
