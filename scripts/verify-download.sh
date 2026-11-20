#!/usr/bin/env sh
# Sourced by repository launchers; no downloads or side effects on source.
xl_expected_sha256() {
  awk -v key="$2" '$2 == key { print $1; found++ } END { if (found != 1) exit 1 }' "$1"
}
xl_verify_sha256() {
  case "$1" in *[!0-9a-f]*|'') echo "Invalid expected SHA-256" >&2; return 1 ;; esac
  [ "${#1}" -eq 64 ] || { echo "Invalid expected SHA-256 length" >&2; return 1; }
  if command -v sha256sum >/dev/null 2>&1; then
    xl_actual=$(sha256sum "$2") || return 1
  elif command -v shasum >/dev/null 2>&1; then
    xl_actual=$(shasum -a 256 "$2") || return 1
  else
    echo "A SHA-256 tool (sha256sum or shasum) is required" >&2; return 1
  fi
  [ "${xl_actual%% *}" = "$1" ] || { echo "SHA-256 mismatch: $2" >&2; return 1; }
}
