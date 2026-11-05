#!/usr/bin/env python3
"""Check remote hook pins using the repository's block-style pre-commit configuration."""
import re
import sys
from pathlib import Path


def violations(text):
    errors = []
    remote = None
    pinned = False
    for number, line in enumerate(text.splitlines(), 1):
        stripped = line.strip()
        if stripped.startswith("#"):
            continue
        repo = re.fullmatch(r"- repo: (\S+)\s*(?:#.*)?", stripped)
        if repo:
            if remote and not pinned:
                errors.append(f"remote hook {remote} has no full commit rev")
            remote = None if repo[1] in ("local", "meta") else repo[1]
            pinned = False
        elif re.match(r"rev\s*:", stripped):
            if not re.fullmatch(r"rev: [0-9a-f]{40}(?:\s+#.*)?", stripped):
                errors.append(f"line {number}: hook rev must be a full lowercase commit SHA")
            else:
                pinned = True
    if remote and not pinned:
        errors.append(f"remote hook {remote} has no full commit rev")
    return errors


def main():
    config = Path(__file__).resolve().parents[1] / ".pre-commit-config.yaml"
    errors = violations(config.read_text())
    if errors:
        print("\n".join(errors), file=sys.stderr)
        return 1
    print("Immutable pre-commit hook revisions verified")
    return 0


if __name__ == "__main__":
    sys.exit(main())
