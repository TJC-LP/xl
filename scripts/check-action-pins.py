#!/usr/bin/env python3
"""Require reviewable immutable action references in block-style workflow YAML."""
import re
import sys
from pathlib import Path

REFERENCE = re.compile(r"^\s*(?:-\s*)?uses:\s*['\"]?([A-Za-z0-9_.-]+/[A-Za-z0-9_./-]+)@([0-9a-f]{40})['\"]?\s+#\s+v?\d+\.\d+\.\d+(?:[-+][\w.-]+)?\s*$")
KEY = re.compile(r"(?:^|[\s{,])['\"]?uses['\"]?\s*:")


def violations(text):
    return [(number, line.strip()) for number, line in enumerate(text.splitlines(), 1)
            if not line.lstrip().startswith("#") and KEY.search(line) and not REFERENCE.fullmatch(line)]


def main():
    root = Path(__file__).resolve().parents[1]
    files = sorted((root / ".github/workflows").glob("*.yml")) + sorted((root / ".github/workflows").glob("*.yaml"))
    errors = [f"{path.relative_to(root)}:{number}: {line}"
              for path in files for number, line in violations(path.read_text())]
    if errors:
        print("Actions must use a full lowercase commit SHA and a version comment; use block-style uses:", file=sys.stderr)
        print("\n".join(errors), file=sys.stderr)
        return 1
    print(f"Immutable action references verified in {len(files)} workflows")
    return 0


if __name__ == "__main__":
    sys.exit(main())
