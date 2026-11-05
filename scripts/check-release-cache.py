#!/usr/bin/env python3
"""Release jobs must not restore or save shared Actions caches (GH-836)."""
import re
import sys
from pathlib import Path


def violations(text):
    return [(number, line.strip()) for number, line in enumerate(text.splitlines(), 1)
            if not line.lstrip().startswith("#") and
            (re.search(r"actions/cache(?:[/@\s'\"]|$)", line) or
             re.search(r"\brestore-keys\s*:", line))]


def main():
    path = Path(__file__).resolve().parents[1] / ".github/workflows/release.yml"
    errors = violations(path.read_text())
    if errors:
        for number, line in errors:
            print(f"release.yml:{number}: shared cache forbidden: {line}", file=sys.stderr)
        return 1
    print("Release workflow contains no Actions cache steps or restore keys")
    return 0


if __name__ == "__main__":
    sys.exit(main())
