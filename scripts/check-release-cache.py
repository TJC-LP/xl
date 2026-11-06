#!/usr/bin/env python3
"""Reject known shared-cache actions and cache inputs in release jobs (GH-836).
This lexical guard does not audit caching hidden inside arbitrary third-party code.
"""
import re
import sys
from pathlib import Path


def violations(text):
    return [(number, line.strip()) for number, line in enumerate(text.splitlines(), 1)
            if not line.lstrip().startswith("#") and
            (re.search(r"(?:actions/cache|coursier/cache-action)(?:[/@\s'\"]|$)", line, re.IGNORECASE) or
             re.search(r"(?:^|[\s{,])['\"]?(?:restore-keys|cache)['\"]?\s*:", line, re.IGNORECASE))]


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
