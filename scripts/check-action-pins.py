#!/usr/bin/env python3
"""Require reviewable immutable action references in block-style workflow YAML."""
import re
import sys
from pathlib import Path

REFERENCE = re.compile(r"^\s*(?:-\s*)?uses:\s*['\"]?([A-Za-z0-9_.-]+/[A-Za-z0-9_./-]+)@([0-9a-f]{40})['\"]?\s+#\s+v?\d+\.\d+\.\d+(?:[-+][\w.-]+)?\s*$")
KEY = re.compile(r"(?:^|[\s{,])['\"]?uses['\"]?\s*:")


LOCAL_OR_DOCKER = re.compile(r"^\s*(?:-\s*)?uses:\s*(?:\./[A-Za-z0-9_./-]+|docker://[A-Za-z0-9_./:-]+@sha256:[0-9a-f]{64})\s*(?:#.*)?$")
MAPPING_KEY = re.compile(r"^\s*(?:-\s*)?['\"]?uses['\"]?\s*:")
BLOCK_SCALAR = re.compile(r"^\s*(?:-\s*)?[\w-]+:\s*[|>][+\-1-9]*\s*(?:#.*)?$")


def violations(text):
    errors = []
    scalar_indent = None
    for number, line in enumerate(text.splitlines(), 1):
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        indent = len(line) - len(line.lstrip())
        if scalar_indent is not None:
            if indent > scalar_indent:
                continue
            scalar_indent = None
        if BLOCK_SCALAR.fullmatch(line) and not MAPPING_KEY.match(line):
            # A list item's mapping keys are indented beyond its dash.
            scalar_indent = indent + (2 if line.lstrip().startswith("- ") else 0)
            continue
        # Flow mappings remain unsupported and fail closed. Scalar values such
        # as run/name strings are not action references.
        flow_mapping = re.match(r"^\s*(?:[-,]\s*)?(?:\[\s*)?\{", line) or re.match(r"^\s*(?:[\"][^\"]+[\"]|['][^']+[']|[\w-]+):\s*(?:\[\s*)?\{", line)
        if flow_mapping:
            errors.append((number, line.strip()))
            continue
        if MAPPING_KEY.match(line) or (line.lstrip().startswith(",") and KEY.search(line)):
            if not REFERENCE.fullmatch(line) and not LOCAL_OR_DOCKER.fullmatch(line):
                errors.append((number, line.strip()))
    return errors


def main():
    root = Path(__file__).resolve().parents[1]
    files = sorted((root / ".github/workflows").glob("*.yml")) + sorted((root / ".github/workflows").glob("*.yaml"))
    errors = [f"{path.relative_to(root)}:{number}: {line}"
              for path in files for number, line in violations(path.read_text())]
    if errors:
        print("Actions must use a full lowercase commit SHA and a version comment; use block-style workflow mappings:", file=sys.stderr)
        print("\n".join(errors), file=sys.stderr)
        return 1
    print(f"Immutable action references verified in {len(files)} workflows")
    return 0


if __name__ == "__main__":
    sys.exit(main())
