#!/usr/bin/env python3
"""Reject known shared-cache actions and cache inputs in release jobs (GH-836).
This lexical guard does not audit caching hidden inside arbitrary third-party code.
"""
import re
import sys
from pathlib import Path


def without_comment(line):
    """Strip YAML comments, keeping # characters inside quoted values."""
    quote = None
    escaped = False
    for index, char in enumerate(line):
        if escaped:
            escaped = False
        elif quote == '"' and char == "\\":
            escaped = True
        elif char == quote:
            quote = None
        elif quote is None and char in "\"'" and (index == 0 or line[index-1] in " :,[{"):
            quote = char
        elif quote is None and char == '#' and (index == 0 or line[index-1].isspace()):
            return line[:index]
    return line


def violations(text):
    errors = []
    scalar_indent = None
    flow_depth = 0
    for number, raw in enumerate(text.splitlines(), 1):
        line = without_comment(raw)
        if not line.strip():
            continue
        indent = len(line) - len(line.lstrip())
        if scalar_indent is not None:
            if indent > scalar_indent:
                continue
            scalar_indent = None
        mapping = re.match(r"^\s*(?:-\s*)?[\"']?([\w-]+)[\"']?\s*:\s*(.*)$", line)
        in_flow = flow_depth > 0 or re.match(r"^\s*(?:-\s*)?[{\[]", line) or (
            mapping and mapping.group(2).lstrip().startswith(('{', '[')))
        if in_flow:
            # Keep scanning continuation mappings (e.g. name: x, uses: cache)
            # until the surrounding flow collection closes. Quoted braces do
            # not change collection depth.
            structure = re.sub(r"\"(?:[^\"\\]|\\.)*\"|'(?:[^']|'')*'", '', line)
            flow_depth = max(0, flow_depth + sum(structure.count(c) for c in '{[')
                             - sum(structure.count(c) for c in '}]'))
        if mapping and not in_flow:
            key, value = mapping.group(1).lower(), mapping.group(2)
            if key not in ('uses', 'cache', 'restore-keys') and re.fullmatch(r"[|>][+\-1-9]*\s*", value):
                scalar_indent = re.match(r"^\s*(?:-\s+)?", line).end()
                continue
            if key in ('cache', 'restore-keys'):
                errors.append((number, raw.strip()))
                continue
            if key == 'uses':
                if re.search(r"(?:actions/cache|coursier/cache-action)(?:[/@\s'\"]|$)", value, re.IGNORECASE):
                    errors.append((number, raw.strip()))
                continue
            # Ordinary scalar text (run/name/etc.) cannot declare a cache step.
            if not value.lstrip().startswith(('{', '[')):
                continue
        # Retain the existing conservative check for flow mappings and their
        # continuation lines, rather than treating them as shell text.
        if (re.search(r"(?:actions/cache|coursier/cache-action)(?:[/@\s'\"]|$)", line, re.IGNORECASE) or
                re.search(r"(?:^|[\s{,])[\"']?(?:restore-keys|cache)[\"']?\s*:", line, re.IGNORECASE)):
            errors.append((number, raw.strip()))
    return errors


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
