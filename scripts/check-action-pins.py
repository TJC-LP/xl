#!/usr/bin/env python3
"""Enforce immutable uses references in a deliberately restricted block-YAML grammar.

Recognize supported lines positively: unknown structure is an error, not an
ignored line. This is a repository format policy, not a general YAML parser.
"""
import re
import sys
from pathlib import Path

REFERENCE = re.compile(r"^\s*(?:-\s*)?uses:\s*(?P<quote>['\"]?)([A-Za-z0-9_.-]+/[A-Za-z0-9_./-]+)@([0-9a-f]{40})(?P=quote)\s+#\s+v?\d+\.\d+\.\d+(?:[-+][\w.-]+)?\s*$", re.ASCII)
LOCAL_OR_DOCKER = re.compile(r"^\s*(?:-\s*)?uses:\s*(?:\./[A-Za-z0-9_./-]+|docker://[A-Za-z0-9_./:-]+@sha256:[0-9a-f]{64})\s*(?:#.*)?$", re.ASCII)
MAPPING = re.compile(r'''(?P<key>[A-Za-z0-9_.-]+|"[^"\\]*"|'[^'\\]*') *:(?:[ \t]+|$)(?P<value>.*)$''')
QUOTED = re.compile(r'''(?:"(?:[^"\\]|\\.)*"|'(?:[^']|'')*')(?:[ \t]+\#.*)?$''')
BLOCK_HEADER = re.compile(r'[|>](?:[1-9][+-]?|[+-][1-9]?|)$')
PLAIN_LIST = re.compile(r'\[[\w./*+ -]*(?:,[\w./*+ -]*)*\]$')
OPTION_ENTRY = re.compile(r'([A-Za-z0-9_.-]+) *: *[A-Za-z0-9_./+-][A-Za-z0-9_./*+-]*$')


def without_comment(value):
    # Called only for unquoted values. Quotes inside a YAML plain scalar are
    # literal characters, not quote delimiters (e.g. run: echo "$REF#suffix").
    return re.split(r'(?:^|[ \t])#', value, maxsplit=1)[0].rstrip()


def scalar_or_options(value, key, sequence_item):
    if value.startswith(('"', "'")):
        return QUOTED.fullmatch(value) is not None
    value = without_comment(value)
    if not value:
        return True
    if value.startswith('['):
        return (PLAIN_LIST.fullmatch(value) is not None
                and not any(item.strip().startswith('*') for item in value[1:-1].split(',')))
    if value.startswith('{'):
        if re.fullmatch(r'\{[ \t]*\}', value):
            return True
        if key not in {'env', 'with', 'permissions'} or not value.endswith('}'):
            return False
        entries = [OPTION_ENTRY.fullmatch(entry.strip()) for entry in value[1:-1].split(',')]
        return all(entry is not None and entry[1] != 'uses' for entry in entries)
    # A scalar cannot begin with a collection/property/directive indicator.
    if value[0] in '&*!?{}[],|>@`%' or re.match(r'[-:](?:[ \t]|$)', value):
        return False
    # A sequence item's unrecognized key must not fall back to scalar text.
    if sequence_item and re.search(r':(?:[ \t]|$)', value):
        return False
    return True


def violations(text):
    errors = []
    scalar_parent = None
    scalar_content = None
    # YAML permits a stream-leading BOM; it must not mask the first structure.
    if text.startswith('\ufeff'):
        text = text[1:]
    for number, line in enumerate(text.splitlines(), 1):
        indent = len(line) - len(line.lstrip(' '))
        if not line.strip():
            continue
        if scalar_parent is not None:
            if scalar_content is None and indent > scalar_parent:
                scalar_content = indent
            if scalar_content is not None and indent >= scalar_content:
                continue
            scalar_parent = scalar_content = None
        content = line[indent:]
        if content.startswith('#'):
            continue
        if '\ufeff' in line or any(ord(c) < 32 and c != '\t' for c in line):
            errors.append((number, line.strip()))
            continue
        sequence = re.match(r'-(?: +|$)', content)
        if sequence:
            indent += sequence.end()
            content = content[sequence.end():]
            if not content or content.startswith('#'):
                continue
        # Tabs after a dash are not supported indentation/separation. Refuse
        # before scalar fallback so they cannot mask aliases/tags/collections.
        if content.startswith('\t'):
            errors.append((number, line.strip()))
            continue
        mapping = MAPPING.fullmatch(content)
        if mapping is None:
            if not sequence or not scalar_or_options(content, None, True):
                errors.append((number, line.strip()))
            continue
        key = mapping['key']
        if key.startswith(('"', "'")):
            key = key[1:-1]
        value = mapping['value']
        # This precedes every scalar/empty-map/list exemption. A uses value must
        # be an immutable reference even when its value has the wrong YAML type.
        if key == 'uses':
            if not REFERENCE.fullmatch(line) and not LOCAL_OR_DOCKER.fullmatch(line):
                errors.append((number, line.strip()))
            continue
        header = without_comment(value)
        if BLOCK_HEADER.fullmatch(header):
            scalar_parent = indent
            indicator = re.search(r'[1-9]', header)
            scalar_content = indent + int(indicator[0]) if indicator else None
        elif not scalar_or_options(value, key, False):
            errors.append((number, line.strip()))
    return errors


def main():
    root = Path(__file__).resolve().parents[1]
    files = sorted((root / '.github/workflows').glob('*.yml')) + sorted((root / '.github/workflows').glob('*.yaml'))
    errors = [f'{path.relative_to(root)}:{number}: {line}'
              for path in files for number, line in violations(path.read_text(encoding='utf-8'))]
    if errors:
        print('Actions require a full lowercase SHA and version comment; unsupported YAML must use the documented block-style forms:', file=sys.stderr)
        print('\n'.join(errors), file=sys.stderr)
        return 1
    print(f'Immutable action references verified in {len(files)} workflows')
    return 0


if __name__ == '__main__':
    sys.exit(main())
