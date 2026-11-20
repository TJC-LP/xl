#!/usr/bin/env python3
"""Keep the documented native installer aligned with the released plugin version."""
import json
import re
import sys
from pathlib import Path


def violations(root):
    plugin = json.loads((root / "plugin/.claude-plugin/plugin.json").read_text())["version"]
    skill = (root / "plugin/skills/xl-cli/SKILL.md").read_text()
    install = (root / "plugin/skills/xl-cli/reference/INSTALL.md").read_text()
    versions = re.findall(r"^VERSION=([0-9]+\.[0-9]+\.[0-9]+)$", install, re.MULTILINE)
    floors = re.findall(r"Requires xl >= ([0-9]+\.[0-9]+\.[0-9]+)", skill)
    errors = []
    if versions != [plugin]:
        errors.append("INSTALL.md VERSION must equal plugin.json; update it during release preparation")
    if "<!-- unreleased-contract -->" not in skill and floors != [plugin]:
        errors.append("Released skill Requires floor must match the documented installer version")
    return errors


if __name__ == "__main__":
    errors = violations(Path(__file__).resolve().parents[1])
    for error in errors: print(error, file=sys.stderr)
    sys.exit(bool(errors))
