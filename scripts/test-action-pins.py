#!/usr/bin/env python3
import importlib.util
import unittest
from pathlib import Path

spec = importlib.util.spec_from_file_location("pins", Path(__file__).with_name("check-action-pins.py"))
pins = importlib.util.module_from_spec(spec)
spec.loader.exec_module(pins)


class ActionPinsTest(unittest.TestCase):
    def test_current_commit_is_accepted(self):
        self.assertEqual([], pins.violations("  - uses: actions/checkout@" + "a" * 40 + " # v4.4.0"))

    def test_mutable_short_or_missing_references_are_rejected(self):
        for ref in ("v4", "main", "a" * 39, "a" * 41, "${{ inputs.ref }}"):
            with self.subTest(ref=ref):
                self.assertTrue(pins.violations("  uses: owner/action@" + ref))

    def test_version_comment_is_required(self):
        self.assertTrue(pins.violations("uses: owner/action@" + "a" * 40))

    def test_flow_and_quoted_keys_cannot_hide_tags(self):
        for line in ('- {uses: owner/action@v1}', '  "uses": owner/action@v1', "  uses: 'owner/action@v1'"):
            self.assertTrue(pins.violations(line))

    def test_comments_are_ignored(self):
        self.assertEqual([], pins.violations(" # uses: owner/action@v1"))


if __name__ == "__main__":
    unittest.main()
