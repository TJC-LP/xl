#!/usr/bin/env python3
import importlib.util
import unittest
from pathlib import Path

spec = importlib.util.spec_from_file_location("pins", Path(__file__).with_name("check-hook-pins.py"))
pins = importlib.util.module_from_spec(spec)
spec.loader.exec_module(pins)


class HookPinsTest(unittest.TestCase):
    def test_remote_commit_and_local_hooks(self):
        self.assertEqual([], pins.violations("- repo: local\n- repo: https://example.com/hooks\n  rev: " + "a" * 40 + " # v1.0.0"))

    def test_tag_and_short_hash_rejected(self):
        for rev in ("v4.5.0", "main", "a" * 39):
            self.assertTrue(pins.violations("- repo: https://example.com/hooks\n  rev: " + rev))

    def test_missing_remote_rev_rejected(self):
        self.assertTrue(pins.violations("- repo: https://example.com/hooks\n  hooks: []"))

    def test_unsupported_layouts_fail_closed(self):
        for text in (
            "- rev: v4.5.0\n  repo: https://example.com/hooks",
            '- {repo: https://example.com/hooks, rev: v4.5.0}',
            '- repo: https://example.com/hooks\n  "rev": v4.5.0',
        ):
            self.assertTrue(pins.violations(text))

    def test_second_repo_cannot_inherit_first_pin(self):
        self.assertTrue(pins.violations("- repo: https://example.com/one\n  rev: " + "a" * 40 + "\n- repo: https://example.com/two"))


if __name__ == "__main__":
    unittest.main()
