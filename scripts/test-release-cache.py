#!/usr/bin/env python3
import importlib.util
import unittest
from pathlib import Path

spec = importlib.util.spec_from_file_location("policy", Path(__file__).with_name("check-release-cache.py"))
policy = importlib.util.module_from_spec(spec)
spec.loader.exec_module(policy)


class ReleaseCacheTest(unittest.TestCase):
    def test_cache_actions_are_rejected_for_tags_and_pins(self):
        for action in ("actions/cache@v4", "actions/cache@" + "a" * 40,
                       "actions/cache/restore@" + "a" * 40, "actions/cache/save@v4"):
            self.assertTrue(policy.violations("uses: " + action))

    def test_restore_prefixes_are_rejected(self):
        self.assertTrue(policy.violations("restore-keys: release-coursier-"))

    def test_ordinary_release_steps_pass(self):
        self.assertEqual([], policy.violations("uses: actions/checkout@" + "a" * 40 + "\nrun: ./mill -i __.test"))

    def test_comment_is_not_an_active_cache_step(self):
        self.assertEqual([], policy.violations("# actions/cache is deliberately absent"))


if __name__ == "__main__":
    unittest.main()
