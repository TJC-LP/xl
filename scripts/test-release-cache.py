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

    def test_alternate_cache_mechanisms_are_rejected(self):
        for line in ("uses: Actions/Cache@v4", "uses: coursier/cache-action@v6",
                     "cache: maven", '"cache": true'):
            self.assertTrue(policy.violations(line), line)

    def test_restore_prefixes_are_rejected(self):
        self.assertTrue(policy.violations("restore-keys: release-coursier-"))

    def test_ordinary_release_steps_pass(self):
        self.assertEqual([], policy.violations("uses: actions/checkout@" + "a" * 40 + "\nrun: ./mill -i __.test"))

    def test_scalar_text_and_trailing_comments_are_ignored(self):
        for text in ('run: echo "no cache: cold build"',
                     'name: Build # no actions/cache@v4',
                     'run: |\n  echo "no cache: cold build"\n  # actions/cache@v4',
                     'name: "actions/cache@v4 is forbidden"'):
            self.assertEqual([],policy.violations(text),text)

    def test_real_cache_after_scalar_still_fails(self):
        text='  -   run: |\n        echo "cache: harmless"\n      uses: actions/cache@v4 # real step'
        self.assertEqual([3],[n for n,_ in policy.violations(text)])
        self.assertTrue(policy.violations('uses: "actions/cache@v4" # cache step'))
        self.assertTrue(policy.violations('- {uses: actions/cache@v4}'))
        self.assertTrue(policy.violations('with: {cache: maven}'))

    def test_multiline_flow_cache_step_still_fails(self):
        text='steps:\n  - {\n    name: Cache, uses: actions/cache@v4\n    }'
        self.assertEqual([3],[n for n,_ in policy.violations(text)])
        text += '\nrun: echo "no cache: harmless"'
        self.assertEqual([3],[n for n,_ in policy.violations(text)])

    def test_comment_is_not_an_active_cache_step(self):
        self.assertEqual([], policy.violations("# actions/cache is deliberately absent"))


if __name__ == "__main__":
    unittest.main()
