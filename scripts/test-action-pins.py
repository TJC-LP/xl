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

    def test_local_and_digest_references(self):
        for value in ("./.github/actions/setup", "./.github/workflows/build.yml",
                      "docker://alpine@sha256:" + "a" * 64):
            self.assertEqual([], pins.violations("uses: " + value))
        for value in ("docker://alpine:latest", "docker://alpine@sha256:" + "a" * 63):
            self.assertTrue(pins.violations("uses: " + value))

    def test_scalar_text_is_not_an_action_reference(self):
        self.assertEqual([], pins.violations('run: echo "Release uses: tag"'))
        self.assertEqual([], pins.violations('- name: "Check uses: value"'))
        text = '  - run: |\n      uses: example text\n    name: example\n  - uses: owner/action@v1'
        self.assertEqual([4], [n for n, _ in pins.violations(text)])

    def test_folded_action_reference_fails_closed(self):
        self.assertTrue(pins.violations("uses: >-\n  actions/checkout@v4"))

    def test_multiline_flow_mapping_fails_closed(self):
        for text in ("- {name: Checkout\n  , uses: actions/checkout@v4}",
                     "- {\n  name: Checkout, uses: actions/checkout@v4\n  }"):
            self.assertTrue(pins.violations(text))

    def test_quoted_flow_mapping_key_fails_closed(self):
        self.assertTrue(pins.violations('"steps": [{uses: actions/checkout@v4}]'))

    def test_node_properties_aliases_and_explicit_keys_fail_closed(self):
        for text in (r'- "us\u0065s": actions/checkout@v4', '- &co uses: actions/checkout@v4', '- &co {uses: actions/checkout@v4}',
                     '- !!map {uses: actions/checkout@v4}', 'steps: &s [{uses: actions/checkout@v4}]',
                     'steps : [{uses: actions/checkout@v4}]', '- *step', 'steps: *steps',
                     '? uses\n: actions/checkout@v4', '- <<: *step'):
            with self.subTest(text=text): self.assertTrue(pins.violations(text))

    def test_harmless_option_maps_are_accepted(self):
        for text in ('permissions: {}', 'env: {}', 'with: { fetch-depth: 0 }',
                     'permissions: { contents: read }'):
            with self.subTest(text=text): self.assertEqual([], pins.violations(text))
        for text in ('with: { uses: actions/checkout@v4 }', 'env: { x: {uses: x} }'):
            self.assertTrue(pins.violations(text))

    def test_quoted_scalar_filters_are_not_keys(self):
        for text in ("- 'v*'", "- 'src/**' # path filter", '- "scripts/**"'):
            self.assertEqual([],pins.violations(text))

    def test_flow_sequence_mappings_are_rejected(self):
        for text in ('steps: [uses: actions/checkout@v4]',
                     'steps: [name: Checkout, uses: actions/checkout@v4]',
                     'steps: [\n name: Checkout, uses: actions/checkout@v4\n]',
                     '- [uses: actions/checkout@v4]'):
            self.assertTrue(pins.violations(text),text)
        self.assertEqual([],pins.violations('branches: [main, release]'))

    def test_encoded_option_keys_are_rejected(self):
        for key in ('with','env','permissions'):
            self.assertTrue(pins.violations(key + r': { "\x75ses": org/repo/.github/workflows/w.yml@main }'))

    def test_block_scalar_uses_actual_key_column(self):
        for spaces in (1,2,3,5):
            text='  -'+' '*spaces+'name: |\n'+' '*(spaces+6)+'Checkout\n'+' '*(spaces+3)+'uses: actions/checkout@v4'
            self.assertEqual([3],[n for n,_ in pins.violations(text)])

    def test_comments_are_ignored(self):
        self.assertEqual([], pins.violations(" # uses: owner/action@v1"))


if __name__ == "__main__":
    unittest.main()
