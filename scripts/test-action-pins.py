#!/usr/bin/env python3
import importlib.util
import unittest
from pathlib import Path

spec = importlib.util.spec_from_file_location("pins", Path(__file__).with_name("check-action-pins.py"))
pins = importlib.util.module_from_spec(spec)
spec.loader.exec_module(pins)


class ActionPinsTest(unittest.TestCase):
    def test_reference_quotes_must_pair(self):
        for opening in ('', "'", '"'):
            for closing in ('', "'", '"'):
                for ref in ('a' * 40, 'main', 'v1.2.3', 'a' * 39, 'a' * 41):
                    with self.subTest(opening=opening, closing=closing, ref=ref):
                        text = '- uses: ' + opening + 'owner/action@' + ref + closing + ' # v1.2.3'
                        self.assertEqual(opening != closing or ref != 'a' * 40,
                                         bool(pins.violations(text)))

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
                     '- [uses: actions/checkout@v4]', 'branches: [*alias]', 'branches: [main, *alias]'):
            self.assertTrue(pins.violations(text),text)
        self.assertEqual([],pins.violations('branches: [main, release]'))

    def test_encoded_option_keys_are_rejected(self):
        for key in ('with','env','permissions'):
            self.assertTrue(pins.violations(key + r': { "\x75ses": org/repo/.github/workflows/w.yml@main }'))

    def test_block_scalar_uses_actual_key_column(self):
        for spaces in (1,2,3,5):
            text='  -'+' '*spaces+'name: |\n'+' '*(spaces+6)+'Checkout\n'+' '*(spaces+3)+'uses: actions/checkout@v4'
            self.assertEqual([3],[n for n,_ in pins.violations(text)])


    def test_leading_bom_does_not_hide_structure(self):
        self.assertEqual([], pins.violations('\ufeffuses: owner/action@' + 'a' * 40 + ' # v1.2.3'))
        for text in ('\ufeffuses: owner/action@main',
                     '\ufeffjobs: {x: {uses: org/r/.github/workflows/w.yml@main}}\non: push',
                     '# prefix\n\ufeffuses: owner/action@main'):
            self.assertTrue(pins.violations(text), text)

    def test_document_forms_and_unknown_structure_fail_closed(self):
        for text in ('--- {on: push, jobs: {x: {uses: org/r/.github/workflows/w.yml@main}}}',
                     '---\nuses: owner/action@main', '...', '%YAML 1.2',
                     ', "\\x75ses": owner/action@main }',
                     '"us\\x65s": owner/action@main',
                     '- unexpected key: {uses: owner/action@main}',
                     '- - uses: owner/action@main', '\tuses: owner/action@main'):
            with self.subTest(text=text): self.assertTrue(pins.violations(text))

    def test_comment_cannot_close_an_option_map(self):
        for key in ('env', 'with', 'permissions'):
            for tail in ('uses', r'"\x75ses"'):
                text = key + ': { a: 1 # }\n  , ' + tail + ': owner/action@main }'
                with self.subTest(text=text): self.assertTrue(pins.violations(text))
            self.assertTrue(pins.violations(key + ': { a: 1 # }'))
            self.assertEqual([], pins.violations(key + ': { a: 1 } # { uses: harmless }'))
            self.assertEqual([], pins.violations(key + ': { }'))

    def test_uses_precedes_every_harmless_value_exemption(self):
        for key in ('uses', '"uses"', "'uses'", 'uses '):
            for value in ('', '{}', '{ }', '[]', '[main]', '|', '>-', 'null', 'true'):
                with self.subTest(key=key, value=value):
                    self.assertTrue(pins.violations(key + ': ' + value))
        self.assertTrue(pins.violations('"uses": owner/action@' + 'a' * 40 + ' # v1.2.3'))

    def test_scalar_comment_boundaries_preserve_harmless_text(self):
        for text in ('name: "uses: example # literal" # uses: ignored',
                     "name: 'it''s a # uses: example'", '- "uses: a # string"',
                     'run: echo "${GITHUB_REF#refs/tags/v}"',
                     'run: echo "hi" # {uses: harmless}',
                     'if: ${{ github.repository == \'TJC-LP/xl\' }}',
                     '- name: echo "uses: descriptive text"'):
            with self.subTest(text=text): self.assertEqual([], pins.violations(text))
        for text in ('name: "unterminated\n  uses: owner/action@main',
                     "name: 'unterminated\n  uses: owner/action@main",
                     'name: "closed" trailing'):
            self.assertTrue(pins.violations(text), text)

    def test_block_scalar_boundaries_and_modifiers(self):
        for spaces in (1, 3, 5):
            key_indent = 3 + spaces
            for header in ('|', '>', '|-', '>+', '|2', '>2-', '|-2', '|+2'):
                lines = ['  -' + ' ' * spaces + 'name: ' + header,
                         '', ' ' * (key_indent + 2) + '# actual payload',
                         ' ' * (key_indent + 2) + 'uses: harmless payload',
                         ' ' * key_indent + 'uses: owner/action@main']
                text = '\n'.join(lines)
                with self.subTest(spaces=spaces, header=header):
                    self.assertEqual([5], [n for n, _ in pins.violations(text)])
                    lines[-1] = ' ' * key_indent + 'uses: owner/action@' + 'a' * 40 + ' # v1.2.3'
                    self.assertEqual([], pins.violations('\n'.join(lines)))

    def test_implicit_scalar_stops_at_detected_content_dedent(self):
        text = 'name: |\n      harmless\n  uses: owner/action@main'
        self.assertEqual([3], [n for n, _ in pins.violations(text)])
        # This deliberately malformed indentation is refused conservatively;
        # it is not asserted to be an executable GitHub workflow.
        for header in ('|0', '|22', '|+-', '|2-2'):
            self.assertTrue(pins.violations('name: ' + header))

    def test_option_maps_refuse_structural_and_encoded_values(self):
        for value in ('{ x: &anchor value }', '{ x: *alias }', '{ x: [value] }',
                      '{ x: { y: value } }', '{ "uses": owner/action@main }',
                      r'{ "\x75ses": owner/action@main }', '{ uses: value }',
                      '{ x: value,\n  y: value }', '{ x: "# }" }'):
            self.assertTrue(pins.violations('with: ' + value), value)

    def test_sequence_separator_tabs_cannot_hide_structure(self):
        for text in ('- \t*alias', '- \t&anchor', '- \t!tag scalar',
                     '-\t*alias', '- \t{"uses":"owner/action@main"}',
                     '- \tname: value'):
            with self.subTest(text=text): self.assertTrue(pins.violations(text))
        self.assertEqual([], pins.violations('- plain-list-item'))
        self.assertEqual([], pins.violations('- name: "literal\t tab"'))

    def test_actual_workflows_follow_supported_grammar(self):
        root = Path(__file__).resolve().parents[1]
        files = sorted((root / '.github/workflows').glob('*.yml')) + sorted((root / '.github/workflows').glob('*.yaml'))
        self.assertTrue(files)
        for path in files:
            with self.subTest(workflow=path.name):
                self.assertEqual([], pins.violations(path.read_text(encoding='utf-8')))

    def test_comments_are_ignored(self):
        self.assertEqual([], pins.violations(" # uses: owner/action@v1"))


if __name__ == "__main__":
    unittest.main()
