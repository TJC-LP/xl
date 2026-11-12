#!/usr/bin/env python3
import unittest
from datetime import date, datetime, timedelta, timezone
from unittest.mock import patch
import dependency_age as age

class PolicyTests(unittest.TestCase):
    def test_cross_versions_and_comments(self):
        pins = age.extract('mvn"a:b:1.0"\nmvn"a::b:1.0" // pinned until 2026-12-01: compatibility\nmvn"a:::b:1.0"\n// mvn"ignored:x:1.0"', 'test', '3.9.0')
        self.assertEqual(['b', 'b_3', 'b_3.9.0'], [p.artifact for p in pins])
        self.assertEqual('compatibility', pins[1].reason)

    def test_unhandled_expression_is_not_silently_skipped(self):
        for text in ['mvn"a:b:$version"', 'mvn"a:b:1.0;classifier=sources"']:
            with self.subTest(text=text), self.assertRaises(ValueError):
                age.extract(text, 'test', '3.9.0')

    def test_age_boundary_and_justification(self):
        now = datetime(2026, 11, 5, tzinfo=timezone.utc)
        for reason, delta, passed in [('', timedelta(days=365), True),
                                      ('', timedelta(days=365, seconds=1), False),
                                      ('compatibility', timedelta(days=800), True)]:
            pin = age.Pin('a', 'b', '1.0', 'test:1', reason, date(2026, 12, 1) if reason else None)
            self.assertEqual(passed, age.evaluate(pin, now-delta, now)['passed'])
        with self.assertRaises(ValueError):
            age.evaluate(pin, now+timedelta(seconds=1), now)

    def test_unrelated_comment_is_not_justification(self):
        # A blank or undated `pinned` comment is an error (ReviewFollowupTests).
        for comment in ('// unrelated pinned until 2026-12-01: text', '// normal'):
            self.assertFalse(age.extract('mvn"a:b:1" '+comment, 'test', '3.9.0')[0].reason)

    def test_missing_timestamp_and_network_failure_are_errors(self):
        with patch.object(age, 'urlopen') as request:
            request.return_value.__enter__.return_value.headers = {}
            with self.assertRaises(ValueError): age.modified('https://example.test/a')
            request.side_effect = OSError('network unavailable')
            with self.assertRaises(OSError): age.modified('https://example.test/a')


    def test_jmh_artifacts_and_age_exemption(self):
        header = 'object `package` extends build.XLInternalModule with contrib.jmh.JmhModule {'
        now = datetime(2026, 11, 6, tzinfo=timezone.utc)
        for comment, expected in [('', False),
                                   (' // unrelated pinned until 2026-12-01: reason', False),
                                   (' // pinned until 2026-12-01: baseline stability', True)]:
            pins = age.extract_jmh(header + '\n  def jmhCoreVersion = "1.35"' + comment + '\n}')
            self.assertEqual(['jmh-core', 'jmh-generator-bytecode'], [p.artifact for p in pins])
            self.assertEqual(['1.35', '1.35'], [p.version for p in pins])
            for pin in pins:
                self.assertEqual(expected, age.evaluate(pin, now-timedelta(days=800), now)['passed'])
                with self.assertRaises(ValueError):
                    age.evaluate(pin, now+timedelta(seconds=1), now)

    def test_jmh_comments_and_fail_closed_declarations(self):
        header = 'object `package` extends build.XLInternalModule with contrib.jmh.JmhModule {'
        literal = 'def jmhCoreVersion = "1.35"'
        comments = '// def jmhGeneratorByteCodeVersion = "1"\n/*\ndef jmhCoreVersion = "9"\n*/\n'
        self.assertEqual('1.35', age.extract_jmh(comments+header+'\n'+comments+literal+'\n}')[0].version)
        invalid = ['', literal+'\n    .replace("1.35", "1.37")',
                   literal+'\n    + "0"', literal+' + "x"', literal+'; println("x")',
                   literal+'\n'+literal, literal+'\ndef jmhGeneratorByteCodeVersion = "1"',
                   'def helper = {\n'+literal+'\n}', '// '+literal,
                   '/* '+literal+' */', '/* comment */ '+literal,
                   '/* unclosed', 'def jmhCoreVersion = version']
        for body in invalid:
            with self.subTest(body=body), self.assertRaises(ValueError):
                age.extract_jmh(header+'\n'+body+'\n}')

    def test_mill_pin_reason(self):
        header = '//| mvnDeps: ["com.lihaoyi::mill-contrib-jmh:$MILL_VERSION"]\n'
        for extra, expected in [('', ('', None)),
                                ('// unrelated mill-version pinned until 2026-12-01: text\n', ('', None)),
                                ('// mill-version pinned until 2026-12-01: wait for launcher digests\n',
                                 ('wait for launcher digests', date(2026, 12, 1)))]:
            with self.subTest(extra=extra):
                self.assertEqual(expected, age.mill_reason(header + extra + 'import mill._\n'))
        for extra in ('// mill-version pinned: \n',
                      '// mill-version pinned until 2026-12-01: a\n// mill-version pinned until 2026-12-01: b\n'):
            with self.subTest(extra=extra), self.assertRaises(ValueError):
                age.mill_reason(header + extra)

    def test_old_mill_needs_a_reason(self):
        now = datetime(2027, 4, 1, tzinfo=timezone.utc)
        for reason, passed in [('', False), ('wait for launcher digests', True)]:
            pin = age.Pin('com.lihaoyi', 'mill-contrib-jmh_3', '1.1.5', 'build.mill:1', reason,
                          date(2027, 4, 30) if reason else None)
            self.assertEqual(passed, age.evaluate(pin, now-timedelta(days=400), now)['passed'])

    def test_repository_inventory_includes_tooling(self):
        pins = age.inventory(age.ROOT)
        self.assertGreater(len(pins), 28)
        self.assertIn('mill-contrib-jmh_3', [p.artifact for p in pins])
        self.assertEqual(2, len([p for p in pins if p.group == 'org.openjdk.jmh']))
        self.assertIn('scala3-compiler_3', [p.artifact for p in pins])

class ReviewFollowupTests(unittest.TestCase):
    """#839 follow-ups from the #848 review."""

    NOW = datetime(2026, 11, 11, tzinfo=timezone.utc)
    OLD = timedelta(days=800)

    def pin(self, comment):
        return age.extract('mvn"a:b:1.0"' + comment, 'test', '3.9.0')[0]

    # Pin reasons expire on a review date.
    def test_dated_reason_passes_until_its_review_date(self):
        pin = self.pin(' // pinned until 2026-12-01: compatibility')
        self.assertEqual('compatibility', pin.reason)
        self.assertTrue(age.evaluate(pin, self.NOW - self.OLD, self.NOW)['passed'])
        after = datetime(2026, 12, 2, tzinfo=timezone.utc)
        self.assertFalse(age.evaluate(pin, after - self.OLD, after)['passed'])

    def test_review_date_is_the_last_day_it_counts(self):
        pin = self.pin(' // pinned until 2026-11-11: compatibility')
        self.assertTrue(age.evaluate(pin, self.NOW - self.OLD, self.NOW)['passed'])

    def test_undated_or_malformed_pinned_comment_is_an_error(self):
        for comment in (' // pinned: compatibility', ' // pinned until 2026-13-01: x',
                        ' // pinned until 2026-12-01:', ' // pinned until soon: x',
                        ' // pinned until 2026-12-01 compatibility'):
            with self.subTest(comment=comment), self.assertRaises(ValueError):
                self.pin(comment)

    def test_review_date_more_than_a_year_ahead_is_an_error(self):
        pin = self.pin(' // pinned until 2027-11-13: compatibility')
        with self.assertRaises(ValueError):
            age.evaluate(pin, self.NOW - self.OLD, self.NOW)
        ok = self.pin(' // pinned until 2027-11-12: compatibility')
        self.assertTrue(age.evaluate(ok, self.NOW - self.OLD, self.NOW)['passed'])

    def test_fresh_pin_needs_no_reason(self):
        self.assertTrue(age.evaluate(self.pin(''), self.NOW - timedelta(days=10), self.NOW)['passed'])

    def test_jmh_scala_and_mill_reasons_use_the_dated_form(self):
        header = 'object `package` extends build.XLInternalModule with contrib.jmh.JmhModule {'
        jmh = age.extract_jmh(header + '\n  def jmhCoreVersion = "1.35" // pinned until 2026-12-01: baseline\n}')
        self.assertEqual(['baseline', 'baseline'], [p.reason for p in jmh])
        self.assertTrue(age.evaluate(jmh[0], self.NOW - self.OLD, self.NOW)['passed'])
        with self.assertRaises(ValueError):
            age.extract_jmh(header + '\n  def jmhCoreVersion = "1.35" // pinned: baseline\n}')
        mill = age.mill_reason('// mill-version pinned until 2026-12-01: launcher digests\n')
        self.assertEqual(('launcher digests', age.date(2026, 12, 1)), mill)
        with self.assertRaises(ValueError):
            age.mill_reason('// mill-version pinned: launcher digests\n')

    # Block comments are found before line comments, outside string literals.
    def test_url_inside_block_comment_does_not_hide_its_end(self):
        for prefix in ('/** See https://fs2.io */\n', '/**\n * Docs: https://mill-build.org */\n',
                       '/* a */ // b\n'):
            with self.subTest(prefix=prefix):
                pins = age.extract(prefix + 'mvn"a:b:1.0"\n', 'test', '3.9.0')
                self.assertEqual(['b'], [p.artifact for p in pins])

    def test_comment_delimiters_inside_strings_are_code(self):
        pins = age.extract('val url = "https://x/*"\nmvn"a:b:1.0"\n', 'test', '3.9.0')
        self.assertEqual(['b'], [p.artifact for p in pins])

    def test_dependency_inside_or_beside_block_comment_still_fails_closed(self):
        for text in ('/*\nmvn"a:b:1.0"\n*/', '/* x */ mvn"a:b:1.0"', 'mvn"a:b:1.0" /* x */',
                     '/* unclosed\n', 'stray */\n', '/* a /* nested */ */\n'):
            with self.subTest(text=text), self.assertRaises(ValueError):
                age.extract(text, 'test', '3.9.0')

    def test_jmh_url_inside_block_comment(self):
        header = 'object `package` extends build.XLInternalModule with contrib.jmh.JmhModule {'
        text = '/** See https://openjdk.org/jmh */\n' + header + '\n  def jmhCoreVersion = "1.35"\n}'
        self.assertEqual('1.35', age.extract_jmh(text)[0].version)

    # The live age check cannot skip unrelated CI steps.
    def test_age_check_runs_in_its_own_ci_job(self):
        workflow = (age.ROOT / '.github/workflows/ci.yml').read_text()
        jobs = age.re.split(r'(?m)^  (?=[A-Za-z0-9_-]+:\s*$)', workflow.split('\njobs:\n', 1)[1])
        [job] = [j for j in jobs if 'scripts/dependency_age.py' in j]
        steps = age.re.findall(r'(?m)^      - (?:name|uses): (.+)$', job)
        self.assertEqual(2, len(steps), steps)  # checkout + the age policy
        self.assertTrue(steps[0].startswith('actions/checkout@'), steps)

    # Renovate derives the full-cross suffix from build.mill's Scala version.
    def test_full_cross_dependencies_follow_the_scala_version(self):
        import json
        config = json.loads((age.ROOT / 'renovate.json').read_text())
        [manager] = [m for m in config['customManagers'] if ':::' in ''.join(m['matchStrings'])]
        self.assertEqual(['/^build\\.mill$/'], manager['managerFilePatterns'])
        self.assertEqual('combination', manager.get('matchStringsStrategy'))
        self.assertIn('{{{scalaVersion}}}', manager['depNameTemplate'])
        self.assertNotIn('3.9.0', manager['depNameTemplate'])

    def test_full_cross_dependencies_are_limited_to_one_in_build_mill(self):
        with self.assertRaises(ValueError):
            age.check_full_cross({'build.mill': 'mvn"a:::b:1"\nmvn"c:::d:1"\n'})
        with self.assertRaises(ValueError):
            age.check_full_cross({'build.mill': '', 'x/package.mill': 'mvn"a:::b:1"\n'})
        age.check_full_cross({'build.mill': 'mvn"a:::b:1"\n', 'x/package.mill': 'mvn"a::b:1"\n'})


if __name__ == '__main__': unittest.main()
