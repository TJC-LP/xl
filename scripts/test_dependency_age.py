#!/usr/bin/env python3
import unittest
from datetime import datetime, timedelta, timezone
from unittest.mock import patch
import dependency_age as age

class PolicyTests(unittest.TestCase):
    def test_cross_versions_and_comments(self):
        pins = age.extract('mvn"a:b:1.0"\nmvn"a::b:1.0" // pinned: compatibility\nmvn"a:::b:1.0"\n// mvn"ignored:x:1.0"', 'test', '3.9.0')
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
            pin = age.Pin('a', 'b', '1.0', 'test:1', reason)
            self.assertEqual(passed, age.evaluate(pin, now-delta, now)['passed'])
        with self.assertRaises(ValueError):
            age.evaluate(pin, now+timedelta(seconds=1), now)

    def test_blank_or_unrelated_comment_is_not_justification(self):
        for comment in ('// pinned: ', '// unrelated pinned: text', '// normal'):
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
        for comment, expected in [('', False), (' // pinned: ', False),
                                   (' // unrelated pinned: reason', False),
                                   (' // pinned: baseline stability', True)]:
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
        for extra, expected in [('', ''), ('// mill-version pinned: \n', ''),
                                ('// unrelated mill-version pinned: text\n', ''),
                                ('// mill-version pinned: wait for launcher digests\n',
                                 'wait for launcher digests')]:
            with self.subTest(extra=extra):
                self.assertEqual(expected, age.mill_reason(header + extra + 'import mill._\n'))
        with self.assertRaises(ValueError):
            age.mill_reason(header + '// mill-version pinned: a\n// mill-version pinned: b\n')

    def test_old_mill_needs_a_reason(self):
        now = datetime(2027, 4, 1, tzinfo=timezone.utc)
        for reason, passed in [('', False), ('wait for launcher digests', True)]:
            pin = age.Pin('com.lihaoyi', 'mill-contrib-jmh_3', '1.1.5', 'build.mill:1', reason)
            self.assertEqual(passed, age.evaluate(pin, now-timedelta(days=400), now)['passed'])

    def test_repository_inventory_includes_tooling(self):
        pins = age.inventory(age.ROOT)
        self.assertGreater(len(pins), 28)
        self.assertIn('mill-contrib-jmh_3', [p.artifact for p in pins])
        self.assertEqual(2, len([p for p in pins if p.group == 'org.openjdk.jmh']))
        self.assertIn('scala3-compiler_3', [p.artifact for p in pins])

if __name__ == '__main__': unittest.main()
