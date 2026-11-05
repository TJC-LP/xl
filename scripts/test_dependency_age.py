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

    def test_repository_inventory_includes_tooling(self):
        pins = age.inventory(age.ROOT)
        self.assertGreater(len(pins), 28)
        self.assertIn('mill-contrib-jmh_3', [p.artifact for p in pins])
        self.assertIn('scala3-compiler_3', [p.artifact for p in pins])

if __name__ == '__main__': unittest.main()
