#!/usr/bin/env python3
"""Exercise the native smoke harness without paying for a native rebuild per case."""
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / 'scripts/smoke-native-xml-security.sh'


class NativeXmlSmokeTest(unittest.TestCase):
    def run_smoke(self, mode):
        with tempfile.TemporaryDirectory(prefix='xl smoke ') as tmp:
            fake = Path(tmp) / 'fake cli.py'
            record = Path(tmp) / 'fixture-path'
            fake.write_text('''import pathlib, sys, zipfile
import xml.etree.ElementTree as ET
mode, record, flag, fixture, verb = sys.argv[1:]
assert flag == '-f' and verb == 'sheets'
pathlib.Path(record).write_text(fixture)
with zipfile.ZipFile(fixture) as book:
    xml = book.read('xl/workbook.xml')
    assert len(xml) > 6_000_000
    root = ET.fromstring(xml)
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    names = root.findall('s:definedNames/s:definedName', ns)
    assert len(names) == 120000
    assert names[0].get('name') == 'n0'
    assert names[-1].get('name') == 'n119999'
    assert names[-1].text == 'Sheet1!$A$100'
    assert book.read('xl/worksheets/sheet1.xml')
if mode == 'exit':
    print('parse failed', file=sys.stderr)
    sys.exit(7)
if mode == 'jaxp':
    print('JAXP00010003: entity exceeds limit', file=sys.stderr)
else:
    print('Sheet1')
''')
            result = subprocess.run(
                ['bash', str(SCRIPT), sys.executable, str(fake), mode, str(record)],
                # The script must locate fixtures from outside the repository too.
                cwd=tmp, capture_output=True, text=True,
            )
            self.assertTrue(record.exists(), result.stdout + result.stderr)
            fixture = ROOT / record.read_text()
            self.assertFalse(fixture.parent.exists(), 'temporary fixture was not cleaned up')
            return result

    def test_valid_120k_name_fixture_passes(self):
        result = self.run_smoke('success')
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn('Sheet1', result.stdout)

    def test_nonzero_cli_exit_fails_gate_and_cleans_fixture(self):
        result = self.run_smoke('exit')
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertIn('::error::', result.stdout)

    def test_jaxp_diagnostic_fails_even_with_zero_exit(self):
        result = self.run_smoke('jaxp')
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertIn('JAXP00010003', result.stdout)

    def test_command_is_required(self):
        result = subprocess.run(['bash', str(SCRIPT)], capture_output=True, text=True)
        self.assertEqual(result.returncode, 2)
        self.assertIn('Usage:', result.stderr)


if __name__ == '__main__':
    unittest.main()
