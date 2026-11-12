#!/usr/bin/env python3
"""Synthetic setup regressions: no real credentials, network or downloaded tools."""
import hashlib
import os
from pathlib import Path
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]

class RemoteSetupTest(unittest.TestCase):
    def run_setup(self, system='Linux', arch='x86_64', cached=False,
                  external=False, version='2.1.24', payload=False, tamper=False, user_owned=False):
        with tempfile.TemporaryDirectory() as td:
            t = Path(td)
            project = t / 'project'; (project / 'scripts').mkdir(parents=True)
            for name in ['remote-setup.sh', 'verify-download.sh', 'launcher-checksums.txt']:
                (project / 'scripts' / name).write_bytes((ROOT / 'scripts' / name).read_bytes())
            for name in ['.mill-version', '.mill-jvm-version']:
                (project / name).write_bytes((ROOT / name).read_bytes())
            tools = t / 'tools'; tools.mkdir()
            managed = t / 'managed'; managed.mkdir()
            private = managed / '.xl-remote'; private.mkdir()
            marker = t / 'executed'
            def executable(path, body):
                path.write_text('#!/bin/sh\n' + body + '\n'); path.chmod(0o755)
            executable(tools / 'uname', f'case "$1" in -s) echo {system};; -m) echo {arch};; esac')
            executable(tools / 'java', 'exit 1')
            executable(tools / 'scala-cli', 'exit 0')
            executable(tools / 'curl', 'echo called >> "$FIXTURE/curl-called"; exit 7')
            if cached:
                executable(private / 'cs', 'echo bad >> "$FIXTURE/executed"; exit 1')
            if user_owned:
                executable(managed / 'cs', 'echo user-owned >> "$FIXTURE/executed"; exit 1')
            if external:
                executable(tools / 'cs', 'echo trusted >> "$FIXTURE/executed"; exit 1')
            if payload:
                import gzip
                native = b'#!/bin/sh\necho verified >> "$FIXTURE/executed"\nexit 1\n'
                archive = gzip.compress(native)
                (t / 'payload.gz').write_bytes(archive if not tamper else b'wrong bytes')
                manifest = project / 'scripts/launcher-checksums.txt'
                manifest.write_text(hashlib.sha256(archive).hexdigest() + '  cs-2.1.24-x86_64-pc-linux.gz\n' + hashlib.sha256(native).hexdigest() + '  cs-2.1.24-x86_64-pc-linux\n')
                executable(tools / 'curl', 'cp "$FIXTURE/payload.gz" "$4"')
            env = {'PATH':str(managed)+':'+str(tools)+':/usr/bin:/bin','HOME':str(t),
                   'FIXTURE':str(t),'CLAUDE_PROJECT_DIR':str(project),
                   'CLAUDE_CODE_REMOTE':'true','XL_REMOTE_BIN':str(managed),
                   'CLAUDE_ENV_FILE':str(t/'exports'), 'XL_REMOTE_COURSIER_VERSION':version}
            p = subprocess.run(['/bin/bash',str(project/'scripts/remote-setup.sh'),'--force'],
                               env=env,capture_output=True,text=True,timeout=10)
            if user_owned:
                self.assertTrue((managed/'cs').exists(), 'caller tool was deleted')
                self.assertIn('user-owned', (managed/'cs').read_text())
            return p, (t/'exports').exists(), marker.read_text() if marker.exists() else '', (t/'curl-called').exists()

    def test_bad_cached_tool_is_not_executed_and_exports_continue(self):
        p, exports, executed, _ = self.run_setup(cached=True)
        self.assertEqual(0,p.returncode,p.stderr); self.assertTrue(exports)
        self.assertEqual('',executed); self.assertIn('WARN',p.stderr)

    def test_bad_cache_is_replaced_only_with_verified_download(self):
        for tamper in [False,True]:
            with self.subTest(tamper=tamper):
                p,exports,executed,_=self.run_setup(cached=True,payload=True,tamper=tamper)
                self.assertEqual(0,p.returncode,p.stderr);self.assertTrue(exports)
                self.assertNotIn('bad',executed)
                self.assertEqual(not tamper,'verified' in executed)

    def test_unknown_version_degrades_without_download(self):
        p,exports,executed,download=self.run_setup(version='2.1.25')
        self.assertEqual(0,p.returncode,p.stderr);self.assertTrue(exports)
        self.assertFalse(download);self.assertEqual('',executed)

    def test_existing_coursier_precedes_arm_fallback(self):
        for system,arch in [('Darwin','arm64'),('Linux','aarch64'),('Darwin','x86_64')]:
            with self.subTest(system=system,arch=arch):
                p,exports,executed,download=self.run_setup(system,arch,external=True)
                self.assertEqual(0,p.returncode,p.stderr);self.assertTrue(exports)
                self.assertIn('trusted',executed);self.assertFalse(download)

    def test_user_owned_default_path_is_preserved(self):
        for arch in ('x86_64', 'aarch64'):
            with self.subTest(arch=arch):
                p,exports,executed,download=self.run_setup(arch=arch,user_owned=True)
                self.assertEqual(0,p.returncode,p.stderr);self.assertTrue(exports)
                self.assertEqual('',executed)
                self.assertEqual(arch == 'x86_64',download)

    def test_arm_without_java_warns_without_download(self):
        p,exports,executed,download=self.run_setup(arch='aarch64')
        self.assertEqual(0,p.returncode,p.stderr);self.assertTrue(exports)
        self.assertFalse(download);self.assertEqual('',executed)
        self.assertIn('working Java',p.stderr)

if __name__ == '__main__': unittest.main()
