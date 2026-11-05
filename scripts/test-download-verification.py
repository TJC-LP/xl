#!/usr/bin/env python3
"""Offline behavior checks: no real network, credentials, or downloaded code execution."""
import hashlib
import os
from pathlib import Path
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]


class DownloadVerificationTest(unittest.TestCase):
    def test_checksum_helper_accepts_exact_bytes_and_rejects_tampering(self):
        with tempfile.TemporaryDirectory() as td:
            artifact = Path(td) / "artifact"
            artifact.write_bytes(b"synthetic launcher")
            digest = hashlib.sha256(artifact.read_bytes()).hexdigest()
            command = ['sh', '-c', '. "$1"; xl_verify_sha256 "$2" "$3"', 'verify',
                       str(ROOT / 'scripts/verify-download.sh'), digest, str(artifact)]
            self.assertEqual(0, subprocess.run(command, capture_output=True).returncode)
            artifact.write_bytes(b"tampered")
            self.assertNotEqual(0, subprocess.run(command, capture_output=True).returncode)

    def test_cached_mill_is_checked_before_execution(self):
        with tempfile.TemporaryDirectory() as td:
            cached = Path(td) / '1.1.5'
            marker = Path(td) / 'executed'
            cached.write_text('#!/bin/sh\ntouch "' + str(marker) + '"\n')
            cached.chmod(0o755)
            env = dict(os.environ, MILL_VERSION='1.1.5-jvm', MILL_FINAL_DOWNLOAD_FOLDER=td)
            result = subprocess.run(['sh', str(ROOT / 'mill'), '--version'], cwd=ROOT,
                                    env=env, capture_output=True, text=True)
            self.assertNotEqual(0, result.returncode)
            self.assertIn('SHA-256 mismatch', result.stderr)
            self.assertFalse(marker.exists())

    def test_unknown_mill_version_fails_closed(self):
        env = dict(os.environ, MILL_VERSION='0.0.0-jvm')
        result = subprocess.run(['sh', str(ROOT / 'mill'), '--version'], cwd=ROOT,
                                env=env, capture_output=True, text=True)
        self.assertNotEqual(0, result.returncode)
        self.assertIn('No committed checksum', result.stderr)

    def test_installer_preserves_existing_binary_on_bad_or_missing_manifest(self):
        platforms = [('Linux', 'linux-amd64'), ('MINGW64_NT-10.0', 'windows-amd64.exe'),
                     ('MSYS_NT-10.0', 'windows-amd64.exe'), ('CYGWIN_NT-10.0', 'windows-amd64.exe')]
        for system, asset in platforms:
            self.check_installer_platform(system, asset)

    def check_installer_platform(self, system, asset):
        for mode in ('valid', 'bad', 'missing', 'duplicate'):
            with self.subTest(system=system, mode=mode), tempfile.TemporaryDirectory() as td:
                temp = Path(td)
                mocks = temp / 'bin'; mocks.mkdir()
                target = temp / 'install'; target.mkdir()
                (target / 'xl').write_text('old binary')
                payload = b'synthetic new binary'
                (temp / 'payload').write_bytes(payload)
                digest = hashlib.sha256(payload).hexdigest()
                manifest = digest + f'  xl-0.0.0-{asset}\n'
                if mode == 'bad': manifest = '0' * 64 + f'  xl-0.0.0-{asset}\n'
                if mode == 'missing': manifest = ''
                if mode == 'duplicate': manifest += manifest
                (temp / 'manifest').write_text(manifest)
                (mocks / 'uname').write_text('#!/bin/sh\ncase "$1" in -s) echo ' + system + ';; -m) echo x86_64;; esac\n')
                (mocks / 'curl').write_text('#!/bin/sh\ncase "$2" in */SHA256SUMS) cp "$XL_TEST_FIXTURES/manifest" "$4";; *) cp "$XL_TEST_FIXTURES/payload" "$4";; esac\n')
                for file in mocks.iterdir(): file.chmod(0o755)
                script = temp / 'installer.sh'
                script.write_text((ROOT / 'scripts/install-cli.sh').read_text().replace('__VERSION__', '0.0.0'))
                env = dict(os.environ, PATH=str(mocks) + os.pathsep + os.environ['PATH'],
                           XL_INSTALL_DIR=str(target), XL_TEST_FIXTURES=td)
                result = subprocess.run(['bash', str(script)], env=env, capture_output=True)
                if mode == 'valid':
                    self.assertEqual(0, result.returncode, result.stderr)
                    self.assertEqual(payload, (target / 'xl').read_bytes())
                else:
                    self.assertNotEqual(0, result.returncode)
                    self.assertEqual('old binary', (target / 'xl').read_text())
                self.assertEqual(['xl'], [p.name for p in target.iterdir()])

    def test_download_paths_are_wired_to_verification(self):
        mill = (ROOT / 'mill').read_text()
        self.assertLess(mill.index('xl_verify_sha256 "$XL_MILL_SHA" "$MILL_TEMP_DOWNLOAD_FILE"'),
                        mill.index('chmod +x "${MILL_TEMP_DOWNLOAD_FILE}"'))
        remote = (ROOT / 'scripts/remote-setup.sh').read_text()
        self.assertNotIn('coursier/launchers/raw/master', remote)
        self.assertIn('xl_verify_sha256 "$CS_JAR_SHA" "$CS_JAR_PATH.tmp" || exit 1', remote)
        self.assertIn('xl_verify_sha256 "$CS_ARCHIVE_SHA" "$BIN/cs.gz.tmp" || exit 1', remote)
        release = (ROOT / '.github/workflows/release.yml').read_text()
        self.assertIn('sha256sum xl-* > SHA256SUMS', release)
        self.assertIn('release-assets/SHA256SUMS', release)
        self.assertIn('cp scripts/install-cli.sh skill-dist/install-cli.sh', release)

    def test_committed_digest_table_is_well_formed(self):
        rows = [line.split() for line in (ROOT / 'scripts/launcher-checksums.txt').read_text().splitlines()
                if line and not line.startswith('#')]
        self.assertEqual(8, len(rows))
        self.assertEqual(len(rows), len({row[1] for row in rows}))
        for digest, name in rows:
            self.assertRegex(digest, r'^[0-9a-f]{64}$')


if __name__ == '__main__':
    unittest.main()
