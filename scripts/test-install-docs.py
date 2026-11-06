#!/usr/bin/env python3
"""Execute documented installers only against synthetic downloads and tool stubs."""
import importlib.util
import json
import gzip
import hashlib
import os
from pathlib import Path
import re
import subprocess
import tempfile
import unittest
ROOT = Path(__file__).resolve().parents[1]

class InstallDocsTest(unittest.TestCase):
    def check(self, skill, system, arch, bad=False, future=False, manifest="valid"):
        doc=(ROOT/f'plugin/skills/{skill}/reference/INSTALL.md').read_text()
        # The first executable block is the documented native installation.
        script=doc.split('```bash\n',1)[1].split('```',1)[0]
        self.assertTrue(script.startswith('(\n'));self.assertTrue(script.rstrip().endswith(')'))
        with tempfile.TemporaryDirectory() as td:
            t=Path(td);tools=t/'tools';tools.mkdir();target=t/'target';target.mkdir()
            name='xl' if skill=='xl-cli' else 'scala-cli'
            (target/name).write_text('old')
            content=b'synthetic binary';payload=content if name=='xl' else gzip.compress(content)
            (t/'payload').write_bytes(payload)
            if future: script=script.replace('VERSION=0.25.0', 'VERSION=0.26.0')
            digest=hashlib.sha256(payload).hexdigest()
            asset='xl-0.26.0-linux-amd64'
            entry=('0'*64 if bad else digest)+'  '+asset+'\n'
            (t/'manifest').write_text(entry if manifest=='valid' else entry*2 if manifest=='duplicate' else '')
            # Substitute expected digests only for the synthetic fixtures.
            script=re.sub(r'EXPECTED=[0-9a-f]{64}', 'EXPECTED='+('0'*64 if bad else digest),script)
            for command,body in {'uname':f'case "$1" in -s) echo {system};; -m) echo {arch};; esac',
                                 'curl':'case "$2" in */SHA256SUMS) cp "$FIXTURE/manifest" "$4";; *) cp "$FIXTURE/payload" "$4";; esac'}.items():
                p=tools/command;p.write_text('#!/bin/sh\n'+body+'\n');p.chmod(0o755)
            env={'PATH':str(tools)+':/usr/bin:/bin','HOME':str(t),'FIXTURE':str(t),
                 'XL_INSTALL_DIR':str(target),'SCALA_CLI_INSTALL_DIR':str(target)}
            bad = bad or (future and manifest != 'valid')
            # A failed install must not exit the caller or replace its EXIT trap.
            shell='trap "echo caller-trap" EXIT\n'+script+'\nresult=$?\nfalse\necho caller-survives:$result\n'
            p=subprocess.run(['/bin/bash','-c',shell],env=env,capture_output=True,text=True,timeout=10)
            self.assertEqual(0,p.returncode,p.stderr)
            self.assertIn('caller-survives:'+('1' if bad else '0'),p.stdout)
            self.assertIn('caller-trap',p.stdout)
            self.assertEqual(b'old' if bad else content,(target/name).read_bytes())
            self.assertEqual([name],[x.name for x in target.iterdir()])

    def test_cli_platforms_and_mismatch(self):
        for system,arch in [('Linux','x86_64'),('Linux','aarch64'),('Darwin','x86_64'),('Darwin','arm64'),('MINGW64_NT','x86_64')]:
            for bad in [False,True]:
                with self.subTest(system=system,arch=arch,bad=bad):self.check('xl-cli',system,arch,bad)

    def test_linux_scala_cli_without_coursier(self):
        for arch in ['x86_64','aarch64']:
            for bad in [False,True]:
                with self.subTest(arch=arch,bad=bad):self.check('xl-scripting','Linux',arch,bad)

    def test_next_release_requires_unique_manifest_and_matching_bytes(self):
        for bad,manifest in [(False,'valid'),(True,'valid'),(False,'missing'),(False,'duplicate')]:
            with self.subTest(bad=bad,manifest=manifest):
                self.check('xl-cli','Linux','x86_64',bad=bad,future=True,manifest=manifest)

    def test_installer_version_cannot_lag_a_release(self):
        spec=importlib.util.spec_from_file_location('install_version',ROOT/'scripts/check-install-version.py')
        checker=importlib.util.module_from_spec(spec);spec.loader.exec_module(checker)
        with tempfile.TemporaryDirectory() as td:
            root=Path(td)
            for name in ('plugin/.claude-plugin/plugin.json','plugin/skills/xl-cli/SKILL.md',
                         'plugin/skills/xl-cli/reference/INSTALL.md'):
                dest=root/name;dest.parent.mkdir(parents=True,exist_ok=True);dest.write_bytes((ROOT/name).read_bytes())
            self.assertEqual([],checker.violations(root))
            plugin=root/'plugin/.claude-plugin/plugin.json';data=json.loads(plugin.read_text());data['version']='0.26.0';plugin.write_text(json.dumps(data))
            skill=root/'plugin/skills/xl-cli/SKILL.md';skill.write_text(skill.read_text().replace('Requires xl >= 0.25.0','Requires xl >= 0.26.0'))
            self.assertTrue(checker.violations(root))
            install=root/'plugin/skills/xl-cli/reference/INSTALL.md';install.write_text(install.read_text().replace('VERSION=0.25.0','VERSION=0.26.0'))
            self.assertEqual([],checker.violations(root))

if __name__=='__main__':unittest.main()
