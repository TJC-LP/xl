#!/usr/bin/env python3
"""Execute documented installers only against synthetic downloads and tool stubs."""
import importlib.util
import json
import gzip
import hashlib
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import unittest
ROOT = Path(__file__).resolve().parents[1]

def next_release(version):
    major, minor, patch = map(int, version.split('.'))
    return f"{major}.{minor + 1}.0"


def set_installer_version(text, version):
    return re.sub(r'^VERSION=[0-9]+\.[0-9]+\.[0-9]+$',
                  lambda _: 'VERSION=' + version, text, flags=re.MULTILINE)


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
            version = None
            if name == 'xl':
                version = re.search(r'^VERSION=([0-9]+\.[0-9]+\.[0-9]+)$', script, re.MULTILINE).group(1)
                if future:
                    version = next_release(version)
                    script = set_installer_version(script, version)
            digest=hashlib.sha256(payload).hexdigest()
            platforms = {('Linux','x86_64'):'linux-amd64', ('Linux','aarch64'):'linux-arm64',
                         ('Darwin','x86_64'):'darwin-amd64', ('Darwin','arm64'):'darwin-arm64',
                         ('MINGW64_NT','x86_64'):'windows-amd64.exe'}
            asset=f'xl-{version}-{platforms[(system,arch)]}'
            entry=('0'*64 if bad else digest)+'  '+asset+'\n'
            (t/'manifest').write_text(entry if manifest=='valid' else entry*2 if manifest=='duplicate' else '')
            # Substitute expected digests only for the synthetic fixtures.
            script=re.sub(r'EXPECTED=[0-9a-f]{64}', 'EXPECTED='+('0'*64 if bad else digest),script)
            for command,body in {'uname':f'case "$1" in -s) echo {system};; -m) echo {arch};; esac',
                                 'curl':'case "$2" in */SHA256SUMS) cp "$FIXTURE/manifest" "$4";; *) cp "$FIXTURE/payload" "$4";; esac'}.items():
                p=tools/command;p.write_text('#!/bin/sh\n'+body+'\n');p.chmod(0o755)
            env={'PATH':str(tools)+':/usr/bin:/bin','HOME':str(t),'FIXTURE':str(t),
                 'XL_INSTALL_DIR':str(target),'SCALA_CLI_INSTALL_DIR':str(target)}
            bad = bad or (name == 'xl' and version != '0.25.0' and manifest != 'valid')
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
            plugin=root/'plugin/.claude-plugin/plugin.json';data=json.loads(plugin.read_text());version=next_release(data['version']);data['version']=version;plugin.write_text(json.dumps(data))
            skill=root/'plugin/skills/xl-cli/SKILL.md';skill.write_text(re.sub(r'Requires xl >= [0-9]+\.[0-9]+\.[0-9]+', lambda _: 'Requires xl >= '+version, skill.read_text()))
            self.assertTrue(checker.violations(root))
            install=root/'plugin/skills/xl-cli/reference/INSTALL.md';install.write_text(set_installer_version(install.read_text(),version))
            self.assertEqual([],checker.violations(root))

    def test_entire_fixture_matrix_survives_release_bumps(self):
        version=json.loads((ROOT/'plugin/.claude-plugin/plugin.json').read_text())['version']
        # Exercise two releases, rather than only the first transition away from
        # the legacy inline-digest release. Run named tests to avoid recursion.
        tests=['test_cli_platforms_and_mismatch', 'test_linux_scala_cli_without_coursier',
               'test_next_release_requires_unique_manifest_and_matching_bytes',
               'test_installer_version_cannot_lag_a_release']
        for _ in range(2):
            version=next_release(version)
            with self.subTest(version=version), tempfile.TemporaryDirectory() as td:
                root=Path(td)
                for name in ('scripts/test-install-docs.py','scripts/check-install-version.py',
                             'plugin/.claude-plugin/plugin.json','plugin/skills/xl-cli/SKILL.md',
                             'plugin/skills/xl-cli/reference/INSTALL.md',
                             'plugin/skills/xl-scripting/reference/INSTALL.md'):
                    dest=root/name;dest.parent.mkdir(parents=True,exist_ok=True);dest.write_bytes((ROOT/name).read_bytes())
                plugin=root/'plugin/.claude-plugin/plugin.json';data=json.loads(plugin.read_text());data['version']=version;plugin.write_text(json.dumps(data))
                skill=root/'plugin/skills/xl-cli/SKILL.md';skill.write_text(re.sub(r'Requires xl >= [0-9]+\.[0-9]+\.[0-9]+',lambda _: 'Requires xl >= '+version,skill.read_text()))
                install=root/'plugin/skills/xl-cli/reference/INSTALL.md';install.write_text(set_installer_version(install.read_text(),version))
                result=subprocess.run([sys.executable,'-B',str(root/'scripts/test-install-docs.py'),
                                       *['InstallDocsTest.'+test for test in tests]],
                                      capture_output=True,text=True,timeout=60)
                self.assertEqual(0,result.returncode,result.stdout+result.stderr)

if __name__=='__main__':unittest.main()
