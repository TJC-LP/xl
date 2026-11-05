#!/usr/bin/env python3
"""Fail closed when a literal direct Maven pin exceeds 365 days without a reason."""
import concurrent.futures
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
import json
from pathlib import Path
import re
import subprocess
import sys
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parents[1]
PIN = re.compile(r'mvn"([A-Za-z0-9_.-]+)(:{1,3})([A-Za-z0-9_.-]+):([0-9][A-Za-z0-9_.+-]*)"')

@dataclass(frozen=True)
class Pin:
    group: str
    artifact: str
    version: str
    location: str
    reason: str

    @property
    def url(self):
        return ('https://repo.maven.apache.org/maven2/' + self.group.replace('.', '/') +
                f'/{self.artifact}/{self.version}/{self.artifact}-{self.version}.pom')


def extract(text, filename, scala):
    pins = []
    block = False
    for number, line in enumerate(text.splitlines(), 1):
        # Build declarations use one literal per line. Refuse unhandled active forms.
        code, _, comment = line.partition('//')
        if '/*' in code or block:
            if 'mvn"' in code:
                raise ValueError(f'{filename}:{number}: dependency in block comment context is unsupported')
            block = '*/' not in code
            continue
        if '*/' in code:
            raise ValueError(f'{filename}:{number}: unmatched block comment')
        matches = list(PIN.finditer(code))
        if code.count('mvn"') != len(matches):
            raise ValueError(f'{filename}:{number}: unsupported dependency expression')
        reason = comment.partition('pinned:')[2].strip() if comment.strip().startswith('pinned:') else ''
        for m in matches:
            group, cross, artifact, version = m.groups()
            suffix = {'': '', ':': '', '::': '_' + scala.split('.')[0], ':::': '_' + scala}[cross]
            pins.append(Pin(group, artifact + suffix, version, f'{filename}:{number}', reason))
    if block:
        raise ValueError(f'{filename}: unclosed block comment')
    return pins


def inventory(root):
    build = (root / 'build.mill').read_text()
    match = re.search(r'^val scala3Version = "([0-9.]+)"(?: // pinned: (.+))?$', build, re.M)
    if not match:
        raise ValueError('Unsupported Scala version declaration')
    scala = match[1]
    files = subprocess.check_output(['git', 'ls-files', '-z', '*.mill'], cwd=root).decode().split('\0')
    pins = [p for f in files if f for p in extract((root / f).read_text(), f, scala)]
    if not pins:
        raise ValueError('No direct dependency declarations found')
    # Include the direct toolchain compiler and the build's Mill-versioned JMH plugin.
    pins.append(Pin('org.scala-lang', 'scala3-compiler_3', scala, 'build.mill:scala3Version',
                    match[2].strip() if match[2] else ''))
    headers = [line for line in build.splitlines() if line.startswith('//| mvnDeps:')]
    if headers != ['//| mvnDeps: ["com.lihaoyi::mill-contrib-jmh:$MILL_VERSION"]']:
        raise ValueError('Review changed build-header dependencies')
    pins.append(Pin('com.lihaoyi', 'mill-contrib-jmh_3', (root / '.mill-version').read_text().strip(),
                    'build.mill:1', ''))
    return pins


def modified(url):
    with urlopen(Request(url, method='HEAD'), timeout=20) as response:
        header = response.headers.get('Last-Modified')
        if not header:
            raise ValueError('Missing Maven POM Last-Modified: ' + url)
        date = parsedate_to_datetime(header)
        if date.tzinfo is None:
            raise ValueError('Last-Modified lacks timezone: ' + url)
        return date.astimezone(timezone.utc)


def evaluate(pin, date, now):
    if date > now:
        raise ValueError('Future POM Last-Modified: ' + pin.url)
    age = now - date
    return {'location': pin.location, 'coordinate': f'{pin.group}:{pin.artifact}:{pin.version}',
            'pom': pin.url, 'lastModified': date.isoformat(), 'ageDays': age.days,
            'reason': pin.reason, 'passed': age <= timedelta(days=365) or bool(pin.reason)}


def main():
    try:
        pins = inventory(ROOT)
        now = datetime.now(timezone.utc)
        urls = sorted({p.url for p in pins})
        with concurrent.futures.ThreadPoolExecutor(max_workers=6) as pool:
            dates = dict(zip(urls, pool.map(modified, urls)))
        rows = [evaluate(p, dates[p.url], now) for p in pins]
        print(json.dumps({'observedAt': now.isoformat(), 'pins': rows}, indent=2))
        return 0 if all(r['passed'] for r in rows) else 1
    except Exception as error:
        print(str(error), file=sys.stderr)
        return 1

if __name__ == '__main__':
    sys.exit(main())
