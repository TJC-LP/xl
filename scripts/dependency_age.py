#!/usr/bin/env python3
"""Fail closed when a literal direct Maven pin exceeds 365 days without a reason."""
import concurrent.futures
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
import json
from pathlib import Path
import re
import subprocess
import sys
from typing import Optional
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
    until: Optional[date] = None

    @property
    def url(self):
        return ('https://repo.maven.apache.org/maven2/' + self.group.replace('.', '/') +
                f'/{self.artifact}/{self.version}/{self.artifact}-{self.version}.pom')


REASON = re.compile(r'pinned until (\d{4}-\d{2}-\d{2}):(.*)')


def parse_reason(comment, where):
    """`pinned until YYYY-MM-DD: reason` -> (reason, review date); no pin comment -> ('', None)."""
    text = comment.strip()
    if not re.match(r'pinned\b', text):
        return '', None
    match = REASON.fullmatch(text)
    reason = match[2].strip() if match else ''
    if not reason:
        raise ValueError(f'{where}: write the pin reason as "pinned until YYYY-MM-DD: reason"')
    try:
        return reason, date.fromisoformat(match[1])
    except ValueError:
        raise ValueError(f'{where}: invalid pin review date {match[1]}') from None


def split_comments(line, block, where):
    """Split one line into (code, line comment, block-comment text, still in block, touched a block comment).

    Delimiters are found left to right and ignored inside "..." literals, so a URL's
    `//` inside a block comment cannot hide its `*/`. Nested block comments are refused.
    """
    code, hidden, i, touched, quoted = [], [], 0, block, False
    while i < len(line):
        pair = line[i:i + 2]
        if block:
            if pair == '/*':
                raise ValueError(f'{where}: nested block comment unsupported')
            if pair == '*/':
                block, i = False, i + 2
                continue
            hidden.append(line[i])
            i += 1
        elif quoted:
            if pair in ('\\"', '\\\\'):
                code.append(pair)
                i += 2
                continue
            quoted = line[i] != '"'
            code.append(line[i])
            i += 1
        elif pair == '//':
            return ''.join(code), line[i + 2:], ''.join(hidden), False, touched
        elif pair == '/*':
            block, touched, i = True, True, i + 2
        elif pair == '*/':
            raise ValueError(f'{where}: unmatched block comment')
        else:
            quoted = line[i] == '"'
            code.append(line[i])
            i += 1
    return ''.join(code), '', ''.join(hidden), block, touched


def extract(text, filename, scala):
    pins = []
    block = False
    for number, line in enumerate(text.splitlines(), 1):
        # Build declarations use one literal per line. Refuse unhandled active forms.
        where = f'{filename}:{number}'
        code, comment, hidden, block, touched = split_comments(line, block, where)
        if touched:
            if 'mvn"' in code or 'mvn"' in hidden:
                raise ValueError(f'{where}: dependency in block comment context is unsupported')
            continue
        matches = list(PIN.finditer(code))
        if code.count('mvn"') != len(matches):
            raise ValueError(f'{where}: unsupported dependency expression')
        reason, until = parse_reason(comment, where)
        for m in matches:
            group, cross, artifact, version = m.groups()
            suffix = {'': '', ':': '', '::': '_' + scala.split('.')[0], ':::': '_' + scala}[cross]
            pins.append(Pin(group, artifact + suffix, version, where, reason, until))
    if block:
        raise ValueError(f'{filename}: unclosed block comment')
    return pins


def extract_jmh(text, filename="xl-benchmarks/package.mill"):
    """Recognize this benchmark module's literal setting, not arbitrary Scala."""
    active = []
    block = False
    for number, line in enumerate(text.splitlines(), 1):
        code, comment, _, block, touched = split_comments(line, block, f'{filename}:{number}')
        code = code.strip()
        if touched:
            if code:
                raise ValueError(f'{filename}:{number}: mixed block comment/code unsupported')
            continue
        if code:
            active.append((number, code, comment))
    if block:
        raise ValueError(f'{filename}: unclosed block comment')
    declarations = [(i, row) for i, row in enumerate(active) if 'jmhCoreVersion' in row[1]]
    if len(declarations) != 1 or any('jmhGeneratorByteCodeVersion' in row[1] for row in active):
        raise ValueError(f'{filename}: require one JMH core literal and inherited generator version')
    i, (number, code, comment) = declarations[0]
    header = 'object `package` extends build.XLInternalModule with contrib.jmh.JmhModule {'
    match = re.fullmatch(r'def jmhCoreVersion = "([0-9][A-Za-z0-9_.+-]*)"', code)
    if not match or i == 0 or active[i - 1][1] != header:
        raise ValueError(f'{filename}:{number}: unsupported benchmark JMH declaration')
    # A following Scala selection or infix line can continue a literal RHS.
    # Recognize the existing next declaration (or the end of an empty module).
    if i + 1 >= len(active) or active[i + 1][1] not in {'override def moduleDeps = Seq(', '}'}:
        raise ValueError(f'{filename}:{number}: unsupported continuation after JMH literal')
    reason, until = parse_reason(comment, f'{filename}:{number}')
    return [Pin('org.openjdk.jmh', artifact, match[1], f'{filename}:{number}', reason, until)
            for artifact in ('jmh-core', 'jmh-generator-bytecode')]


MILL_REASON = re.compile(r'^// mill-version (pinned.*)$', re.M)


def mill_reason(build):
    """(reason, review date) for keeping the .mill-version pin past 365 days.

    .mill-version holds only a bare version and the build header line is checked
    verbatim, so the reason lives on its own `// mill-version pinned until
    YYYY-MM-DD: ...` line in build.mill. Without one, an old Mill fails like any
    other unexplained pin.
    """
    reasons = MILL_REASON.findall(build)
    if len(reasons) > 1:
        raise ValueError('build.mill: more than one mill-version pinned reason')
    return parse_reason(reasons[0], 'build.mill') if reasons else ('', None)


def check_full_cross(texts):
    """Renovate derives `:::` suffixes from build.mill's Scala version with one
    combination match, which supports exactly one full-cross dependency there."""
    for name, text in texts.items():
        count = len(re.findall(r'mvn"[^"]*:::', text))
        if name != 'build.mill' and count:
            raise ValueError(f'{name}: full-cross (:::) dependencies belong only in build.mill')
        if count > 1:
            raise ValueError('build.mill: Renovate supports one full-cross (:::) dependency')


def inventory(root):
    build = (root / 'build.mill').read_text()
    match = re.search(r'^val scala3Version = "([0-9.]+)"(?: //(.*))?$', build, re.M)
    if not match:
        raise ValueError('Unsupported Scala version declaration')
    scala = match[1]
    files = subprocess.check_output(['git', 'ls-files', '-z', '*.mill'], cwd=root).decode().split('\0')
    texts = {f: (root / f).read_text() for f in files if f}
    check_full_cross(texts)
    pins = [p for f, text in texts.items() for p in extract(text, f, scala)]
    if not pins:
        raise ValueError('No direct dependency declarations found')
    # Include the direct toolchain compiler and the build's Mill-versioned JMH plugin.
    pins.append(Pin('org.scala-lang', 'scala3-compiler_3', scala, 'build.mill:scala3Version',
                    *parse_reason(match[2] or '', 'build.mill:scala3Version')))
    headers = [line for line in build.splitlines() if line.startswith('//| mvnDeps:')]
    if headers != ['//| mvnDeps: ["com.lihaoyi::mill-contrib-jmh:$MILL_VERSION"]']:
        raise ValueError('Review changed build-header dependencies')
    pins.append(Pin('com.lihaoyi', 'mill-contrib-jmh_3', (root / '.mill-version').read_text().strip(),
                    'build.mill:1', *mill_reason(build)))
    pins.extend(extract_jmh((root / 'xl-benchmarks/package.mill').read_text()))
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


def evaluate(pin, modified_at, now):
    if modified_at > now:
        raise ValueError('Future POM Last-Modified: ' + pin.url)
    today = now.date()
    if pin.until and pin.until > today + timedelta(days=366):
        raise ValueError(f'{pin.location}: pin review date {pin.until} is more than a year ahead')
    age = now - modified_at
    justified = bool(pin.reason) and pin.until is not None and today <= pin.until
    return {'location': pin.location, 'coordinate': f'{pin.group}:{pin.artifact}:{pin.version}',
            'pom': pin.url, 'lastModified': modified_at.isoformat(), 'ageDays': age.days,
            'reason': pin.reason, 'reviewBy': pin.until.isoformat() if pin.until else None,
            'passed': age <= timedelta(days=365) or justified}


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
