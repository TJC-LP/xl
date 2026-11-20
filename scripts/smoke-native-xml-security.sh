#!/usr/bin/env bash
# GH-457 / GH-485: shared by CI and every release platform. Pass the built CLI command.
set -euo pipefail

if [ "$#" -eq 0 ]; then
  echo "Usage: $0 <native-binary> [args ...]" >&2
  exit 2
fi

cd "$(dirname "$0")/.."
# A private, relative directory also works with Python under Git Bash on Windows.
WORK=$(mktemp -d .xl-xml-smoke.XXXXXX)
trap 'rm -rf "$WORK"' EXIT
PY=$(command -v python3 || command -v python)
"$PY" - "$WORK/name-bloated.xlsx" <<'PY'
import sys
import zipfile

src = 'xl-ooxml/test/resources/fixtures/small-values.xlsx'
names = ''.join(
    f'<definedName name="n{i}">Sheet1!$A${(i % 100) + 1}</definedName>'
    for i in range(120000)
)
with zipfile.ZipFile(src) as zin, \
        zipfile.ZipFile(sys.argv[1], 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'xl/workbook.xml':
            text = data.decode('utf-8')
            assert text.count('</workbook>') == 1
            data = text.replace(
                '</workbook>',
                f'<definedNames>{names}</definedNames></workbook>'
            ).encode('utf-8')
        zout.writestr(item, data)
PY

STATUS=0
OUTPUT=$("$@" -f "$WORK/name-bloated.xlsx" sheets 2>&1) || STATUS=$?
echo "$OUTPUT"
if [ "$STATUS" -ne 0 ] || echo "$OUTPUT" | grep -q 'JAXP000'; then
  echo "::error::native binary cannot read a name-bloated workbook (GH-457 / GH-485)"
  exit 1
fi
