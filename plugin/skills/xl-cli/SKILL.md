---
name: xl-cli
description: "LLM-friendly Excel operations via the `xl` CLI. Read cells, view ranges, search, evaluate formulas, export (CSV/JSON/PNG/PDF), style cells, modify rows/columns. Use when working with .xlsx files or spreadsheet data."
---

# XL CLI - Excel Operations

**Requires xl >= 0.21.1.** Check with `xl --version`. Older binaries lack `--json`, `xl schema`,
`xl batch --schema`, `describe`, `audit`, `deps`, the 0/1/2/3 exit table and globals-anywhere; every
statement in this skill assumes 0.21.1 or later.

The binary documents itself and is the reference: `xl <verb> --help` for a verb's flags,
`xl schema` for every verb, `xl batch --schema` for every batch op and field, `xl functions --json`
for every formula function. This skill is the map; those are the territory.

## Installation

Check if installed: `which xl || echo "not installed"`; then `xl --version` (must print `0.21.1`
or later).

**Release download** — the latest published native binary (no JDK required):

**macOS/Linux (recommended):**
```bash
# Auto-detect platform and install latest release
REPO="TJC-LP/xl"
LATEST=$(curl -s "https://api.github.com/repos/$REPO/releases/latest" | grep '"tag_name"' | cut -d'"' -f4)
VERSION=${LATEST#v}
case "$(uname -s)-$(uname -m)" in
  Linux-x86_64)  BINARY="xl-$VERSION-linux-amd64" ;;
  Linux-aarch64) BINARY="xl-$VERSION-linux-arm64" ;;
  Darwin-x86_64) BINARY="xl-$VERSION-darwin-amd64" ;;
  Darwin-arm64)  BINARY="xl-$VERSION-darwin-arm64" ;;
  *) echo "Unsupported: $(uname -s)-$(uname -m)" && exit 1 ;;
esac
mkdir -p ~/.local/bin
curl -fsSL "https://github.com/$REPO/releases/download/$LATEST/$BINARY" -o ~/.local/bin/xl || {
  echo "Error: no $BINARY published for $LATEST — use the JAR distribution xl-cli-$VERSION.tar.gz instead" >&2
  exit 1
}
chmod +x ~/.local/bin/xl
echo "Installed xl $VERSION to ~/.local/bin/xl"
xl --version
```

**Alternative using GitHub CLI:**
```bash
# If gh is installed (simpler, handles auth for private repos)
gh release download --repo TJC-LP/xl --pattern "xl-*-$(uname -s | tr A-Z a-z)-$(uname -m | sed 's/x86_64/amd64/;s/aarch64/arm64/')" -D /tmp
mv /tmp/xl-* ~/.local/bin/xl && chmod +x ~/.local/bin/xl
```

**Windows (PowerShell):**
```powershell
$repo = "TJC-LP/xl"
$latest = (Invoke-RestMethod "https://api.github.com/repos/$repo/releases/latest").tag_name
$version = $latest -replace '^v', ''
$url = "https://github.com/$repo/releases/download/$latest/xl-$version-windows-amd64.exe"
Invoke-WebRequest -Uri $url -OutFile "$env:LOCALAPPDATA\xl.exe"
Write-Host "Installed xl $version"
```

Ensure `~/.local/bin` is in your PATH: `export PATH="$HOME/.local/bin:$PATH"`

---

## Mental model (read once)

```
usage: xl [-f FILE] [-s SHEET] [-o OUT | -i] [--json] <verb> …
```

1. **Global flags go anywhere** on the command line, before or after the verb: `-f`/`--file`,
   `-s`/`--sheet`, `-o`/`--output`, `-i`/`--in-place`, `--stream`, `--max-size`, `--backend`,
   `--no-recalc` (`--preserve-caches`), `--strict`, `--json`. `xl view A1:B2 -f f -s Data` is
   `xl -f f -s Data view A1:B2`. The one exception: `--strict` right after `view` is view's own
   `--eval` gate, not the write gate. After `--` every token is data —
   `xl -f a.xlsx search -- --json` searches for the text `--json` (without the `--` the flag is
   hoisted and `search` has no pattern: exit 2).
2. **ONE sheet rule**, for every verb, batch op and `--stream` path: a sheet-qualified ref
   (`'Q1 Report'!A1:D9`) names the sheet; otherwise `-s`; otherwise the only sheet of a
   single-sheet book (a `SHEET_AUTOSELECTED` warning under `--json`); otherwise `SHEET_REQUIRED`,
   exit 3, with the sheet names as candidates. `search` (without `-s`), `sheets`, `names`, `diff`,
   `lint`, `describe` and `audit` read the whole book instead (a `-s` given to `describe`, `names`
   or `sheets` must still name a real sheet: `SHEET_NOT_FOUND`, exit 3). Start with
   `xl -f file describe`.
3. **Reads need `-f`; writes need `-o` (a new file) or `-i` (in place).** A write with neither is
   `OUTPUT_REQUIRED`, exit 2, before anything is read. Writes are atomic: the output appears only
   when the whole command succeeded. **The file is never positional**: the first non-flag token
   is the verb, so `xl data.xlsx view A1:B4` takes `data.xlsx` for a verb and fails
   `UNKNOWN_VERB` (exit 2) — a verb's positionals are its own (the range, the ref, the formula,
   `copy`'s source and target, `delete-rows`' row and count).
4. **Always pass `--json` when a program reads the result.** Every verb, success or failure, then
   prints exactly one envelope on stdout — `{ok, exitCode, verb, version, data, warnings, error}`
   — and nothing else there. `ok` is `true` exactly when `error` is `null`. On a failure (exit 2
   or 3) `data` is `null`; on findings and gates (exit 1: `diff` differs, `lint` findings,
   `audit --fail-on-findings`, `--strict`) `ok` is `false` but `data` keeps the report. Either way
   stderr carries one `Error: <message>` line. For `view`, `filter`, `diff` and `lint`, `--json`
   alone selects the verb's JSON payload as `data` (no `--format json` needed; an explicit text
   `--format` rides inside as `data.text`); prose verbs yield `data.text` plus
   `data.saved`/`data.written`.
5. **Exit codes** (branch on these and on `error.code`, never on message text):

   | exit | meaning | file written? |
   |---|---|---|
   | `0` | ok | as requested |
   | `1` | completed with findings or a failed gate (`diff` differs, `lint` findings, `audit --fail-on-findings`, `--strict`) — **never a failure** | `-o`: yes; `-i`: no |
   | `2` | usage — the command line is wrong (unknown verb, `-o` missing, `-i` with `-o`, unsupported under `--stream`) | no |
   | `3` | failed — the operation could not complete (sheet not found, bad ref, formula error, unreadable file) | no |

6. **`xl <verb> --help` is the documentation** for a verb's arguments and flags; `xl schema`
   lists every verb with what it needs, how it can exit and its batch twin.
7. **`batch` is the primary write surface**: one JSON array of ops, applied in order, atomically,
   with the recalculation once at the end. `xl batch --schema` prints the JSON Schema of the
   document (every op, field, alias, example); `xl batch --dry-run ops.json` validates without a
   workbook. Prefer one `batch` over a chain of single verbs.

```bash
xl -f model.xlsx --json describe                   # what is in this file?
xl -f model.xlsx --json audit --fail-on-findings   # anything already broken? (exit 1 if so)
xl -f model.xlsx -s Data --json view A1:D20 | jq '.data.rows'   # --json alone selects the JSON payload
xl -f model.xlsx -s Data -o out.xlsx --json batch ops.json | jq -e '.ok' >/dev/null || echo "batch failed"
```

---

## Task → verb

| Task | Verb | Notes |
|------|------|-------|
| Orient in an unknown workbook | `describe` (`--full` for counts) | sheets with state and dimension, defined names, date system; metadata-only, works under `--stream` |
| "This number looks wrong" | `audit` | error values, uncached/unparseable formulas, cycles (notes, not findings, when iterative calculation is on), unresolved names; `--fail-on-findings` exits 1 for CI |
| Where a cell's value comes from / what reads it | `deps <ref>` | `--direction precedents\|dependents\|both`, `--depth n\|all` |
| One cell: value, style, comment, direct deps | `cell <ref>` | |
| Read a block | `view <range>` | `--format markdown\|json\|csv\|html\|svg\|png\|jpeg\|webp\|pdf`, `--eval`, `--formulas`, `--limit`, `--show-labels` |
| Find text or a number | `search <regex>` | all sheets unless `-s`; `--limit` stops the scan (`total` is then a lower bound, `totalExact: false`); `--total` for the exact count |
| Rows matching a predicate | `filter --where "B > 100 AND D = TRUE"` | `--header` uses row 1 names; `--columns A,C:E` |
| Used range, numeric summary | `bounds`, `stats <range>` | |
| What-if without writing | `eval "=…" --with "A1=5"`, `evala` (arrays, `--at` to spill) | no `-f` for constants |
| Write values / formulas | `put`, `putf` — or a `batch` | one formula over a range drags with `$` anchoring |
| Style, merge, comments, hyperlinks | `style`, `merge`/`unmerge`, `comment`/`remove-comment` — or `batch` ops | styles merge unless `--replace` |
| Copy, fill, sort or clear a block | `copy <source> <target> [--values-only]`, `fill <source> <target> [--right]`, `sort <range> --by <col>`, `clear <range> [--all\|--styles\|--comments]` — or the batch ops `copy` and `clear` | `copy` shifts relative references like Excel; the target is a cell (expanded to the source's size) or a range, and either side may be sheet-qualified: `{"op":"copy","source":"Data!A1:B2","target":"Summary!A1","valuesOnly":false}`. `fill` and `sort` have no batch twin |
| Rows and columns | `row <n>`, `col <letter>`, `autofit [--columns A:F]`, `group-rows <10:20>`/`group-cols <E:H>` (`--level n`, `--collapsed`), `insert-rows <at-row> [count]`/`delete-rows <at-row> [count]`, `insert-cols <at-col\|C:E> [count]`/`delete-cols <at-col\|C:E> [count]` | `at-row` is one 1-based row and `count` defaults to 1: `delete-rows 7 5` deletes rows 7-11 — there is **no** `7:11` form (only the column verbs take `C:E`). Structural edits rewrite formulas on every sheet, `#REF!` on loss; they have no batch twin |
| Sheets | `add-sheet`, `remove-sheet`, `rename-sheet`, `move-sheet`, `copy-sheet`, `sheets hide\|show`, `name add\|rm` | `rename-sheet` rewrites every reference to the sheet |
| Deliverable finish | `sheet-view`, `tab-color`, `page-setup`, `header-footer`, `autofilter`, `freeze`, `cf add`, `chart add`, `add-image` | every one but `add-image` has a batch twin |
| Import data | `import <csv>`, `import-md <table.md\|->` | `--new-sheet`, type detection |
| Refresh cached values | `recalc` (`--tables`, `--parallel n`) | `--strict` exits 1 on formula errors |
| Compare, validate before sending | `diff -g other.xlsx`, `lint` | exit 1 = differences / findings |
| New workbook | `new out.xlsx --sheet Data --sheet Summary` | |
| What can the binary do? | `schema`, `functions`, `rasterizers`, `batch --schema` | no `-f` |

---

## Recipes

### Explore, then act

```bash
xl -f data.xlsx describe --full             # sheets, names, date system, per-sheet counts
xl -f data.xlsx -s Sheet1 view A1:E20       # preview (markdown; add --limit 0 for all rows)
xl -f data.xlsx -s Sheet1 stats B2:B100
xl -f data.xlsx -s Sheet1 deps C5 --depth all
xl -f data.xlsx -s Sheet1 eval "=SUM(A1:A10)" --with "A1=500,A5=0"
```

### Build a formatted report in one atomic batch

```bash
xl -f template.xlsx -o report.xlsx batch - <<'EOF'
[
  {"op": "put",   "sheet": "Data", "ref": "A1", "value": "Sales Report"},
  {"op": "style", "sheet": "Data", "range": "A1:E1", "bold": true, "bg": "navy", "fg": "white"},
  {"op": "put",   "sheet": "Data", "ref": "A2:E2", "values": ["Date", "Company", "Revenue", "Growth", "Status"]},
  {"op": "put",   "sheet": "Data", "ref": "C3", "value": 1234.5, "format": "currency"},
  {"op": "putf",  "sheet": "Data", "ref": "D3:D10", "value": "=C3/C$3-1", "from": "D3", "format": "percent"},
  {"op": "colwidth", "sheet": "Data", "col": "A", "width": 25},
  {"op": "autofit",  "sheet": "Data", "columns": "B:E"},
  {"op": "freeze",   "sheet": "Data", "ref": "A3"},
  {"op": "add-sheet", "name": "Summary", "after": "Data"},
  {"op": "putf", "sheet": "Summary", "ref": "B2", "value": "=SUM(Data!C3:C10)", "format": "currency"}
]
EOF
```

Batch essentials (the complete, generated field list is one command away: `xl batch --schema`):

- **Native JSON types**: numbers, booleans and `null` are stored as such. Strings are
  smart-detected — `"$1,234.56"` → currency, `"59.4%"` → 0.594 with a percent format,
  `"2025-11-10"` → a date; `"detect": false` keeps a string as text.
- **`format`** on `put`/`putf` is explicit: a name (`general`, `integer`, `decimal`, `currency`,
  `percent`, `date`, `datetime`, `time`, `text`) or any Excel format code (`"0.0x"`,
  `"$#,##0;($#,##0)"`), and it **replaces** the cell's number format. A detected format only
  applies to a General cell. A string that is neither a name nor code-shaped is ignored with a
  `FORMAT_HINT_IGNORED` warning that names it and lists the known names (the cell stays General).
- **`values`** writes a row-major array over a range; `putf` with a single `value` over a range
  drags it from `from` (Excel `$` anchoring); `putf` `values` writes each formula as-is.
- **`sheet`** on any op (except `add-sheet`/`rename-sheet`) names the sheet for its unqualified
  refs, so a batch can touch several sheets and never needs shell quoting for sheet names with
  spaces. A qualified ref (`"Summary!B2"`) wins over it.
- **Property names** are accepted in camelCase or kebab-case; `format`/`numFormat`,
  `from`/`anchor`, `target`/`url`, `align`/`halign` and `value`/`formula` (on `putf`) are aliases.
  An unknown property is an `UNKNOWN_PROPERTY` warning, not an error.
- **Validate first**: `xl batch --dry-run ops.json` (no workbook needed). The dry run's parse
  warnings are `Warning[CODE]:` lines on stderr, or the envelope's `warnings[]` under `--json`
  (with `location.opIndex`); `data.ops[].index` is 1-based — the index a `BATCH_OP_FAILED`
  names at apply time.

### Formula dragging and anchors

```bash
xl -f f.xlsx -s S1 -o o.xlsx putf B2:B10 "=A2*1.1"          # B2: =A2*1.1, B3: =A3*1.1, …
xl -f f.xlsx -s S1 -o o.xlsx putf C2:C10 "=SUM(\$A\$1:A2)"   # running total: C3: =SUM($A$1:A3), …
```

| Syntax | Behavior |
|--------|----------|
| `$A$1` | Absolute (never shifts) |
| `$A1`  | Column absolute, row relative |
| `A$1`  | Column relative, row absolute |
| `A1`   | Fully relative (shifts both ways) |

`put` writes text and values (`put A1 "Total Revenue"`); `putf` always parses a formula. Batch
`put`: `put A1:D1 "Q1" "Q2" "Q3" "Q4"` (row-major), `put A1:A10 "TBD"` (fill), `--csv` to split
one comma-separated value across the range, `--no-detect` to keep dates/numbers as text,
`--value "-100"` (or `put A1 -- -5`) for a negative number.

### Cross-sheet references and shell quoting

Cross-sheet references use `!`: `=Data!B5`, `=SUM('Q1 Sales'!A1:A100)`. In bash, single-quote
the formula so `!` is not history-expanded (`putf A1 '=Sheet2!B1'`); when the sheet name itself
needs single quotes, double-quote the argument (`putf B4 "='Income Statement'!G8"`) or move the
edit into a batch heredoc (`<<'EOF'`), where nothing needs escaping.

### Output formats and images

`view --format json` gives typed cells (`{ref, type, value, formatted}`; formula cells add
`formula`); `csv` with `--show-labels` keeps row numbers visible; `--format html` renders inline
CSS (fonts, fills, number formats) with no rasterizer; `svg` is pure vector, no backend either;
`png`/`jpeg`/`webp`/`pdf` need `--raster-output <path>` and a rasterizer — `xl rasterizers` lists
what is available (the native binary needs one external tool: `pip install cairosvg` or
`apt install librsvg2-bin`). Add `--eval` when formula cells should show computed values.

```bash
xl -f data.xlsx -s Sheet1 view A1:F20 --format png --raster-output /tmp/sheet.png --show-labels --eval
```

### Large files (100k+ rows)

`--stream` runs in O(1) memory for the reads `search`, `stats`, `bounds`, `view`
(markdown/csv/json; `view --limit 0` streams the whole sheet row by row — csv and json from the
first row, markdown after one pass for the column widths — so dump a big sheet with `--format csv`
or `json`), `cell`, `filter`, `describe` (the metadata card), `sheets` (the listing), `names` and
`lint` (the SAX lint, same findings), and for the writes `put`, `putf`, `style` and `batch` — the
last for streamable ops only (`xl batch --schema` marks each op `x-streamable`; `batch --help`
marks the others `[not with --stream]`). Every other write verb accepts the flag but loads the
workbook in memory and only writes through the streaming writer, so it saves no memory. Refused up
front with `UNSUPPORTED_IN_STREAM` (exit 2, before any read): `audit`, `deps`, `diff`, `eval`,
`evala`, `new`, `functions`, `rasterizers`, `schema`, `describe --full`, `sheets --stats`,
`view --eval`, `put --csv`, `--strict` on a streamed write, and a batch op the streaming writer
cannot apply or whose `sheet`/qualified ref names a sheet other than the streamed one (refused by
index before any byte is written; a streamed `style` merges as in memory, and an op that fails to
apply is `BATCH_OP_FAILED` with its index). `view --format html|svg|png|jpeg|webp|pdf` needs the
styles and is not available under `--stream`. `xl schema --json` publishes each verb's answer as
`stream`: `o1`, `backend` or `refused`. Under `--stream`, `--max-size` bounds the shared-string
table — the one part a streaming read holds in memory (default 100 MB; `SECURITY_ERROR` past it,
`0` lifts it). Under `--json` the table streams too (csv/markdown as `data.text`, json spliced into
`data`): a failure before the first row is a normal `ok: false` envelope, one after the first byte
leaves the envelope unterminated — branch on the exit code (3) and stderr, never on stdout parsing
alone. Streaming never recalculates. For everything else, load in memory
with `--max-size 0` (lifts the 100 MB security limit) or `--max-size 500`. That lifts the limit,
not the heap: the native binary's heap is capped at 8 GB unless `-Xmx<size>` is passed — put it
before `-f` to be safe (`xl -Xmx64g -f big.xlsx …`; the JAR takes `java -Xmx64g -jar`), and an in-memory load
needs 16–20× its worksheet XML for dense cells (a million rows × 8 numbers, 346 MB of XML, loads
in 6 GB and not in 5), 33–41× for wide text rows, 2–3× the `sharedStrings.xml` bytes for long
text — so stream large files. With `--max-size` lifted, the load is sized before parsing: one
that cannot fit even at the lower estimate (`14 × sheet XML + 2 × SST`) is refused as
`RESOURCE_LIMIT` (exit 3); one that may not fit (upper estimate `30 × sheet + 3 × SST` above the
heap) proceeds under a `MEMORY_PRESSURE` warning; one that then exhausts the heap — loading,
recalculating, evaluating, rendering or serialising — fails as `RESOURCE_LIMIT` with the
`--stream`/`-Xmx` hint, never a raw `OutOfMemoryError`. `--max-size` below 0 is a usage error.

```bash
xl -f huge.xlsx --stream search "pattern" --limit 10
xl -f huge.xlsx -o out.xlsx --stream putf A2 "=B2*1.1"
xl -f huge.xlsx --max-size 0 sheets
xl -Xmx32g -f huge.xlsx --max-size 0 audit          # native image: -Xmx anywhere; before -f to be safe
```

### Cache posture and strict pipelines

Writes recalculate the edit's dependency cone and report formula errors advisorily (exit 0).
`--strict` turns those reports into exit 1 (`RECALC_GATE`): with `-o` the file is still written,
with `-i` the input is left untouched. `--no-recalc` (`--preserve-caches`) applies the edit and
recalculates nothing — for books whose numbers come from another engine; structural edits then
leave the formulas they invalidated uncached rather than re-stamp stale numbers. `recalc`
refreshes every cached value (`--tables` also seeds data-table interiors).

---

## Gotchas

- **`view` and `search` clip at `--limit` (default 50).** The clip is visible — `view` markdown
  appends a `… showing N of M rows` trailer, `view` json carries `truncated`/`totalRows`, and
  csv/html/svg emit a `TRUNCATED` warning; `search` says it in its own payload
  (`count`/`total`/`totalExact`) or its trailer, never as a warning — but a 50-row result is not
  the whole range: pass `--limit 0` for everything. `search` also *stops scanning* at the limit:
  `Found at least 11 matches` / `"total": 11, "totalExact": false` means the rest of the sheet was
  never read; pass `--total` when the count itself is the answer.
- **Use `--show-labels` whenever row numbers matter** in CSV output: hidden rows shift positional
  counting. `view` renders hidden rows and marks them (`--skip-hidden` to omit).
- **`putf` for formulas only.** `putf A1 "Total Revenue"` is a parse error; use `put`.
- **`format` replaces, detection defers.** An explicit `format` on `put`/`putf` overwrites the
  cell's number format; a detected one (`"$1,234"`) leaves an existing non-General format alone.
- **Batch keys are forgiving, unknown keys are warnings.** camelCase and kebab-case both work and
  the aliases above are silent; a typo (`"bolds": true`) is an `UNKNOWN_PROPERTY` warning on
  stderr (or `warnings[]`) and the op still applies — read the warnings.
- **`rename-sheet` rewrites references** in formulas, defined names, conditional-formatting rules
  and charts, on every sheet; inside a batch a rename of the default sheet retargets the ops that
  follow. `move-sheet` changes tab order only.
- **Negative numbers** look like flags: `put A1 --value "-100"` or `put A1 -- -5`. After `--`
  every token is data, so `search -- --json` searches for the text `--json`.
- **`--strict` after `view`** is view's `--eval` gate (exit 1 on evaluation failure, nothing
  rendered); everywhere else it is the write gate.
- **PNG/PDF on the native binary needs an external rasterizer** — `xl rasterizers` tells you.
- **A file that will not read** fails with `IO_READ` (exit 3) and a message naming the construct.
  Rebuild it with openpyxl, then xl works on the rebuilt file — and report the message upstream.
- **Formula caches**: `xl lint` catches structure Excel would repair; `xl audit` catches numbers
  that are wrong or uncached; `xl recalc` fills caches for readers that never recalculate
  (pandas, `openpyxl data_only=True`, previewers).

---

## Reference

Generated from the binary and CI-gated (a stale page fails the build), so these never drift:

- [Verbs](https://github.com/TJC-LP/xl/blob/main/docs/reference/generated/cli-verbs.md) — every verb, what it needs, how it exits, its batch twin (`xl schema`)
- [Batch operations](https://github.com/TJC-LP/xl/blob/main/docs/reference/generated/batch-ops.md) — every op, field, alias and example (`xl batch --schema`)
- [Functions](https://github.com/TJC-LP/xl/blob/main/docs/reference/generated/functions.md) — every formula function with arity and argument slots (`xl functions --json`)
- [Exit codes](https://github.com/TJC-LP/xl/blob/main/docs/reference/generated/exit-codes.md) and [error/warning codes](https://github.com/TJC-LP/xl/blob/main/docs/reference/generated/error-codes.md) (`xl schema --json`)

In this skill:

- [reference/FORMULAS.md](reference/FORMULAS.md) — the function table plus hand-written semantics notes
- [reference/COLORS.md](reference/COLORS.md) — color names
- [reference/OUTPUT-FORMATS.md](reference/OUTPUT-FORMATS.md) — format specs

Full prose reference: [docs/reference/cli.md](https://github.com/TJC-LP/xl/blob/main/docs/reference/cli.md).
