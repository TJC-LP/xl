# XL CLI Reference

The `xl` CLI provides a command-line interface for Excel operations, designed for LLM agents and automation.

**Design Philosophy**:
- **Stateless** — Each command is self-contained
- **Explicit cell refs** — Always use `A1`, `B5:D10` notation
- **Global flags** — `-f` for file, `-s` for sheet, `-o` for output
- **LLM-optimized output** — Markdown tables, token-efficient

---

## Installation

```bash
git clone https://github.com/TJC-LP/xl.git
cd xl
make install
```

This builds a native binary (no JDK required, instant startup) and installs `xl` to `~/.local/bin/`. Ensure it's in your PATH:

```bash
export PATH="$HOME/.local/bin:$PATH"
```

**JAR install** (requires JDK 17+): `make install-jar`

**Uninstall**: `make uninstall`

**Update**: After `git pull`, run `make install` (or `make install-jar`) again.

---

## Quick Reference

```bash
# Global flags (accepted anywhere on the command line, before or after the verb)
-f, --file <path>     # Input file (required for most commands)
-s, --sheet <name>    # Sheet to operate on (a qualified ref wins; a single-sheet book needs neither)
-o, --output <path>   # Output file for mutations
-i, --in-place        # Edit file in place (same as -o matching -f)
--stream              # O(1) memory streaming for large files (search/stats/bounds/filter, a bounded view window, writes; `view --limit 0` materialises the whole sheet — #635)
--max-size <MB>       # Max uncompressed size for in-memory load (default 100, 0 = unlimited; the heap still bounds what fits — see below)
--backend <name>      # XML backend: scalaxml (default) or saxstax (faster)
--no-recalc           # Write verbs: apply the edit, recalculate nothing (alias --preserve-caches)
--preserve-caches     # Same flag, spelled for the intent
--strict              # Write verbs: exit 1 when the write's recalculation reports problems

# Exit codes: 0 ok · 1 findings/gate (diff, lint, --strict; never a failure) · 2 usage · 3 failed
# Results on stdout; errors (Error: <message> + code:/hint: lines) and warnings on stderr

# Read-only operations
xl -f model.xlsx sheets                    # List all sheets
xl -f model.xlsx names                     # List defined names (named ranges)
xl -f model.xlsx -s "P&L" bounds           # Show used range
xl -f model.xlsx -s "P&L" view A1:D20      # View range as markdown
xl -f model.xlsx cell B5                   # Get single cell details (single-sheet books auto-select for every verb)
xl -f model.xlsx search "Revenue"          # Find cells by content (all sheets)
xl -f model.xlsx -s "P&L" stats B1:B100    # Numeric statistics for a range
xl -f model.xlsx -s "P&L" eval "=SUM(B1:B10)"   # Evaluate formula (what-if)
xl -f model.xlsx -s "P&L" evala "=TRANSPOSE(A1:C2)"  # Evaluate array formula (spill grid)

# Mutations (require -o or -i)
xl -f model.xlsx -s S1 -o output.xlsx put B5 1000000       # Write value
xl -f model.xlsx -s S1 -o output.xlsx putf C5 "=B5*1.1"    # Write formula

# What-if analysis with overrides
xl -f model.xlsx -s S1 eval "=B1*1.1" --with "B1=100"      # Evaluate with temporary values

# No file needed
xl new report.xlsx --sheet Data --sheet Summary   # Create a blank workbook
xl functions                                       # List the supported functions (--json: typed rows)
xl rasterizers                                     # List available PNG/PDF backends
xl schema                                          # Every verb: what it needs, how it exits, its batch twin
xl schema --json                                   # The whole contract as one JSON document (see below)
xl batch --schema                                  # JSON Schema of the batch document (every op, field, alias)
```

> **The binary documents itself.** `xl <verb> --help` is the reference for a verb's own flags;
> `xl schema` lists every verb; `xl schema --json` publishes the contract — exit codes, error and
> warning codes, global flags, verbs, the batch JSON Schema, the function registry and the `--json`
> envelope's schema — and the pages under [`generated/`](generated/cli-verbs.md) are rendered from
> it and CI-gated, so they cannot drift from the code.

> **Global flags go anywhere.** `xl -f x.xlsx --strict recalc` and `xl -f x.xlsx recalc --strict`
> are the same command line; so are `xl -f f -s Data view A1:B2` and `xl view A1:B2 -f f -s Data`.
> The one exception is `view --strict`: after `view` it is view's own `--eval` gate, not the write
> gate. After `--` every token is data and nothing is hoisted: `xl -f a.xlsx search -- --json`
> searches for the text `--json` (without the `--` the flag is hoisted and `search` is left with no
> pattern: exit 2). An unknown verb exits 2 with `code: UNKNOWN_VERB` and a `did you mean:` line;
> any other wrong command line exits 2 with a one-line usage, the parser's error and
> `run \`xl <verb> --help\``.

> **The file is always `-f <path>`, never positional.** The first non-flag token on the command
> line is the verb, so `xl data.xlsx view A1:B4` takes `data.xlsx` for a verb and fails
> `UNKNOWN_VERB` (exit 2). A verb's positional arguments are its own — the range, the ref, the
> formula, `copy`'s source and target, `delete-rows`' row and count.

> **`--max-size` lifts the security limit, not the heap.** `--max-size 0` ("unlimited") only stops
> the reader refusing large uncompressed content; what fits is bounded by the process heap. The
> native binary is built with an 8 GB ceiling (`-R:MaxHeapSize=8g`), raised by passing `-Xmx<size>`
> on the command line — the native runtime consumes it anywhere; put it before `-f` to be safe
> (`xl -Xmx64g -f big.xlsx --max-size 0 audit`; the JAR takes the JVM's own `java -Xmx64g -jar
> xl.jar …`). `--max-size` below 0 is a usage error (exit 2), not "unlimited". An in-memory load
> needs many times its uncompressed XML.
> Measured with the reader on 1,000,000-row books (smallest heap that loads): dense numeric or
> short-text cells need 16–20× their worksheet XML (346 MB of sheet XML loads in 6 GB, not in 5);
> the wide text rows of the 0.21.0 dogfood book needed 33–41× (1.09 GB of sheet XML, 36–45 GB of
> heap); long text costs only 2–3× the bytes of `sharedStrings.xml` (380 MB of mostly strings loads
> in 1.75 GB). Large files belong to `--stream`. When `--max-size` is lifted, the load is sized from
> the zip's central directory before anything is parsed, in two bands: a *lower* estimate of
> `14 × worksheet XML + 2 × shared-string XML` already above the heap is hopeless and is refused
> (`code: RESOURCE_LIMIT`, exit 3, the file in `error.location`); an *upper* estimate of
> `30 × worksheet + 3 × strings` above the heap proceeds under a `Warning[MEMORY_PRESSURE]`
> carrying both figures and the hint; below both the load is silent (`diff` sizes its two books
> together). A load that does exhaust the heap is reported the same way — `code: RESOURCE_LIMIT`,
> exit 3, with the `--stream`/`-Xmx` hint — never a raw `java.lang.OutOfMemoryError` with exit 1
> and, under `--json`, no envelope. The catch covers every stage that builds a large structure:
> the load, the recalculation after an edit and `recalc` itself, `view --eval`'s evaluation, the
> html/svg/raster renders and the serialisation. Two residuals: an `OutOfMemoryError` raised first
> on another thread (a `recalc --parallel` fiber) can still end the process the old way, and a
> non-heap `OutOfMemoryError` (Metaspace, "unable to create native thread") is reported with the
> heap wording.

> **ONE sheet rule**, for every verb, batch op and `--stream` path: a sheet-qualified ref
> (`'Q1 Report'!A1:D9`) names the sheet; otherwise `-s`/`--sheet` (for a batch op, its `sheet` key
> comes first); otherwise the only sheet of a single-sheet book (under `--json` a
> `SHEET_AUTOSELECTED` warning says so); otherwise `SHEET_REQUIRED`, exit 3, with the sheet names as
> candidates. `search` (without `-s`), `sheets`, `names`, `diff`, `lint`, `describe` and `audit` read
> the whole book instead (a `-s` given to `describe`, `names` or `sheets` is still validated: a name
> that is not a sheet is `SHEET_NOT_FOUND`, exit 3). Streaming reads on a multi-sheet book without a
> sheet are `SHEET_REQUIRED` too (they no longer default to the first sheet), and qualified refs
> work under `--stream` writes.

### Cache posture on writes (`--no-recalc` / `--preserve-caches`)

Every write verb that changes cell content ends with a recalculation scoped to the edit's **dirty
dependency cone** — the changed cells plus their transitive dependents (cross-sheet included) plus
the always-dirty `INDIRECT`/`OFFSET` cells. A formula the parser rejects (an omitted argument, an
unsupported function, a name whose definition it cannot read) has no graph edges, so its TEXT
decides: it joins the cone only when a cell or range it spells — resolved through sheet qualifiers,
3-D spans and defined names — contains a dirty cell; a text the scanner cannot bound (a structured
or external reference, an unknown name, `INDIRECT` inside it) is dirty on every edit. Inside the
cone such a formula is evaluated, fails, and is left uncached and reported; outside it, the cache
the file already carried stays untouched (#606). Formula writes are evaluated against the completed
edit; an authored cycle or host failure is reported and left uncached with its affected dependents.
Unaffected caches keep the values supplied by their original calculator.

`--no-recalc` (alias `--preserve-caches`) skips the post-edit recalculation. The edit still lands;
newly written formulas can carry caches produced by the verb's authoring step. Use it when an external calculator owns the numbers. Honored by `put`,
`putf`, `fill`, `copy`, `batch`, `insert-rows`, `insert-cols`, `delete-rows`, `delete-cols`;
`recalc` rejects it as a contradiction.

How completely caches survive depends on the verb:

- **Non-structural** (`put`, `putf`, `fill`, `copy`, `batch`): explicitly written cells take their
  new content; existing formula caches elsewhere are preserved, including dependents.
- **Structural** (`insert-rows`, `insert-cols`, `delete-rows`, `delete-cols`): a structural edit
  moves cells, rewrites formula text and rewrites defined names, so xl invalidates the cache of
  every formula the edit could have changed and writes those cells **without a `<v>`**. It never
  re-asserts a pre-edit cache: whether a formula the edit *did* reach still has its old answer
  cannot be decided without recalculating, which is precisely what the flag refuses. (A formula
  reached through a shrunk defined name, or a static dependent of an `INDIRECT` cell, keeps its text
  byte-identical while its answer changes underneath it.)

  A cache survives only when the edit provably could not have changed it: the formula's text is
  byte-identical after the rewrite, it did not relocate, and nothing it transitively reads was moved
  or removed. So an edit *below* or *beside* your data preserves everything, while an edit *inside*
  a block preserves the rows above the cut and drops the rest. The summary counts both halves:

  ```
  Recalculation skipped (--no-recalc): 4 cached value(s) preserved, 7 formula(s) invalidated by
  the edit left uncached (recalculate externally)
  ```

  One consequence worth knowing: **volatile formulas** (`TODAY()`, `NOW()`, `RAND()`) above the cut
  keep their cached values too. The flag means "do not recalculate", and a volatile is no exception
  — Excel refreshes them on open regardless.

  Reopening the file in Excel, or a later `xl recalc`, fills those back in. A missing `<v>` is a
  visible gap that any recalculation repairs; a wrong one is silent and permanent, which is why xl
  never **re-asserts** a cache the edit invalidated. Unresolved named readers, aliases, and their
  dependents are invalidated on both default and `--no-recalc` paths. Name resolution respects
  sheet-local shadowing. An unsupported reference rewrite that could change a defined name's
  meaning is refused before the edit; top-level unions remain supported for rewriting even though
  the evaluator cannot compute them. **If you need every formula cached after a structural edit,
  do not pass `--no-recalc`.**

```bash
xl -f external-model.xlsx -s Data -o out.xlsx --no-recalc put B5 1000
xl -f external-model.xlsx -s Data -o out.xlsx --preserve-caches batch ops.json
```

### Strict exit codes (`--strict`)

By default a write reports formula-evaluation errors, iterative non-convergence and data-table seed
warnings in its summary and still exits 0. `--strict` promotes those to **exit 1** while printing
the same summary — for CI and scripted pipelines. Excel error *values* (`#DIV/0!`, `#N/A`) are data
conditions, not failures, and never gate.

```bash
xl -f model.xlsx -o out.xlsx --strict recalc    # exit 1 if any formula failed to evaluate
```

`recalc --parallel N` evaluates statically independent formula waves on up to `N` workers, then
folds results in the same deterministic order as ordinary `recalc`. It is an opt-in throughput
lever, not an `N`-times-speed promise: narrow dependency chains remain sequential, and range-heavy
books can be limited by allocation and memory bandwidth. If the workbook declares iterative
calculation (`<calcPr iterate="1"/>`), xl keeps the sequential fixpoint path and reports that
`--parallel` was ignored in the command summary.

With `-o` the output file is written even on a strict failure (the gate only sets the exit code).
With `-i` the temp file is discarded and the input is left byte-identical; the summary then says
`NOT saved (--strict failure): <file> left untouched` instead of `Saved:`. `--strict` is refused
together with `--stream` (streaming writes never recalculate, so the gate could never fire). Verbs
that perform no recalculation, such as presentation-only verbs, have no calculation outcome to
gate. `put`, `putf`, `fill`, and `copy` include authored formulas and affected dependents in their
reported outcomes. Structural and batch writes recalculate the whole book but report a failure only
for a cell inside the cache-write cone or one the written file leaves uncached; a formula outside
the cone whose cache was kept is never reported as "left uncached" (#606). Use `--strict` without
`--no-recalc` when the command must validate calculation results.

For `recalc --tables`, unsupported dynamic source cones produce a named skip warning; failed
source/axis/member evaluations report unseeded counts and retain unresolved-precedent diagnostics.
These warnings also fail `--strict`.

---

## Commands

### Operation Categories

| Category | Commands | Purpose |
|----------|----------|---------|
| **Info** (no `-f`) | `functions`, `rasterizers`, `schema`, `new` | Capability listing, the contract itself, blank workbook |
| **Navigate** | `sheets`, `bounds`, `names` | Find your way around |
| **Explore** | `view`, `cell`, `search`, `stats` | Read data incrementally |
| **Analyze** | `eval`, `evala` | What-if formula evaluation (scalar + array) |
| **Inspect** | `describe`, `audit`, `deps` | Orient in one call, find every reason a number is wrong, trace one cell's precedents/dependents |
| **Mutate cells** | `put`, `putf`, `style`, `fill`, `clear`, `copy`, `sort`, `merge`, `unmerge`, `comment`, `remove-comment`, `batch`, `import` | Make changes (require `-o` or `-i`) |
| **Rows/columns** | `row`, `col`, `autofit`, `insert-rows`, `delete-rows`, `insert-cols`, `delete-cols` | Sizing, visibility, structural editing |
| **Sheets & view** | `add-sheet`, `remove-sheet`, `rename-sheet`, `move-sheet`, `copy-sheet`, `sheets hide/show`, `freeze`, `unfreeze`, `name` | Workbook structure |
| **Appearance & print** | `sheet-view`, `tab-color`, `page-setup`, `header-footer` | Deliverable finish: gridlines, zoom, tab colors, print setup, footers |
| **Conditional formatting** | `cf add`, `cf list` | Highlight rules, color scales, data bars, top-N, text matches |

The categories are a hand-written orientation; the complete, CI-gated verb list is the generated
[`generated/cli-verbs.md`](generated/cli-verbs.md).

Argument shapes that are easy to guess wrong (the file is always `-f`; these are the verb's own
positionals):

| Verb | Shape | Batch twin |
|------|-------|------------|
| `insert-rows` / `delete-rows` | `<at-row> [count]` — a 1-based row and a count (default 1): `delete-rows 7 5` deletes rows 7-11. There is **no** `7:11` form | — |
| `insert-cols` / `delete-cols` | `<at-col> [count]` — a letter and a count; the column verbs also take an inclusive `C:E`, which overrides the count | — |
| `group-rows` / `group-cols` | `<10:20>` / `<E:H>` `[--level n] [--collapsed]` (`ungroup-rows`/`ungroup-cols` take the span alone) | `group-rows`, `group-cols` |
| `copy` | `<source> <target> [--values-only]` — relative references shift like Excel; the target is a cell (expanded to the source's size) or a range | `copy` (`source`, `target`, `valuesOnly`; each side may be sheet-qualified) |
| `fill` | `<source> <target> [--right]` — Excel Ctrl+D / Ctrl+R | — |
| `sort` | `<range> --by <col> [--then-by <col>] [--desc] [--numeric] [--header]` | — |
| `clear` | `<range> [--all \| --styles \| --comments]` — contents by default | `clear` |

### Command Summary

The verb table is generated from the binary and CI-gated: **[`generated/cli-verbs.md`](generated/cli-verbs.md)**
lists every verb in `xl --help` order with what it needs (`-f`, one sheet, `-o`/`-i`, `--stream`),
the exit codes it can end with, its batch twin and the release it is documented from — the same
rows `xl schema` prints and `xl schema --json` publishes under `verbs`. Each verb's own arguments
and flags are in `xl <verb> --help` and in the details below. Companion pages, all generated:
[`generated/exit-codes.md`](generated/exit-codes.md), [`generated/error-codes.md`](generated/error-codes.md)
(every `code:` value with its exit code, plus the warning codes),
[`generated/batch-ops.md`](generated/batch-ops.md) (every batch op, field, alias and example) and
[`generated/functions.md`](generated/functions.md) (every formula function with its arity, argument
slots and flags). Regenerate them with
`XL_UPDATE_DOCS=1 ./mill xl-cli.test.testOnly com.tjclp.xl.cli.contract.DocsGenSpec`; a stale page
fails `DocsGenSpec`, the same discipline as the golden corpus.

---

## Command Details

### `xl sheets [list|hide|show]`

Sheet operations. With no subcommand, defaults to `list`.

**Subcommands**:
| Subcommand | Arguments | Description |
|------------|-----------|-------------|
| `list` | `[--stats]` | List all sheets (`--stats` adds cell/formula counts; slower) |
| `hide` | `<sheet-name> [--very]` | Hide a sheet (`--very` = very hidden, VBA-only; requires `-o`) |
| `show` | `<sheet-name>` | Show a hidden sheet (requires `-o`) |

**Output** (`list --stats`):
```markdown
| # | Name        | Range    | Cells | Formulas |
|---|-------------|----------|-------|----------|
| 1 | Assumptions | A1:F50   | 234   | 12       |
| 2 | Revenue     | A1:M100  | 892   | 156      |
| 3 | P&L         | A1:N120  | 978   | 76       |
```

---

### `xl names` / `xl name add|rm`

List and manage defined names (named ranges).

```bash
xl -f model.xlsx names                                       # List all defined names
xl -f model.xlsx -o out.xlsx name add Tax 'Sheet1!$A$1'      # Add or replace
xl -f model.xlsx -o out.xlsx name rm Tax                     # Remove
```

---

### `xl bounds [--scan]`

Show the used range (bounding box of non-empty cells) for the sheet selected with `-s`. Instant by default (reads the worksheet's dimension element); `--scan` forces a full streaming scan for accurate bounds.

**Output**:
```
Sheet: Revenue
Used range: A1:M100
Rows: 1-100 (100 total)
Columns: A-M (13 total)
Non-empty: 892 cells
```

---

### `xl view [range]`

View a rectangular range — markdown table by default, or JSON/CSV/HTML/SVG/PNG/JPEG/WebP/PDF.
Without a range, the sheet's used range; `--offset` and `--limit` page through the rows and
`--max-cols` caps the columns.

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `range` | string | No | used range | Cell range (e.g., "A1:D20"); absent, the sheet's used range. From the loaded workbook that is the bounding box of every stored cell, styled-but-empty ones included (the `<dimension>` the library's writer records); `--stream` trusts the worksheet's `<dimension>` as written — a stale one, or openpyxl's merged-extent one, can differ from the stored-cell box — and when the file has no readable `<dimension>`, or it names a single cell (Excel's `A1` on an empty sheet), uses the bounding box of the non-empty cells. An empty sheet renders `(empty sheet)`, `""` for csv, `{"sheet", "range": null, "rows": []}` for json from both sources |
| `--format` | string | No | markdown | Output format: markdown, json, csv, html, svg, png, jpeg, webp, pdf |
| `--formulas` | flag | No | false | Show formulas instead of values |
| `--eval` | flag | No | false | Evaluate formulas (compute live values) |
| `--strict` | flag | No | false | Fail on formula evaluation errors (with `--eval`) |
| `--limit` | int | No | 50 | Max rows to display (0 = no limit). Under `--stream`, `--limit 0` is the one read that is NOT constant-memory: the whole sheet is materialised before the first row is rendered ([#635](https://github.com/TJC-LP/xl/issues/635)) — page a very large sheet with `--limit`/`--offset` or a range instead. When output is clipped, a truncation marker is reported: markdown appends a "… showing X of Y rows" trailer; json adds `truncated`/`totalRows` fields (under `--stream` too, since 0.21.0); csv/svg note on stderr; html notes on stderr and appends an HTML comment; raster formats append the notice to the `Exported:` line |
| `--offset` | int | No | 0 | Rows to skip from the top of the range before `--limit` applies; the trailer then reads "… showing rows X–Y of N". An offset past the last row is a usage error |
| `--max-cols` | int | No | 0 | Max columns to display, from the left (0 = all); json adds `totalCols` when clipped, the other formats report "… showing X of Y columns" like the row notice |
| `--skip-empty` | flag | No | false | Skip empty cells (JSON) or empty rows/columns (tabular) |
| `--skip-hidden` | flag | No | false | Omit hidden rows/columns. **Default renders them** — a range you named never silently loses cells (GH-474) |
| `--show-labels` | flag | No | false | Include column letters and row numbers |
| `--header-row` | int | No | — | Use values from this row as keys in JSON output (1-based; `0` or less is a usage error) |
| `--raster-output` | path | For raster | — | Output file (required for png/jpeg/webp/pdf) |
| `--dpi` | int | No | 144 | Resolution for raster output |
| `--quality` | int | No | 90 | JPEG quality 1-100 |
| `--rasterizer` | string | No | batik | PNG/JPEG export uses Apache Batik by default (pure JVM, no external tools). Force another backend: cairosvg, rsvg-convert, resvg, imagemagick (explicit opt-in since 0.11.3). Native binaries need an external backend — see [`xl rasterizers`](#xl-rasterizers) |
| `--gridlines` | flag | No | false | Show cell gridlines in SVG output |
| `--print-scale` | flag | No | false | Apply print scaling (for PDF-like output) |

**Output**:
```markdown
|   | A           | B       | C          | D       |
|---|-------------|---------|------------|---------|
| 1 | Revenue     |         | $1,000,000 |         |
| 2 | COGS        |         | $400,000   |         |
| 4 | Gross Profit|         | =C1-C2     |         |
```

**Hidden rows and columns** (GH-474): a range you addressed explicitly renders in full — hidden
lines included — and every data format carries a marker:

| Format | Marker |
|--------|--------|
| markdown | trailing `note: range includes hidden row(s) … and column(s) …` line |
| csv | the same note on **stderr** (stdout stays machine-parseable) |
| json | top-level `"hiddenRows": [5]` / `"hiddenCols": ["C"]` fields (emitted only when the range holds hidden lines) |

`--skip-hidden` restores the visible-only view, and the same marker then names what was dropped
(`note: omitted hidden …`). `html`/`svg`/`png`/`jpeg`/`webp`/`pdf` are pictures of the sheet: they
mirror Excel's display and always omit hidden lines. Streaming (`--stream`) never read row/column
properties, so it has always rendered every addressed cell — and for the same reason it cannot
honour `--skip-hidden` or emit the hidden-line marker: passing `--skip-hidden` with `--stream`
prints `note: --skip-hidden is ignored with --stream …` on stderr and renders everything, and the
absence of `hiddenRows`/`hiddenCols` (or of the note) under `--stream` means *unknown*, not
*none*. The typed records of `cell --json` and `search --json` say so explicitly: their `hidden`
field is `true`/`false` from the loaded workbook and `null` under `--stream`. Drop `--stream` when
you need hidden lines elided or flagged.

**One projection, two sources** (since 0.21.0): `view`, `cell`, `search`, `stats` and `filter`
render the same `CellRecord`s whether the cells come from the loaded workbook or from the
streaming reader, so `--stream` changes what a verb *can* answer, never how it prints: streaming
`view --format json` is the same `{sheet, range, rows}` document as in memory (it used to be a bare
array of strings), streaming `search` reports the same total, `totalExact` and trailer as the
in-memory one (both stop scanning at `--limit`; `--total` for the exact count), and `filter`
streams. What each source can answer is the `capabilities` table of `xl schema --json`
(`values`, `styles`, `formulas`, `comments` from both; `hidden`, `merges`, `hyperlinks`, `graph`,
`eval`, `render` from the loaded workbook only). A query needing more than `--stream` has is
refused before the file is opened: `UNSUPPORTED_IN_STREAM`, exit 2, with the in-memory
alternative as the hint. A field that belongs to a missing capability is reported unknown rather
than guessed: `hidden`, `dependencies` and `dependents` are `null` (text: `(not available in
streaming mode)`) under `--stream`; `mergedInto` and `hyperlink` read `null` there whether the cell
has none or the reader cannot tell. Everything else — values, kinds, formatted text, formulas with
their caches, comments, styles — is byte-identical from either source (a property test writes
generated books — with openpyxl's package-absolute worksheet Targets and Excel's `<dimension
ref="A1"/>` on empty sheets among them — and compares every read verb through both).

The one window the two sources derive differently is the default one of `view` without a range
and of `filter`. The loaded workbook addresses its stored-cell box: every cell it holds,
styled-but-empty ones included, which is the `<dimension>` the library's own writer records.
`--stream` trusts the worksheet's `<dimension>` as written, so a stale dimension, or openpyxl's
merged-extent one, can differ from the stored-cell box; when a file has no readable `<dimension>`,
or it names a single cell (Excel writes `A1` on an empty sheet), the streaming used range is the
bounding box of the non-empty cells, so an empty sheet prints `(empty sheet)` from both sources.
`bounds` and `sheets` report the `<dimension>` as the file declares it, a scan of the non-empty
cells when it has none (`bounds --scan` and `sheets --stats` always the non-empty box,
`Sheet.usedRange`); `view` and `filter` address the stored-cell box.

Why the default: `xl search` finds a value in a hidden row and `xl cell C5` reads it, so a `view`
that silently elided the same cell read as file corruption.

---

### `xl rasterizers`

List SVG-to-raster backends with live availability on this machine (no `-f` needed). PNG/JPEG/WebP/PDF export probes backends in this order: `batik` → `cairosvg` → `rsvg-convert` → `resvg`; `imagemagick` is never probed automatically (fragile SVG delegate) and must be forced with `--rasterizer imagemagick`.

**Platform matrix**:

| Distribution | Rasterization |
|--------------|---------------|
| JAR (`java -jar`, `make install-jar`) | Works out of the box — Batik is bundled (pure JVM, needs AWT) |
| Native binary (GitHub releases, `make install`) | Batik cannot work (no AWT under native-image, by design) — one external tool is required |

**External tool installs** (any one is enough):

```bash
pip install cairosvg          # Python, most portable
apt install librsvg2-bin      # rsvg-convert (Debian/Ubuntu); brew install librsvg (macOS)
cargo install resvg           # or a prebuilt binary: github.com/linebender/resvg/releases
```

When no backend is available, raster exports fail with an error naming the probed chain and pointing back at `xl rasterizers`. `--format svg` always works (pure vector, no backend needed).

---

### `xl functions [--json]`

List every formula function the evaluator supports (no `-f` needed). Text mode prints the names in
columns with the count; `--json` prints typed rows — `{name, minArgs, maxArgs, args, returnsDate,
returnsTime, dynamicDeps, volatile, specialForm}`, `maxArgs` `null` for a variadic function — for
every registry function plus `LET`, the parser-level special form (`specialForm: true`).
`dynamicDeps` and `volatile` are the evaluator's recalculation flags: the cells a call reads are
decided at evaluation time (INDIRECT, OFFSET); the value can change between two recalculations with
no input changing (TODAY, NOW, RAND, RANDBETWEEN — what `xl audit` lists under `volatile`). The
generated page [`generated/functions.md`](generated/functions.md) is rendered from exactly these
rows.

```bash
xl functions                                  # names in columns, "Supported Excel Functions (N total)"
xl --json functions | jq '.data[] | select(.dynamicDeps) | .name'   # INDIRECT, OFFSET
xl --json functions | jq '.data[] | select(.volatile) | .name'      # NOW, RAND, RANDBETWEEN, TODAY
```

---

### `xl schema [--json]`

The CLI contract from the binary itself (no `-f` needed). Text mode prints the verb table: every
verb in `xl --help` order with `NEEDS` (`-f` an input workbook; `-s` ONE sheet; `-o|-i` an output;
`--stream` accepted), `EXIT` (the exit codes the verb can end with), `BATCH TWIN` (the batch op
that makes the same edit), `SINCE` (the release the verb is documented from) and a one-line
summary, preceded by the global flags and the exit-code table.

`--json` publishes the whole contract as one document — the envelope's `data`:

| key | contents |
|---|---|
| `version` | the `xl` version |
| `exitCodes` | `[{code, meaning}]` — the four exit codes |
| `errorCodes` | `[{code, exit}]` — every `code:` value an error can carry (CLI codes, then domain codes) with the exit code it implies |
| `warningCodes` | `["READER_WARNING", "TRUNCATED", ...]` |
| `globals` | `[{name, short, takesValue, doc}]` — the global flags (`--file`/`-f`, ...) |
| `verbs` | `[{path, summary, needs: {file, sheet, output, streaming}, exit, batchTwin, since}]` — `path` is the subcommand path (`["sheets", "hide"]`), joined by a space it is the envelope's `verb` |
| `batchOps` | the batch document's JSON Schema — what `xl batch --schema` prints alone |
| `functions` | the `xl functions --json` rows |
| `envelope` | the JSON Schema of the `--json` envelope itself |

```bash
xl schema                                     # the verb table
xl --json schema | jq -r '.data.verbs[] | select(.needs.output) | .path | join(" ")'   # every write verb
xl --json schema | jq '.data.errorCodes[] | select(.exit == 1)'                          # the gates and findings
```

The pages under [`generated/`](generated/cli-verbs.md) are rendered from this document by
`DocsGenSpec` and compared with the checked-in files on every test run.

---

### `xl cell <ref>`

Get complete information about a single cell.

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `ref` | string | Yes | Cell reference (e.g., "A1", "B5") |

**Output** (formula cell):
```
Cell: C4
Type: Formula
Formula: =C1-C2
Cached Value: 600000
Formatted: $600,000
```

**Output** (value cell):
```
Cell: A1
Type: Text
Value: Revenue
```

`Dependencies` lists the cells the formula reads — single references exactly, ranges as their
occupied cells — and `Dependents` the formulas that read the cell, by name or through a range that
contains it (an empty cell inside a summed range still names the sum). Empty cells inside a range
and ranges over a sheet the workbook does not have are not listed (since 0.20.0; before, `SUM(A:A)`
listed 1,048,576 entries). Same-sheet refs are unqualified; cross-sheet ones carry the sheet
quoted as a formula would spell it (`'On-Premise'!G9`, `Sheet2!A1`), the same rendering `deps` and
`search` use, so the text pastes into `putf`/`eval`; both lists are ordered by sheet, then row, then
column. For more than one hop, use `deps`. Under `--stream` the graph is not built and both lines
say so:
`Dependencies: (not available in streaming mode)` / `Dependents: (not available in streaming mode)`
(before 0.21.0 streaming listed the formula's reference tokens — `B1, B1:B3, B3` for
`=SUM(B1:B3)` — which was neither the precedent set nor exact; drop `--stream` for the graph).

`--json`: `{ref, sheet, kind, value, formatted, formula, hidden, mergedInto, style, comment,
hyperlink, dependencies, dependents}` — the typed cell record plus what the sheet attaches to it;
`style` is `{font, fill, numFmt, align, border}` or `null` (`--no-style`). Under `--stream`
`dependencies`, `dependents` and `hidden` are `null` (unknown), and `mergedInto`/`hyperlink` are
`null` whether absent or unknown. The comment text is the same from both sources (the author-prefix
run XL's writer adds is stripped on both paths).

---

### `xl describe [--full]`

Orient in a workbook with one call. Without `--full` it reads metadata only — instant for any file
size and identical under `--stream` — listing every sheet with its visibility state and dimension,
every defined name (hidden ones flagged) and the date system. `--full` loads the book and adds the
per-sheet counts an agent needs before reading a cell, plus the calcPr settings. A workbook verb:
`-s` does not narrow the card, but it is validated — a name that is not a sheet is
`SHEET_NOT_FOUND`, exit 3.

```bash
xl -f model.xlsx describe
xl -f model.xlsx --stream describe                    # same card, O(1) memory
xl -f model.xlsx describe --full
xl -f model.xlsx --json describe --full | jq '.data.sheets[] | {name, formulas, uncachedFormulas}'
```

**Output**:
```
Sheets (3):
  1  Sheet1  A1:A3
       cells 3, formulas 1, merged 1, comments 1, freeze B2
  2  Sheet2  A1:B1
       cells 2, formulas 2
  3  Hidden  A1:A1  hidden
       cells 1, formulas 0, hidden rows 1
Defined names (2):
  Total   Sheet1!$A$3
  Secret  Sheet1!$A$1  (hidden)
Date system: 1900
Calculation: default
```
(The indented count lines and `Calculation:` appear only with `--full`; only the facets a sheet has
are printed — merges, comments, hyperlinks, freeze, tab color, autofilter, tables, charts, pictures,
cf, dv, hidden rows/cols.)

**JSON** (`--json`): `data` = `{sheets: [{name, index, state, dimension}], definedNames: [{name,
refersTo, scope, hidden}], date1904}`; `--full` adds to each sheet `cells, formulas,
uncachedFormulas, mergedRanges, comments, hyperlinks, freeze, tabColor, autoFilter, tables, charts,
pictures, conditionalFormats, dataValidations, hiddenRows, hiddenCols` and the top-level `calcPr`
(`{iterativeCalculation, maxIterations, maxChange, calcMode, fullCalcOnLoad, calcId}` or `null`).
`--stream describe --full` is refused (`UNSUPPORTED_IN_STREAM`, exit 2): the counts need the book.

---

### `xl audit [--fail-on-findings]`

Every reason a number can be wrong, bucketed in one pass over the loaded workbook.

**Findings** (they make the book dirty): `Error values` — a cached Excel error on a formula or a
bare error cell; `Uncached formulas` — no cached value (`xl recalc` fills them); `Unparseable
formulas` — this evaluator cannot parse them, with the parser's diagnostic in context; `Cycles` —
circular references (one line per strongly connected component; a note, not a finding, when the
workbook's calcPr enables iterative calculation — `iterativeCycles` in the JSON report, so
`cycles` holds only findings); `Unresolved names` — formulas
reading a defined name the graph cannot resolve. **Notes** (reported, never findings): `Volatile`
(TODAY/NOW/RAND/RANDBETWEEN cells), `Dynamic` (INDIRECT/OFFSET readers), `External references`
(other-workbook refs, whose caches are pinned), `Calculation` (the file's calcPr, when it has one).

Text mode prints the headline, then one section per non-empty bucket, findings before notes, every
list in workbook order (sheet, row, column). `-s <sheet>` restricts the cell buckets to one sheet (a
cycle stays when any member is on it). Exit 0 regardless of findings; `--fail-on-findings` exits 1
with code `AUDIT_FINDINGS` on a dirty book and keeps the report (text mode prints it with no
`Error:` line; `--json` keeps it as `data`), so a CI lane can gate on it. Not available under
`--stream`.

```bash
xl -f model.xlsx audit
xl -f model.xlsx -s Summary audit
xl -f model.xlsx --json audit --fail-on-findings | jq '.data.errorCells, .data.cycles'
```

**Output**:
```
Audit: 6 findings
Error values (1):
  Calc!A1  #DIV/0!
Uncached formulas (2):
  Calc!B1
  Notes!A1
Unparseable formulas (1):
  Calc!H1
    UNSUPPORTED(1)
    ^
    Formula error in 'UNSUPPORTED(1)': Unknown function 'UNSUPPORTED' at position 0
Cycles (1):
  Calc!E1, Calc!F1
Unresolved names (1):
  Calc!I1
Volatile (1):
  Calc!C1
Dynamic (1):
  Calc!D1
External references (1):
  Calc!G1
Calculation: iterative (100 iterations, max change 0.001)
```
A clean book prints `Audit: clean`.

**JSON** (`--json`): `data` = `{clean, findings, errorCells: [{ref, error}], uncachedFormulas: [ref],
unparseable: [{ref, message}], volatile: [ref], dynamic: [ref], cycles: [[ref]], externalRefs: [ref],
unresolvedReaders: [ref], calcPr}` — refs as `Sheet!A1` (quoted when the name needs it).

---

### `xl deps <ref> [--direction precedents|dependents|both] [--depth n|all]`

Trace one cell hop by hop. **Precedents** are the cells the formula reads — single references
exactly, ranges as their occupied cells (a full-column reference never expands to a million rows);
**dependents** are the formulas that read the cell, by name or through a range that contains it.
Layer k holds the cells exactly k hops away that no earlier layer listed; each node carries its
depth, formula and value. `--depth` defaults to `1`; `all` (or `0`) follows the whole cone (a cycle
ends the walk once every member is seen). The ref follows the sheet rule: a qualified ref names the sheet,
else `-s`, else the only sheet of a single-sheet book (`SHEET_REQUIRED` otherwise); a range is
refused. Not available under `--stream`.

```bash
xl -f model.xlsx deps Summary!B4                                   # both directions, one hop
xl -f model.xlsx -s Data deps B4 --direction precedents --depth 3
xl -f model.xlsx --json deps Summary!B4 --direction dependents --depth all | jq '.data.dependents[].ref'
```

**Output**:
```
Cell: Sheet2!A1
Formula: =Sheet1!A1*2
Value: 10
Precedents (depth 1): 1
  1  Sheet1!A1  5
Dependents (depth 1): 1
  1  Sheet2!B1  =A1+1 -> 11
```
Each node line is `<depth>  <ref>  <value>` for a constant and `<depth>  <ref>  <formula> -> <cached
value>` for a formula (`(uncached)` when it has none); an empty side prints `(none)`.

**JSON** (`--json`): `data` = `{ref, formula, value, direction, depth, precedents, dependents}` —
`depth` is the number or `"all"`, each side is `[{ref, depth, formula, value}]` or `null` when the
direction was not requested; `formula` is `null` for a constant, `value` a formula's cached value.

---

### `xl search <pattern>`

Find cells containing text matching pattern. Searches all sheets by default (no `-s` needed).
Matches are listed in row-major order per sheet; the text matched is the cell's raw lexeme (a
number with every digit, `TRUE`/`FALSE`, an error token, an uncached formula's expression), not
its formatted display.

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `pattern` | string | Yes | — | Search pattern (supports regex) |
| `--sheets` | string | No | all | Comma-separated list of sheets to search |
| `--limit` | int | No | 50 | Max results (0 = no limit). The scan stops one match past the limit (since 0.21.0, #637): a clipped hit list reads "Found at least Y matches" with a "… showing first X matches; more exist" trailer, and the rest of the sheet is never read — a bounded read on a million-row sheet, in memory and under `--stream` alike. A hit list that fits is exact ("Found Y matches") |
| `--total` | flag | No | off | Scan every cell for the exact total: "Found Y matches" and a "… showing X of Y matches" trailer when clipped (what `--limit 0` also gives, listing everything) |

`--json`: `{pattern, sheets, count, total, totalExact, matches: [{ref, sheet, kind, value,
formatted, formula, hidden, mergedInto}]}` — `count` the matches listed, `total` the matches the
scan counted, `totalExact` whether the scan read every cell (`true` with `--total`, `--limit 0` or
a hit list that fit; `false` when it stopped at `--limit`, and `total` is then a lower bound —
`limit + 1`, since one more match was seen), each match the typed cell record (`value` an exact
JSON lexeme, `formula` an object or `null`, `hidden` a boolean or `null` under `--stream`).
Matches are the occupied cells in row-major order; a cell that carries a style but no value is not
occupied (it matches nothing, from either source).

**Output** (`-s Data`, two hits, the scan ran to the end):
```markdown
Found 2 matches in Data:

| Ref      | Value          |
|----------|----------------|
| Data!A1  | Revenue        |
| Data!A10 | Revenue Growth |
```

Clipped (`search Revenue --limit 2` with more hits below): the scan stopped one match past the
limit, so the count is a lower bound and the trailer names `--total`:
```markdown
Found at least 3 matches in Data:

| Ref      | Value          |
|----------|----------------|
| Data!A1  | Revenue        |
| Data!A10 | Revenue Growth |

… showing first 2 matches; more exist (use --limit to raise; --limit 0 = no limit; --total for the exact count)
```
Without `-s` the header counts sheets (`in 2 sheets`) and the `Ref` column carries each hit's
sheet. No `TRUNCATED` warning accompanies a clipped `search`: the clip is in the payload
(`count`/`total`/`totalExact`) or the trailer.

---

### `xl eval <formula> [--with overrides]`

Evaluate a formula without modifying the file (what-if analysis). `-f` is optional for constant formulas (`xl eval "=PI()*2"`).

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `formula` | string | Yes | Formula to evaluate |
| `--with`, `-w` | string | No | Temporary cell overrides (e.g., "B1=100,B2=200"; repeatable) |

**Examples**:
```bash
xl -f model.xlsx -s Sheet1 eval "=SUM(B1:B10)"
xl -f model.xlsx -s Sheet1 eval "=B1*1.1" --with "B1=100"
```

---

### `xl evala <formula> [--at <ref>]`

Evaluate an **array formula** and display the result grid, or spill it into the sheet. Requires `-f` (array formulas need sheet context).

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `formula` | string | Yes | Array formula to evaluate |
| `--at` | string | No | Target cell for array spill (default: display only) |
| `--with`, `-w` | string | No | Temporary cell overrides (repeatable) |

**Examples**:
```bash
xl -f data.xlsx -s Sheet1 evala "=TRANSPOSE(A1:C2)"          # Display result grid
xl -f data.xlsx -s Sheet1 evala "=SEQUENCE(5)" --at E1       # Spill starting at E1
xl -f data.xlsx -s Sheet1 evala "=A1:B2*10"                  # Array arithmetic with broadcasting
```

---

### `xl stats <range>`

Calculate statistics (count, sum, min, max, average, ...) for numeric values in a range. Supports `--stream` for large files.

```bash
xl -f data.xlsx -s Sheet1 stats B2:B10000
xl -f huge.xlsx --stream stats A1:E100000
```

`--json`: `{sheet, range, count, sum, min, max, mean}` with every number an exact lexeme (never
rounded through a `Double`); the text form keeps its two-decimal rendering.

---

### `xl filter --where <predicate> [options]`

Show rows of the used range matching a predicate. Read-only (no `-o`); phase 1 of GH-134 — predicate filtering only, no SQL-style SELECT/GROUP BY.

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `--where` | string | Yes | — | Filter predicate (grammar below) |
| `--columns` | string | No | all used | Output columns, e.g. `A,C:E`. A column outside the used range is blank (`null` in JSON); a repeated column is a `USAGE` error |
| `--limit` | int | No | 50 | Max matching rows to display; `0` shows none and reports the match count alone |
| `--format` | string | No | markdown | `markdown`, `csv`, or `json` |
| `--header` | flag | No | false | First used row holds column names (excluded from matching) |

**Predicate grammar** (keywords case-insensitive; `NOT` > `AND` > `OR`, parens allowed):

| Form | Example |
|------|---------|
| Comparison (`=` `!=` `<>` `>` `>=` `<` `<=`) | `B > 100`, `A = 'Widget'`, `C = TRUE` |
| Wildcard match (`%` only) | `A LIKE 'Widget%'` |
| Inclusive range | `B BETWEEN 10 AND 100` |
| Set membership | `A IN ('x', 'y', 'z')` |
| Blank test | `A IS EMPTY`, `A IS NOT EMPTY` |

**Semantics**:
- Column refs are letters (`A`, `B`) or, with `--header`, header names from the first used row (case-insensitive; header names win over letters on collision)
- Numbers compare numerically, strings case-insensitively, booleans against `TRUE`/`FALSE`
- Type mismatch (e.g. text cell vs number literal) means the row doesn't match — never an error
- Formula cells compare by their cached value; `IS EMPTY` is true for missing cells, empty cells, and blank text

```bash
xl -f sales.xlsx -s Q1 filter --where "Revenue > 10000 AND Region = 'EMEA'" --header
xl -f data.xlsx -s Sheet1 filter --where "A LIKE 'Widget%'" --columns A,C:E --format csv
xl -f data.xlsx -s Sheet1 filter --where "B BETWEEN 10 AND 99" --format json
```

**Output**: matching rows keep their original row numbers, present whatever `--columns` selects. Markdown adds a `Row` column and a match-count footer; CSV starts with a `row,<labels>` header line; JSON is an array of `{"row": n, "cells": {<label>: <typed value>}}` objects (labels are header names with `--header`, letters otherwise; two selected columns under one header name share the key, which keeps the first's position and the last's value).

**Streaming**: `--stream` scans the used range in O(1) memory, keeping only the matching rows
(since 0.21.0). The rows and cells are those the in-memory run renders over the same window; the
window is each source's own used range, which can differ for a file another producer wrote — see
**One projection, two sources** above. No date literals in predicates yet.

---

### `xl put <ref|range> <value...>`

Write value(s) to a cell or range.

**Modes**:
| Mode | Example | Behavior |
|------|---------|----------|
| Single | `put A1 100` | Write 100 to A1 |
| Fill | `put A1:A10 "TBD"` | Fill range with the same value |
| Batch | `put A1:C1 "X" "Y" "Z"` | One value per cell (row-major) |
| CSV split | `put A1:C1 "X,Y,Z" --csv` | Split one comma-separated value across the range |

**Type Inference**:
- Numbers and formatted numbers: `1000`, `$1,234.56`, `50%`
- ISO dates: `2024-01-15`
- Booleans: `true`, `false`
- Text: Everything else

Use `--no-detect` to preserve all positional values as text, including numbers and ISO date-like
strings.

**Negative numbers**: use the `--value` flag (a leading `-` is parsed as a flag), or put the value
after `--`, which makes every following token data:
```bash
xl -f input.xlsx -s S1 -o output.xlsx put A1 --value "-500"
xl -f input.xlsx -s S1 -o output.xlsx put A1 -- -500
```

**Example**:
```bash
xl -f input.xlsx -s S1 -o output.xlsx put B5 1000000
```

---

### `xl putf <ref|range> <formula...>`

Write formula(s) to a cell or range with Excel-style dragging.

**Modes**:
| Mode | Example | Behavior |
|------|---------|----------|
| Single | `putf C1 "=A1+B1"` | One formula, one cell |
| Drag | `putf B2:B10 "=A2*1.1"` | Single formula + range: references shift per cell (`$` anchors pin) |
| Batch | `putf D1:D3 "=A1+B1" "=A2*B2" "=A3-B3"` | One formula per cell, applied as-is (no dragging) |

**Anchor modes** (`$` controls shifting when dragging): `$A$1` absolute, `$A1` column-absolute, `A$1` row-absolute, `A1` fully relative.

**Edge and whole-axis rules** (since 0.21.0, [#612](https://github.com/TJC-LP/xl/issues/612)): a whole-column reference (`$A:$A`, `E:E`) drags only along columns and a whole-row reference (`1:1`) only along rows, exactly as Excel; a reference the drag would carry off the grid — before row 1 / column A or past row 1048576 / column XFD — is written as `#REF!` (per reference: `=A1+B3` copied up one row is `=#REF!+B2`, `SUM(A1:A2)` becomes `SUM(#REF!)`), never a clamped or non-existent address.

**Examples**:
```bash
xl -f input.xlsx -s S1 -o output.xlsx putf C5 "=B5*1.1"
xl -f input.xlsx -s S1 -o output.xlsx putf C2:C10 "=SUM(\$B\$2:B2)"   # Running total
```

(The batch JSON `putf` op additionally accepts a `"from"` field to drag from an explicit source cell.)

**Formula records (GH-430)**: legacy CSE array formulas (`{=...}`) and Data Table cells read from a file
survive all rewrites — `view --formulas` and `cell` render them braced (`{=SUM(A1:A3*10)}`,
`{=TABLE(A1,A2)}`) and JSON output carries an additive `"formulaKind": "array" | "dataTable"` field.
`putf` rejects a top-level `TABLE(` expression: `TABLE(...)` is a data-table record's derived display
text, not a real function (Excel would show `#NAME?`); data-table *authoring* is tracked in GH-419.
Writing any value or formula onto a record cell replaces the record; `copy` of a data-table cell
pastes its cached constant (Excel's paste behavior) and `copy` of an array anchor pastes a plain
shifted formula.

---

### `xl style <range> [options]`

Apply styling to cells.

Styles **merge** with existing formatting by default; `--replace` overwrites.

**Arguments**:
| Arg | Type | Description |
|-----|------|-------------|
| `range` | string | Cell/range reference |
| `--bold` | flag | Bold text |
| `--italic` | flag | Italic text |
| `--underline` | flag | Underlined text |
| `--font-size` | double | Font size in points |
| `--font-name` | string | Font family (e.g., "Arial") |
| `--bg` | string | Background color (name, #hex, or rgb(r,g,b)) |
| `--fg` | string | Text color |
| `--align` | string | Horizontal alignment: left, center, right |
| `--valign` | string | Vertical alignment: top, middle, bottom |
| `--wrap` | flag | Enable text wrapping |
| `--format` | string | Number format: general, number, currency, percent, date, text |
| `--border` | string | Border style for all sides: none, thin, medium, thick |
| `--border-top` / `--border-right` / `--border-bottom` / `--border-left` | string | Per-side border style |
| `--border-color` | string | Border color (applies to all specified borders) |
| `--replace` | flag | Replace entire style instead of merging |

**Example**:
```bash
xl -f input.xlsx -s S1 -o output.xlsx style A1:D1 --bold --bg yellow --align center
```

---

### `xl row <n> [options]`

Set row properties (height, hide/show).

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `n` | int | Yes | — | Row number (1-based) |
| `--height` | double | No | — | Row height in points |
| `--hide` | flag | No | false | Hide the row |
| `--show` | flag | No | false | Show (unhide) the row |

**Examples**:
```bash
xl -f input.xlsx -o output.xlsx row 5 --height 30
xl -f input.xlsx -o output.xlsx row 10 --hide
xl -f input.xlsx -o output.xlsx row 10 --show
```

---

### `xl col <letter> [options]`

Set column properties (width, hide/show, auto-fit).

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `letter` | string | Yes | — | Column letter or range (e.g., "A", "AA", "A:F") |
| `--width` | double | No | — | Column width in character units (~8.43 default) |
| `--auto-fit` | flag | No | false | Auto-fit width based on cell content |
| `--hide` | flag | No | false | Hide the column |
| `--show` | flag | No | false | Show (unhide) the column |

**Behavior**:
- `--auto-fit` calculates optimal width based on longest cell content in the column
- Adds 2 characters of padding to the calculated width
- Minimum width is 8.43 (Excel default)
- If both `--width` and `--auto-fit` are specified, `--auto-fit` takes precedence

**Examples**:
```bash
# Set explicit width
xl -f input.xlsx -o output.xlsx col B --width 20

# Auto-fit column width based on content
xl -f input.xlsx -o output.xlsx col A --auto-fit

# Hide/show columns
xl -f input.xlsx -o output.xlsx col C --hide
xl -f input.xlsx -o output.xlsx col C --show
```

---

### `xl clear <range>`

Clear cell contents, styles, or comments from a range.

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `range` | string | Yes | — | Cell or range reference (e.g., "A1", "A1:D10") |
| `--all` | flag | No | false | Clear contents, styles, and comments |
| `--styles` | flag | No | false | Clear styles only (reset to default) |
| `--comments` | flag | No | false | Clear comments only |

**Behavior**:
- **Default** (no flags): Clears cell contents only
- **`--all`**: Clears everything (contents, styles, comments)
- **`--styles`**: Clears formatting but keeps contents and comments
- **`--comments`**: Clears comments but keeps contents and styles
- Flags can be combined: `--styles --comments` clears both
- Merged regions overlapping the cleared range are automatically unmerged

**Examples**:
```bash
# Clear contents (default)
xl -f input.xlsx -o output.xlsx clear A1:D10

# Clear everything
xl -f input.xlsx -o output.xlsx clear A1:D10 --all

# Clear styles only (keep data)
xl -f input.xlsx -o output.xlsx clear A1:D10 --styles

# Clear comments only
xl -f input.xlsx -o output.xlsx clear B5 --comments

# Clear styles and comments, keep contents
xl -f input.xlsx -o output.xlsx clear A1:D10 --styles --comments
```

---

### `xl fill <source> <target> [--right]`

Fill cells with source value/formula (Excel Ctrl+D/Ctrl+R equivalent).

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `source` | string | Yes | — | Source cell or range (e.g., "A1", "A1:C1") |
| `target` | string | Yes | — | Target range to fill (e.g., "A1:A10", "A1:C10") |
| `--right` | flag | No | false | Fill rightward instead of downward |

**Behavior**:
- **Fill Down** (default): Source row(s) are repeated down through target range
  - Columns must match between source and target
  - Example: `fill A1 A1:A10` copies A1 to A2:A10
  - Example: `fill A1:C1 A1:C10` copies row 1 to rows 2-10
- **Fill Right** (`--right`): Source column(s) are repeated right through target range
  - Rows must match between source and target
  - Example: `fill A1 A1:E1 --right` copies A1 to B1:E1
  - Example: `fill A1:A5 A1:E5 --right` copies column A to columns B-E
- Formulas are shifted using Excel anchor rules (`$` anchors are preserved)

**Examples**:
```bash
# Fill value down a column (Ctrl+D equivalent)
xl -f input.xlsx -o output.xlsx fill A1 A1:A100

# Fill multiple columns down together
xl -f input.xlsx -o output.xlsx fill A1:E1 A1:E100

# Fill value right across a row (Ctrl+R equivalent)
xl -f input.xlsx -o output.xlsx fill A1 A1:J1 --right

# Fill multiple rows right together
xl -f input.xlsx -o output.xlsx fill A1:A5 A1:J5 --right

# Formula shifting example: =A1*2 in B1 fills to =A2*2, =A3*2, etc.
xl -f input.xlsx -o output.xlsx fill B1 B1:B10
```

---

### `xl autofit [--columns A:F]`

Auto-fit column widths based on content (defaults to all used columns).

```bash
xl -f input.xlsx -s S1 -o output.xlsx autofit
xl -f input.xlsx -s S1 -o output.xlsx autofit --columns A:F
```

---

### `xl copy <source> <target> [--values-only]`

Copy a range to another location with Excel-style formula adjustment (`$` anchors preserved). `--values-only` copies values without adjusting formulas. The target is a single cell (expanded to the source's size) or a range; either side may be sheet-qualified.

```bash
xl -f input.xlsx -s S1 -o output.xlsx copy A1:C10 E1
xl -f input.xlsx -s S1 -o output.xlsx copy A1:C10 E1 --values-only
xl -f input.xlsx -o output.xlsx copy "Data!A1:C10" "Summary!A1"       # across sheets
```

Batch twin: `{"op": "copy", "source": "A1:C10", "target": "E1", "valuesOnly": false}` (`values-only` is an accepted spelling).

---

### `xl sort <range> --by <col> [options]`

Sort rows in a range by one or more columns.

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `range` | string | Yes | Range to sort |
| `--by`, `-b` | string | Yes | Primary sort column |
| `--then-by` | string | No | Secondary sort column(s) (repeatable) |
| `--desc` | flag | No | Sort descending (default: ascending) |
| `--numeric` | flag | No | Force numeric comparison ("10" > "9") |
| `--header` | flag | No | First row is header (exclude from sort) |

**Behavior**: empty cells sort last; formulas sort by cached value; rows move together.

```bash
xl -f f.xlsx -s S1 -o o.xlsx sort A1:D100 --by B --desc --numeric
xl -f f.xlsx -s S1 -o o.xlsx sort A1:D100 --by B --then-by C --header
```

---

### `xl merge <range>` / `xl unmerge <range>`

Merge or unmerge cells in a range.

```bash
xl -f input.xlsx -s S1 -o output.xlsx merge A1:D1
xl -f input.xlsx -s S1 -o output.xlsx unmerge A1:D1
```

---

### `xl comment <ref> <text> [--author name]` / `xl remove-comment <ref>`

Add or remove a cell comment.

```bash
xl -f input.xlsx -s S1 -o output.xlsx comment A1 "Verify this figure" --author "Analyst"
xl -f input.xlsx -s S1 -o output.xlsx remove-comment A1
```

---

### `xl freeze <ref>` / `xl unfreeze`

Freeze panes at a cell reference (rows above and columns to the left are locked), or remove them.

```bash
xl -f input.xlsx -s S1 -o output.xlsx freeze B2    # Freeze row 1 + column A
xl -f input.xlsx -s S1 -o output.xlsx unfreeze
```

---

### Appearance & print setup: `sheet-view`, `tab-color`, `page-setup`, `header-footer`

The "deliverable finish" commands (GH-358). Each merges into the sheet's current settings:
unspecified options are preserved. All require `-o` (or `-i`).

```bash
# Gridlines off + 85% zoom
xl -f in.xlsx -s Model -o out.xlsx sheet-view --gridlines off --zoom 85

# Tab colors: named, #hex, rgb(r,g,b), or theme:<slot>[:<tint>]
xl -f in.xlsx -s Model -o out.xlsx tab-color "#1F4E79"
xl -f in.xlsx -s Model -o out.xlsx tab-color theme:accent2:0.25
xl -f in.xlsx -s Model -o out.xlsx tab-color --clear     # clears a modeled color only

# Landscape, fit to one page wide and tall
xl -f in.xlsx -s Model -o out.xlsx page-setup --orientation landscape \
   --fit-to-width 1 --fit-to-height 1

# Confidential footer (&L/&C/&R sections; &P page, &N total, &D date, &F file, &A sheet)
xl -f in.xlsx -s Model -o out.xlsx header-footer \
   --odd-footer "&LProprietary & Confidential&RPage &P of &N"
```

**Options**:

| Command | Options |
|---------|---------|
| `sheet-view` | `--gridlines on\|off`, `--zoom <10-400>`, `--tab-selected on\|off` |
| `tab-color` | `<color>` or `--clear` |
| `page-setup` | `--orientation portrait\|landscape`, `--scale <10-400>`, `--fit-to-width <n>`, `--fit-to-height <n>`, `--fit-to-page on\|off` |
| `header-footer` | `--odd-header/--odd-footer`, `--even-header/--even-footer`, `--first-header/--first-footer`, `--different-odd-even`, `--different-first` |

**Notes**:
- Validation is up-front with clean errors (zoom/scale 10-400, orientation values, fit counts >= 1).
- `tab-color --clear` removes the *modeled* color; a tab color already present in the source file's
  XML is preserved on write (preserve-if-None semantics) and cannot be stripped by the CLI.
- `page-setup --fit-to-page` is tri-state: omitted derives the sheetPr `fitToPage` flag from
  `--fit-to-width`/`--fit-to-height` and preserves whatever the source carries; `on` forces the
  flag; `off` actively strips a preserved flag.
- Even-page text sets `different-odd-even` automatically, first-page text sets `different-first`
  (Excel ignores the text while the corresponding flag is off).
- Each command has a batch-op twin (`sheet-view`, `tab-color`, `page-setup`, `header-footer`) —
  the full deliverable finish is one batch file:

```bash
cat > finish.json <<'EOF'
[
  {"op": "sheet-view", "gridlines": false, "zoom": 85},
  {"op": "tab-color", "color": "#1F4E79"},
  {"op": "page-setup", "orientation": "landscape", "fitToWidth": 1, "fitToHeight": 1},
  {"op": "header-footer", "oddFooter": "&LProprietary & Confidential&RPage &P of &N"}
]
EOF
xl -f in.xlsx -s Model -o out.xlsx batch finish.json
```

---

### `xl cf add --range <range> --rule <dsl>` / `xl cf list`

Author conditional formatting (GH-324). `cf add` appends one rule to a range (requires `-o`);
`cf list` shows the sheet's rules (read-only). Priorities are auto-assigned in add order
(lower priority wins in Excel) — the CLI never hand-stamps them.

**Rule DSL** (`--rule`):

| Family | Syntax | Example |
|--------|--------|---------|
| Cell value | `cellIs:<op>:<value>` | `cellIs:greaterThan:100` (ops: `lessThan`/`lt`, `lessThanOrEqual`/`lte`, `equal`/`eq`, `notEqual`/`ne`, `greaterThanOrEqual`/`gte`, `greaterThan`/`gt`) |
| Range | `between:<lo>:<hi>`, `notBetween:<lo>:<hi>` | `between:10:100` |
| Formula | `expression:<formula>` | `expression:MOD(ROW(),2)=0` (formula may contain `:`) |
| Color scale | `colorScale:<c1>:<c2>[:<c3>]` | `colorScale:red:white:green` (3-point mid at 50th percentile) |
| Data bar | `dataBar:<color>` | `dataBar:#638EC6` |
| Top/bottom N | `top10:<n>[:percent]`, `bottom10:<n>[:percent]` | `top10:5:percent` |
| Text match | `text:<op>:<s>` | `text:contains:overdue` (ops: `contains`, `notContains`, `beginsWith`, `endsWith`; `<s>` may contain `:`) |

**Format flags** (highlight rules — `cellIs`, `between`, `notBetween`, `expression`, `top10`,
`bottom10`, `text` — require at least one; `colorScale`/`dataBar` carry inline colors and reject
them): `--bold`, `--italic`, `--underline`, `--strike`, `--bg <color>`, `--fg <color>`.
Flag colors accept the full color syntax including `theme:accent1[:tint]`; color tokens *inside*
`colorScale:`/`dataBar:` rule strings accept named/`#hex`/`rgb(r,g,b)` only (the `:` separator
conflicts with theme syntax).

```bash
# Red highlight for values over 100
xl -f f.xlsx -s S1 -o o.xlsx cf add --range A1:A10 \
   --rule 'cellIs:greaterThan:100' --bold --bg '#FFC7CE' --fg '#9C0006'

# 3-point color scale, then inspect
xl -f f.xlsx -s S1 -o o.xlsx cf add --range B2:B20 --rule 'colorScale:red:white:green'
xl -f o.xlsx -s S1 cf list
```

**Batch op** `cf` mirrors the command:

```json
{"op": "cf", "range": "A1:A10", "rule": "cellIs:greaterThan:100", "bold": true, "bg": "#FFC7CE"}
```

---

### `xl chart add --type <t> --data <range> --at <ref> [options]`

Add a typed chart built from sheet data ranges. Supported types: `column`, `bar` (horizontal),
`line`, `pie`.

```bash
# Column chart: one series per data column, categories down column A
xl -f in.xlsx -s Data -o out.xlsx chart add --type column \
   --data B2:D10 --categories A2:A10 --series-names "Q1,Q2,Q3" \
   --title "Revenue" --at F2:K15

# Stacked horizontal bars, no legend, single-cell placement (default ~5.6x2.8cm size)
xl -f in.xlsx -s Data -o out.xlsx chart add --type bar --grouping stacked \
   --data B2:D10 --legend none --at F2

# Pie over one series
xl -f in.xlsx -s Data -o out.xlsx chart add --type pie \
   --data B2:B6 --categories A2:A6 --at E2:J12
```

| Flag | Description |
|------|-------------|
| `--type, -t` | `column`, `bar`, `line`, `pie` (required) |
| `--grouping` | `clustered` (default), `stacked`, `percent-stacked` — column/bar only |
| `--data` | Values range; qualified refs (`Data!B2:D10`) accepted (required) |
| `--categories` | Categories vector (one row or one column) |
| `--series-names` | Comma-separated literal names, applied positionally |
| `--series-colors` | Comma-separated colors (`#307FE2,#005670`), applied positionally; unset series cycle the theme accents. Pie: colors map per **slice** (`c:dPt`), slices past the list continue the accent cycle |
| `--title` | Chart title |
| `--legend` | `right` (default), `left`, `top`, `bottom`, `top-right`, `none` |
| `--at` | Placement: a range (chart stretches over it) or a single cell (required) |

**Series split** (deterministic): orientation follows the categories vector — column categories
(`A2:A10`) make one series per data **column**, row categories (`B1:D1`) one per data **row**,
absent categories default to per-column. A dimension mismatch between categories and data is an
error, never a guess. Pie charts require exactly one series.

---

### `xl add-image <image-file> --at <ref> [--size WxH]`

Embed an image. PNG/JPEG/GIF/BMP get natural-size sniffing for a single-cell `--at`; TIFF/EMF/WMF
need an explicit `--size`. A range `--at` stretches the image over the range.

```bash
xl -f in.xlsx -s S1 -o out.xlsx add-image logo.png --at B2               # natural size
xl -f in.xlsx -s S1 -o out.xlsx add-image logo.png --at B2 --size 320x240
xl -f in.xlsx -s S1 -o out.xlsx add-image banner.jpeg --at A1:F4         # stretch over range
```

---

### `xl import <csv-file> [start-ref] [options]`

Import CSV data with automatic type detection (numbers, booleans, ISO dates).

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `csv-file` | string | Yes | — | CSV file path |
| `start-ref` | string | No | A1 | Cell where import starts |
| `--delimiter` | char | No | `,` | Field separator |
| `--encoding` | string | No | UTF-8 | Input encoding |
| `--skip-header` | flag | No | false | Skip first row (do not import) |
| `--new-sheet` | string | No | — | Create new sheet for imported data |
| `--no-type-inference` | flag | No | false | Treat all values as text |

```bash
xl -f f.xlsx -s S1 -o o.xlsx import data.csv A1 --delimiter ";" --skip-header
xl -f f.xlsx -o o.xlsx import data.csv --new-sheet "Imported"
```

**Limitations**: entire CSV is loaded into memory (recommended <50k rows); dates must be ISO 8601 (`YYYY-MM-DD`).

---

### `xl import-md <md-file|-> [--start ref] [options]`

Import a GFM (GitHub Flavored Markdown) pipe table with smart type detection. Use `-` to read from stdin — handy for LLM agents that generate tables inline.

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `md-file` | string | Yes | — | Markdown file path, or `-` for stdin |
| `--start` | string | No | A1 | Top-left cell for the imported table |
| `--skip-header` | flag | No | false | Skip the table's header row (do not import) |
| `--new-sheet` | string | No | — | Create new sheet for imported data |
| `--no-type-inference` | flag | No | false | Treat all values as text |

```bash
xl -f f.xlsx -s S1 -o o.xlsx import-md table.md --start B2
xl -f f.xlsx -o o.xlsx import-md table.md --new-sheet "Data"
printf '| A | B |\n|---|---|\n| 1 | 2 |\n' | xl -f f.xlsx -s S1 -o o.xlsx import-md -
```

**Table format** (GFM): header row, delimiter row (`|---|---|`), body rows. The first table found in the input is imported (preamble prose is skipped); the table ends at the first blank line. Outer pipes are optional, `\|` inside a cell is a literal pipe, and cell whitespace is trimmed. Body rows are padded/truncated to the delimiter row's column count.

**Type detection** (per cell, same smart detection as batch `put`): currency `$1,234.56` → Number + Currency format, percent `45.5%` → `0.455` + Percent format, ISO dates `2025-01-15` → date-formatted cell, plain numbers and `true`/`false` → typed values, everything else → text. Opt out with `--no-type-inference`.

**Alignment**: GFM markers map to cell horizontal alignment — `:---` left, `:---:` center, `---:` right (no marker leaves alignment unset).

**Limitations**: input is read as UTF-8 and parsed in memory; one table per import (first wins).

---

### Sheet management: `add-sheet`, `remove-sheet`, `rename-sheet`, `move-sheet`, `copy-sheet`

```bash
xl -f f.xlsx -o o.xlsx add-sheet Summary --after Sheet1      # or --before <name>
xl -f f.xlsx -o o.xlsx remove-sheet Scratch
xl -f f.xlsx -o o.xlsx rename-sheet "Old Name" "New Name"
xl -f f.xlsx -o o.xlsx move-sheet Summary --to 0             # or --after/--before <name>
xl -f f.xlsx -o o.xlsx copy-sheet Template "Q2 Report"
```

`rename-sheet` rewrites every reference to the old name (formulas on every sheet, defined names,
conditional formats, data validations). A dependent that mentions the sheet but cannot be parsed
refuses the whole rename before anything is written: `FORMULA_ERROR` (exit 3) with `location`
naming the cell (`sheet` and `ref`) — or the sheet alone for a conditional format or data
validation — the parser's own diagnostic in the message and a hint to fix or replace that formula
first. An unknown sheet on any of these verbs is `SHEET_NOT_FOUND`, a name Excel would reject
`INVALID_SHEET_NAME`, and a `move-sheet` without `--to`/`--after`/`--before` is `USAGE` (exit 2).

---

### Structural editing: `insert-rows`, `delete-rows`, `insert-cols`, `delete-cols`

Insert or delete rows/columns with full formula-reference rewriting: references at or past the cut shift, straddling ranges shrink, and references to deleted cells become `#REF!`. Cross-sheet references to the edited sheet are rewritten too.

**Arguments**: position (`<at-row>` 1-based, or `<at-col>` letter) and optional `<count>` (default 1). Column commands also accept an inclusive range (`C:E`), which overrides the count.

```bash
xl -f f.xlsx -s S1 -o o.xlsx insert-rows 5 2     # Insert 2 rows at row 5
xl -f f.xlsx -s S1 -o o.xlsx delete-rows 5       # Delete row 5
xl -f f.xlsx -s S1 -o o.xlsx insert-cols B 3     # Insert 3 columns at B
xl -f f.xlsx -s S1 -o o.xlsx delete-cols C:E     # Delete columns C through E
```

Example formula rewriting: deleting row 2 turns `=A1+A3` into `=A1+A2`; `=SUM(A1:A4)` shrinks to `=SUM(A1:A3)`; a direct reference to a deleted cell becomes `#REF!`.

---

### `xl batch <file|->`

Apply multiple operations atomically from JSON input.

**Arguments**:
| Arg | Type | Required | Description |
|-----|------|----------|-------------|
| `file` | string | No (default `-`) | JSON file path or `-` for stdin |
| `--dry-run` | flag | No | Validate JSON and show a summary without writing (works without `-f`/`-o`) |
| `--schema` | flag | No | Print the document's JSON Schema (draft 2020-12) and exit — every op, field, alias and example; works without `-f`/`-o` |

**The document**: a JSON array of operations, applied in order, atomically — nothing is written
unless every op applies:

```json
[
  {"op": "put", "ref": "A1", "value": "Hello"},
  {"op": "putf", "ref": "B1", "value": "=A1*2"},
  {"op": "style", "range": "A1:B1", "bold": true},
  {"op": "merge", "range": "A1:D1"},
  {"op": "colwidth", "col": "A", "width": 15.5},
  {"op": "rowheight", "row": 1, "height": 30},
  {"op": "comment", "ref": "A1", "text": "Revenue figure", "author": "Analyst"},
  {"op": "autofit", "columns": "A:D"},
  {"op": "add-sheet", "name": "Summary", "after": "Sheet1"},
  {"op": "put", "sheet": "Summary", "ref": "A1", "value": "Total"}
]
```

**Supported operations**: the op table is generated from the registry that also parses the
document — **[`generated/batch-ops.md`](generated/batch-ops.md)** lists every op with its
fields, types, required flags, aliases, example, streamability and CLI twin, and
`xl batch --schema` prints the same as a JSON Schema. An unknown `op` is `BATCH_OP_UNKNOWN`
(exit 2) with a `did you mean`; an op that fails to apply is `BATCH_OP_FAILED` with
`location.opIndex` (1-based) and the op's `sheet` — or, when the cause names a cell (a
`rename-sheet` that cannot rewrite `Summary!I23`), that cell's `sheet` and `ref` — plus the cause's
own `hint`.

**Rules every op follows**:

- **The `sheet` key.** Every op except `add-sheet`/`rename-sheet` accepts `"sheet"`: the sheet
  for its unqualified refs. THE sheet rule applies per op — a sheet-qualified ref
  (`"Summary!A1"`) wins, then the op's `sheet`, then `-s`/`--sheet`, then the only sheet of a
  single-sheet book, else `SHEET_REQUIRED`. A `rename-sheet` of the batch's default sheet
  retargets the ops that follow it.
- **Property names** are accepted in camelCase or kebab-case (`fitToWidth` / `fit-to-width`),
  and a few properties have aliases honoured silently: `format` ↔ `numFormat`, `from` ↔ `anchor`,
  `target` ↔ `url`, `align` ↔ `halign`, `value` ↔ `formula` (on `putf`). An unknown property is an
  `UNKNOWN_PROPERTY` warning (stderr in text mode, `warnings[]` under `--json`), never an error.
- **`format` on `put`/`putf` is explicit** and REPLACES the cell's number format, whatever it was
  ([#560](https://github.com/TJC-LP/xl/issues/560)). A format *detected* from a string value
  (currency, percent, ISO date) is only a hint: it applies when the cell's format is General and
  leaves an existing format alone.
- **`rename-sheet` rewrites references**: every formula, defined name, conditional-formatting
  rule and chart series that named the old sheet now names the new one, on every sheet.
- **Under `--stream`** the streamable ops (see the table) are applied with identical semantics —
  `style` merges, `format` replaces, a `put` with both `value` and `values` is refused — and an
  op's `sheet` or a qualified ref may name the streamed worksheet; any other op, or one whose
  resolved sheet differs from the streamed worksheet, is refused **by index before any byte is
  written** (`UNSUPPORTED_IN_STREAM`, exit 2), and one that fails to apply is `BATCH_OP_FAILED`
  with its 1-based index, as in memory. Streaming never recalculates and never degrades an op.

**Native JSON Types** (recommended):

```json
// Numbers are stored as numeric values (not text)
{"op": "put", "ref": "A1", "value": 99.0}

// Booleans
{"op": "put", "ref": "A2", "value": true}

// With explicit format
{"op": "put", "ref": "A3", "value": 99.0, "format": "currency"}
{"op": "put", "ref": "A4", "value": 0.594, "format": "percent"}
```

**Format Options**:

| Format Name | Description | Example Output |
|-------------|-------------|----------------|
| `general` | Default format | `1234.5` |
| `integer` | Whole numbers | `1235` |
| `decimal` | Two decimal places | `1234.50` |
| `currency` | Currency with symbol | `$1,234.50` |
| `percent` | Percentage | `59%` |
| `percent_decimal` | Percentage with decimals | `59.4%` |
| `date` | Date format | `11/10/25` |
| `datetime` | Date and time | `11/10/25 14:30` |
| `time` | Time only | `14:30:00` |
| `text` | Text format | `1234.5` |
| *custom* | Any Excel format code | See below |

**Custom Format Codes**:

```json
// MOIC/Multiple format (3.5x)
{"op": "put", "ref": "A1", "value": 3.5, "format": "0.0x"}

// Accounting format with negatives in parentheses
{"op": "put", "ref": "A2", "value": -1234, "format": "$#,##0;($#,##0)"}

// Basis points
{"op": "put", "ref": "A3", "value": 50, "format": "0 \"bps\""}

// Custom date format
{"op": "put", "ref": "A4", "value": "2025-11-10", "format": "yyyy-mm-dd"}

// Quoted-literal / semicolon-only codes are codes too (the 1/0 toggle-flag idiom)
{"op": "put", "ref": "A5", "value": 1, "format": "\"Yes \";;\"No \""}
```

**Unrecognized format strings** (GH-475): a string that is neither a known name nor Excel
format-code-shaped is a typo far more often than a code.
- On the put/putf `format` hint it is **ignored, with a `FORMAT_HINT_IGNORED` warning** (stderr in
  text mode, `warnings[]` under `--json`) naming the string and listing the known names
  (`format: "curency"` → warning, cell stays General).
- On the `style` op's `numFormat` it is still **applied as a custom code** (Excel, not xl, is the
  authority on codes) but warns the same way.

The `--stream` batch path applies exactly the same table, including custom codes — it used to know
only six names and dropped everything else in silence.

**Smart String Detection** (enabled by default):

Strings are automatically detected and formatted:

```json
// Currency detected from $ prefix
{"op": "put", "ref": "A1", "value": "$99.00"}  // → Number(99.0), Currency

// Percent detected from % suffix
{"op": "put", "ref": "A2", "value": "59.4%"}   // → Number(0.594), Percent

// ISO date detected
{"op": "put", "ref": "A3", "value": "2025-11-10"}  // → DateTime, Date format

// Plain text (no detection pattern)
{"op": "put", "ref": "A4", "value": "Hello"}       // → Text
```

**Disable Detection**: Set `"detect": false` to treat strings as plain text:

```json
{"op": "put", "ref": "A1", "value": "$99.00", "detect": false}  // → Text "$99.00"
{"op": "put", "ref": "A2", "value": "59.4%", "detect": false}   // → Text "59.4%"
```

**Formula Dragging** (putf with range):

```json
// Single formula dragged across range (uses Excel $ anchoring)
{"op": "putf", "ref": "B2:B10", "value": "=SUM($A$1:A2)", "from": "B2"}

// Explicit formulas for each cell (no dragging)
{"op": "putf", "ref": "B2:B4", "values": ["=A2*2", "=A3*2", "=A4*2"]}
```

**Formula Number Formats** (`putf` with `format`, parity with `put`):

The `format` field accepts the same named formats and custom codes as `put` and
applies the number format to the formula cell(s) — no second `style` pass needed.
Works with all three variants (single, dragging, explicit `values`):

```json
{"op": "putf", "ref": "C1", "value": "=A1*2", "format": "#,##0.0"}
{"op": "putf", "ref": "B2:B10", "value": "=A2/A$1", "from": "B2", "format": "percent"}
{"op": "putf", "ref": "D1:D2", "values": ["=SUM(A:A)", "=SUM(B:B)"], "format": "currency"}
```

**Style Options**:

| Option | Type | Description |
|--------|------|-------------|
| `bold` | boolean | Bold text |
| `italic` | boolean | Italic text |
| `underline` | boolean | Underlined text |
| `bg` | string | Background color (hex: `#FF0000`) |
| `fg` | string | Font color (hex: `#0000FF`) |
| `fontSize` | number | Font size in points |
| `fontName` | string | Font family name |
| `align` | string | Horizontal alignment: `left`, `center`, `right`, `justify` |
| `valign` | string | Vertical alignment: `top`, `middle`, `bottom` |
| `wrap` | boolean | Enable text wrapping |
| `numFormat` | string | Number format (see Format Options above) |
| `border` | string | All borders: `none`, `thin`, `medium`, `thick` |
| `borderTop` | string | Top border style |
| `borderRight` | string | Right border style |
| `borderBottom` | string | Bottom border style |
| `borderLeft` | string | Left border style |
| `borderColor` | string | Border color (hex) |
| `replace` | boolean | Replace style instead of merge (default: false) |

**Note**: `align` is the canonical horizontal-alignment key (`halign` is accepted as its alias);
`numFormat` and `format` are aliases on `style` too. Unknown properties are ignored with an
`UNKNOWN_PROPERTY` warning. The complete, generated field list is
[`generated/batch-ops.md`](generated/batch-ops.md#style).

**Examples**:

```bash
# From file
xl -f input.xlsx -o output.xlsx batch operations.json

# From stdin (pipe)
echo '[{"op": "put", "ref": "A1", "value": 100, "format": "currency"}]' | \
  xl -f input.xlsx -s Sheet1 -o output.xlsx batch -

# Complex workflow
cat <<'EOF' | xl -f input.xlsx -s Sheet1 -o output.xlsx batch -
[
  {"op": "put", "ref": "A1", "value": "Revenue", "format": "text"},
  {"op": "put", "ref": "B1", "value": 1000000, "format": "currency"},
  {"op": "style", "range": "A1:B1", "bold": true, "bg": "#FFFF00"},
  {"op": "merge", "range": "A1:B1"},
  {"op": "colwidth", "col": "A", "width": 20}
]
EOF
```

#### Common Gotchas

| Scenario | Behavior | Workaround |
|----------|----------|------------|
| Leading zeros: `"00123"` | Smart detection converts to number `123` | Use `"detect": false` to preserve as text |
| Mixed patterns: `"50 (50%)"` | First pattern wins (treated as text) | Use explicit `"format"` field |
| `values` array length mismatch | Error raised if array length ≠ range cell count | Ensure exact match |
| Percent as decimal | `"59.4%"` stored as `0.594` | Excel displays correctly with percent format |
| Invalid custom formats | Accepted but may render incorrectly in Excel | Test format codes in Excel first |
| `--stream` mode | Supports formula dragging but not formula evaluation | Use non-streaming for `--eval` |

---

### `xl diff -g <file2> [--format markdown|json]`

Compare two workbooks and report differences. The first file comes from the global `-f`, the second from `-g/--file2`. Optional global `-s/--sheet` restricts the comparison to one sheet.

**Arguments**:
| Arg | Type | Required | Default | Description |
|-----|------|----------|---------|-------------|
| `-g, --file2` | path | Yes | — | Second file to compare against |
| `--format` | string | No | markdown | `markdown` (human) or `json` (stable schema) |

**Exit codes**: `0` identical, `1` differences found, `3` error (unreadable file, sheet filter
matching neither workbook, ...) — the error goes to stderr with a `code:` line (see
[Errors, warnings and exit codes](#errors-warnings-and-exit-codes)).

**What is compared** (per sheet, refs in A1, row-major order):
- **Changed cells** — value, formula text, and resolved style (`styleChanged` boolean). Formula cells compare by formula text; cached values are derived and ignored. Styles compare resolved formatting (style id lookup), so equal formatting under different ids is not a difference.
- **Added / removed cells** — a cell with Empty value, default style, and no hyperlink counts as absent.
- **Sheets added / removed** (by name).
- **Merged ranges, comments, hyperlinks** — separate added/removed/changed deltas per sheet.

```bash
xl -f old.xlsx diff -g new.xlsx                      # Markdown report
xl -f old.xlsx -s Sheet1 diff -g new.xlsx            # One sheet only
xl -f old.xlsx diff -g new.xlsx --format json        # Machine-readable
xl -f old.xlsx diff -g new.xlsx && echo "unchanged"  # Exit-code driven
```

**JSON schema** (stable; `sheets` lists only sheets with differences):

```json
{
  "identical": false,
  "sheetsAdded": [], "sheetsRemoved": [],
  "sheets": [{
    "name": "Sheet1",
    "added":   [{"ref": "D5", "value": "New", "formula": null}],
    "removed": [{"ref": "E5", "value": "Old", "formula": null}],
    "changed": [{"ref": "A5",
                 "before": {"value": "Revenue", "formula": null},
                 "after":  {"value": "Total Revenue", "formula": null},
                 "styleChanged": false}],
    "mergesAdded": [], "mergesRemoved": [],
    "commentsAdded": [], "commentsRemoved": [], "commentsChanged": [],
    "hyperlinksAdded": [], "hyperlinksRemoved": [], "hyperlinksChanged": []
  }]
}
```

**Limitations**: both workbooks load in memory (`--max-size` applies to each); no range-level filter yet.

---

### `xl lint [<file>] [--format text|json]`

Validate the raw package structure against the Excel-repair classes — the defects Excel
repairs loudly (repair dialog, content stripped) but every lenient reader, xl's own
`read` included, accepts silently. Lint inspects the **raw zip parts**, never the parsed
model, so nothing gets normalized before it's checked.

```bash
xl lint deliverable.xlsx                         # Positional file form
xl -f deliverable.xlsx lint                      # Flag form (equivalent)
xl -f deliverable.xlsx lint --format json        # Stable schema for pipelines
xl -f deliverable.xlsx lint && echo "safe to send"
```

**What it flags** (the complete `LintCategory` roster — a test pins this list against
`LintCategory.slug`, so it cannot drift):
- **`child-order`** — child elements of `xl/workbook.xml` (CT_Workbook) or a worksheet
  (CT_Worksheet) out of ECMA-376 schema sequence (e.g. `<externalReferences>` after
  `<extLst>`)
- **`unresolved-rel-id`** — an `r:id` (sheet, externalReference, pivotCache, drawing,
  legacyDrawing, hyperlink, tablePart, …) with no entry in the paired `.rels`
- **`wrong-rel-type`** — the `r:id` resolves, but to a relationship of the wrong type
- **`missing-part`** — the relationship target part is absent from the package
- **`missing-content-type`** — a present-and-referenced part has neither an `<Override>` nor a
  matching extension `<Default>` in `[Content_Types].xml`
- **`ref-out-of-bounds`** — a `ref`/`sqref`/`dimension` token past row 1048576 or column XFD
- **`data-table-torn`** — a `<f t="dataTable">` record whose grid no longer holds together
  (missing record cell, inconsistent inputs, interior no longer matching the record)
- **`data-table-unseeded`** — an uncached data-table interior in a `calcMode="autoNoTable"`
  book: Excel does not recompute data tables on open, so the grid opens BLANK
- **`formula-leading-equals`** — `<f>` text stored with the display form's leading `=`
  (non-spec; strict readers like openpyxl misread it — re-writing the file with xl heals it)
- **`external-ref-dangling`** — a formula or defined name references external workbook `[N]`
  with no N-th `<externalReference>` entry in workbook.xml (the cross-workbook sheet-transplant
  class: ordinals index the SOURCE book's table) — Excel repairs the file by removing every
  such formula plus calcChain
- **`defined-name-invalid`** — a defined name Excel refuses: two same-scope names colliding
  under Excel's case/width/kana-insensitive comparison (`g` vs `ｇ`, `html` vs `ＨＴＭＬ`,
  `ぁ` vs `あ` — legacy fossil corpora carry these), a name past the 255-character limit, or
  whitespace / control characters — Excel repairs the file by removing the named range,
  sometimes without even logging it
- **`calc-chain-stale`** — an `xl/calcChain.xml` entry names a cell that holds no formula, or a
  sheet id `workbook.xml` does not declare — Excel repairs the file on open. The part is optional:
  drop it together with its `[Content_Types].xml` Override and `workbook.xml.rels` Relationship
  (Excel rebuilds it on save), or rebuild it. Both xl writers, in-memory and `--stream`, drop
  the source chain on every write that rewrites a worksheet, so xl output never carries one
- **`xlfn-missing`** — a post-2007 function stored bare (`IFS(`, `XLOOKUP(`, `MAXIFS(`, … where
  Excel stores `_xlfn.IFS(`), or a `LET`/`LAMBDA` whose parameters lack `_xlpm.` (openpyxl's
  `_xlfn.LET(x,1,x+1)`, which Excel reports as unreadable content on open), in a cell `<f>`, a
  conditional-formatting `<formula>`, a data-validation `<formula1>`/`<formula2>` or a
  `<definedName>`: not a repair class but a silent `#NAME?` on the first recalculation, which no
  cached value reveals (the openpyxl class of producer). The rule is the writer's own —
  `FormulaStorage.bareFutureCalls` runs the same scanner as the storage mapping, so the lint flags
  exactly the text xl's writer would still prefix. One finding per part with the bare names (a
  half-prefixed `LET` is listed as `LET`), the first five sites and the total count.
  xl's own writers emit the prefix for every slot they regenerate, but a slot a write leaves
  untouched is copied verbatim, and a CF block, DV container or name table is regenerated only
  when its parsed model no longer equals the source. To heal one, add or change a rule, validation
  or name in that part; re-entering the same formula compares equal to the source (bare and
  prefixed text parse to the same model) and is copied through bare — `xl name add` with the same
  text does not heal, `cf add` appends a rule and therefore does. A cell heals on any in-memory
  edit of its sheet; an untouched worksheet and every cell a `--stream` write does not patch are
  copied verbatim. See [LIMITATIONS.md](../LIMITATIONS.md)

**Exit codes**: `0` no findings · `1` findings reported · `3` error (unreadable file, malformed
core part) · `2` usage (no file, or a file given both ways) — errors go to stderr with a `code:`
line (see [Errors, warnings and exit codes](#errors-warnings-and-exit-codes)).

`xl lint` is read-only — it never repairs or rewrites the file. xl's own output always
lints clean; use it as a pre-send self-check in agent pipelines that splice or post-process
workbooks.

---

## Output Format

### Cell Reference Convention

Always include explicit references in output:

```markdown
# Good - References visible
|   | A        | B       |
|---|----------|---------|
| 1 | Revenue  | $1M     |
| 2 | COGS     | $400K   |
```

### Errors, warnings and exit codes

Results go to **stdout**; diagnostics go to **stderr**. On any failure stdout is empty, so a
pipeline that captures stdout never has to parse an error out of a result. An error is rendered
as today's `Error: <message>` line followed by indented, machine-stable fields:

```
Error: <message>                  human text, unchanged from earlier releases
  code: <CODE>                    stable SCREAMING_SNAKE code (see below)
  did you mean: <a>, <b>          only when there are close candidates (sheet names, ...)
  hint: <text>                    only when the error has an obvious next step
```

Example:
```
$ xl -f book.xlsx -s Dat view A1:B2
Error: Sheet not found: Dat. Available: Data, Summary
  code: SHEET_NOT_FOUND
  did you mean: Data
  hint: list sheets with `xl -f <file> sheets`
```

Warnings are `Warning[<CODE>]: <message>` lines on stderr and never change the exit code — for
instance `Warning[READER_WARNING]: MissingStylesXml` when the input has no `xl/styles.xml`.

`code` is either the domain error's code (the SCREAMING_SNAKE of the `XLError` case: `SHEET_NOT_FOUND`,
`INVALID_CELL_REF`, `FORMULA_ERROR`, `SECURITY_ERROR`, `SHEET_REQUIRED`, `IO_ERROR`, ...) or one of the
CLI-only codes: `USAGE`, `UNKNOWN_VERB`, `OUTPUT_REQUIRED`, `UNSUPPORTED_IN_STREAM`,
`BATCH_JSON_INVALID`, `BATCH_OP_UNKNOWN`, `BATCH_OP_INVALID`, `BATCH_OP_FAILED`,
`RASTERIZER_UNAVAILABLE`, `IO_READ`, `IO_WRITE`, `RESOURCE_LIMIT`, `RECALC_GATE`,
`DIFFERENCES_FOUND`, `LINT_FINDINGS`, `AUDIT_FINDINGS`, `INTERNAL`. The complete vocabulary, each code with the exit it implies, is the
generated [`generated/error-codes.md`](generated/error-codes.md) (also `xl schema --json` →
`errorCodes`/`warningCodes`). The exit code follows from the code alone:

| exit | meaning | examples | file written? |
|---|---|---|---|
| `0` | ok | | as requested |
| `1` | completed with findings or a failed gate — **never a failure** | `diff` differs, `lint` findings, `--strict` gate | `-o`: yes; `-i`: no |
| `2` | usage — the command line is wrong | unknown verb, a verb's own flag before the verb, `-o` missing, `-i` with `-o`, `--stream` with an unsupported verb or flag | no |
| `3` | failed — the operation could not complete | sheet not found, invalid ref, formula parse error, value-count mismatch, unreadable or corrupt file, security limit, a workbook that does not fit in memory (`RESOURCE_LIMIT`) | no |

Two rows worth spelling out:

- `view --eval --strict` whose evaluation fails (a circular reference, an unsupported function) is
  a **gate**, not a failure: exit `1` with `code: RECALC_GATE` on stderr and nothing rendered on
  stdout — the same code the write verbs' `--strict` uses. Without `--strict` the view renders the
  cached values and the failure is a stderr warning, exit `0`.
- A missing or unreadable input file is `code: IO_READ`, exit `3`, on every verb — `sheets`,
  `names`, `view`, `cell`, `diff`, `lint` alike.

The same table is printed by `xl --help`. (Earlier releases exited `1` for usage and failures too,
`2` for `diff`/`lint` runtime errors, and printed errors on stdout.)

### Output contract (`--json`)

**Pass the global `--json` whenever a program reads the result.** It goes anywhere on the command
line like every global flag, and every verb — success or failure — then prints exactly one JSON envelope on stdout,
with the same seven keys every time:

```json
{ "ok": true, "exitCode": 0, "verb": "view", "version": "0.20.0",
  "data": { "sheet": "Data", "range": "A1:C3", "rows": [ ... ] },
  "warnings": [ { "code": "TRUNCATED", "message": "… showing 50 of 120 rows (use --limit to raise; --limit 0 = no limit)" } ],
  "error": null }

{ "ok": false, "exitCode": 3, "verb": "put", "version": "0.20.0",
  "data": null, "warnings": [],
  "error": { "code": "SHEET_NOT_FOUND", "message": "Sheet not found: Sales. Available: Data, Summary",
             "hint": "list sheets with `xl -f <file> sheets`", "candidates": [], "location": null } }
```

| key | meaning |
|---|---|
| `ok` | `true` exactly when `error` is `null` |
| `exitCode` | the process exit code, from the table above (`0`/`1`/`2`/`3`) |
| `verb` | the subcommand path, e.g. `"view"`, `"sheets hide"`, `"cf add"` (best-effort for a usage error raised before dispatch) |
| `version` | the `xl` version that produced the envelope |
| `data` | the verb's payload (below); `null` on a failure |
| `warnings` | `[{code, message}]` — the same notices text mode prints as `Warning[CODE]:` lines (`TRUNCATED`, `HIDDEN_OMITTED`, `EVAL_FAILED` for `--eval` without `--strict`, `FLAG_IGNORED` for `--skip-hidden` under `--stream`, `READER_WARNING`); `location` when known |
| `error` | `null`, or `{code, message, hint, candidates, location}` — the fields of the stderr block; absent ones are `null` / `[]` |

**What `data` holds.** `--json` selects a verb's JSON payload where it has one and wraps the text
otherwise:

- `--format json` keeps printing the bare payload without `--json` — the shapes of `view`
  (`{sheet, range, rows}`), `filter`, `diff` and `lint` are unchanged. With `--json` and no
  explicit `--format`, that same payload is `data`: `xl --json view A1:C3` yields `data` equal to
  what `view --format json` prints bare (`--format json --json` spells the same thing out). An
  explicit text `--format` (markdown, csv, html, …) under `--json` rides inside as `data.text`.
- Typed verbs build `data` directly: `sheets` → `[{name, index, state, dimension}]` (`--stats` adds
  `cells`, `formulas`); `names` → `[{name, refersTo, scope, hidden}]`; `bounds` →
  `{sheet, range, dimension}`; `eval` → `{formula, result: {type, value, formatted}, overrides}`;
  `evala` → `{formula, spillRange, result, overrides}` with `result` in the `view` JSON shape;
  `functions` → `[{name, minArgs, maxArgs, args, returnsDate, returnsTime, dynamicDeps,
  volatile, specialForm}]`; `rasterizers` → `{backends: [{name, status, note}], anyAvailable}`;
  `batch --dry-run` → `{ops: [{index, op, summary}]}` (`index` is the op's 1-based position,
  the index a `BATCH_OP_FAILED` reports; parse warnings ride in the envelope's `warnings[]`); `batch --schema` → the batch document's JSON
  Schema; `schema` → `{version, exitCodes, errorCodes, warningCodes, globals, verbs,
  capabilities, batchOps, functions, envelope}` (see
  [`xl schema`](#xl-schema---json)); `describe` → `{sheets, definedNames, date1904}` (`--full`
  adds per-sheet counts and `calcPr`); `audit` → `{clean, findings, errorCells,
  uncachedFormulas, unparseable, volatile, dynamic, cycles, externalRefs, unresolvedReaders,
  calcPr}`; `deps` → `{ref, formula, value, direction, depth, precedents, dependents}`.
- The record-based reads are typed too (since 0.21.0): `cell` → the cell record with `style`,
  `comment`, `hyperlink`, `dependencies`, `dependents`; `search` → `{pattern, sheets, count,
  total, totalExact, matches}`; `stats` → `{sheet, range, count, sum, min, max, mean}` — every
  number an exact lexeme.
- Every other verb (`view` in a text format, and all writes) yields
  `{"text": <what text mode prints>, "saved": <path or null>, "written": <bool>}`. `saved` is the
  user-visible output path once the write was committed; it is `null` (and `written` is `false`)
  when nothing was — a read, or an `-i` run whose `--strict` gate discarded the staged file.

**Findings and gates are `ok: false` with the report kept.** `diff` with differences, `lint` with
findings, `audit --fail-on-findings` on a dirty book and a `--strict` write that fails its gate exit
`1` and carry `error.code` `DIFFERENCES_FOUND` / `LINT_FINDINGS` / `AUDIT_FINDINGS` /
`RECALC_GATE` — and `data` still holds the report (the diff, the findings, the audit buckets, the
recalculation summary), so nothing text mode showed is lost.

**Channels.** With `--json` the envelope is the only thing on stdout. When a run fails (exit `2`
or `3`), stderr also carries the single line `Error: <message>` so a human tailing a log still
sees it; the `code:` and `hint:` lines live in the envelope instead. Findings and gates (exit `1`)
print nothing on stderr, exactly as text mode does — the report is in `data`. A wrong command line (unknown verb, a
verb-owned flag before the verb, `-i` with `-o`) also produces the envelope — exit `2`,
`code: USAGE` — when `--json` is among the arguments.

```bash
xl -f book.xlsx -s Data --json view A1:C3 | jq '.data.rows'   # --json alone selects the JSON payload
xl -f book.xlsx -s Data -o out.xlsx --json put A1 42 | jq -e '.ok' >/dev/null || echo "put failed"
xl -f book.xlsx --json sheets | jq -r '.data[].name'
```

The envelope's JSON Schema is `xl-cli/resources/schema/envelope.schema.json` — shipped in the
binary and published as `xl schema --json` → `envelope` — and the golden corpus
(`xl-cli/test/resources/golden/*-json.golden`) pins one envelope per shape.

---

## See Also

- [Quick Start Guide](../QUICK-START.md) — Library usage
- [Scripting Guide](scripting.md) — When a task outgrows the CLI (loops, typed extraction, multi-file pipelines)
- [Performance Guide](performance-guide.md) — Streaming for large files
- [GitHub Issues](https://github.com/TJC-LP/xl/issues) — Feature requests and bug reports
