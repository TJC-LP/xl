<!-- GENERATED from `xl schema --json` by DocsGenSpec; do not edit by hand.
     A diff here is a contract change. Regenerate: XL_UPDATE_DOCS=1 ./mill xl-cli.test.testOnly com.tjclp.xl.cli.contract.DocsGenSpec -->

# xl verbs

Every verb in `xl --help` order, with what it needs and how it can exit. Global flags go
anywhere on the command line, before or after the verb. `needs` reads: `-f` an input
workbook; `-s` ONE sheet (a sheet-qualified ref names it, else `-s`, else the only sheet
of a single-sheet book, else `SHEET_REQUIRED`); `-o`/`-i` an output (or in-place edit);
`--stream` that the verb runs in O(1) memory under `--stream` (other write verbs accept
the flag but load the workbook and only write through the streaming writer). `exit` lists
the exit codes the
verb can end with (see [exit-codes.md](exit-codes.md)); `batch twin` is the batch op that
makes the same edit; `since` is the release the verb is documented from. Run
`xl <verb> --help` for a verb's own options and `xl schema --json` for this table as JSON.

## Global flags

| flag | short | takes a value | meaning |
| --- | --- | --- | --- |
| `--json` | — | no | Wrap every result, success or failure, in the JSON envelope |
| `--file` | `-f` | yes | Excel file to operate on |
| `--sheet` | `-s` | yes | Sheet to select; a sheet-qualified ref wins over it and a single-sheet book needs neither |
| `--max-size` | — | yes | Max uncompressed size in MB for an in-memory load (default 100, 0 = unlimited); lifts the security limit only — the heap (native image: 8 GB unless -Xmx<size> is passed; put it before -f to be safe) still bounds what fits: a load estimated not to fit is refused with RESOURCE_LIMIT, one that may not fit proceeds under a MEMORY_PRESSURE warning; below 0 is a usage error |
| `--output` | `-o` | yes | Output file for a write |
| `--in-place` | `-i` | no | Edit the input file in place (instead of -o) |
| `--backend` | — | yes | XML writer backend: scalaxml (default) or saxstax (faster) |
| `--stream` | — | no | O(1)-memory streaming for large files: search, stats, bounds, view, cell, filter, describe, sheets; put, putf, style and the streamable batch ops (other write verbs accept the flag but load the workbook) |
| `--no-recalc` | — | no | Write verbs: apply the edit and recalculate nothing; structural edits leave the formulas they invalidated uncached |
| `--preserve-caches` | — | no | Alias for --no-recalc |
| `--strict` | — | no | Write verbs: exit 1 when the recalculation reports formula errors, non-convergence or data-table seed warnings (after `view` it is view's own --eval gate) |

## Verbs

| verb | needs | exit | batch twin | since | summary |
| --- | --- | --- | --- | --- | --- |
| `rasterizers` | — | 0 2 3 | — | 0.6.1 | List available SVG-to-raster backends and their status |
| `functions` | — | 0 2 3 | — | 0.4.2 | List the supported Excel functions (--json: typed rows) |
| `schema` | — | 0 2 3 | — | 0.20.0 | Print the CLI contract: verbs, globals, exit and error codes, batch ops, functions (--json) |
| `new` | — | 0 2 3 | — | 0.1.0 | Create a blank xlsx file (--sheet <name> repeatable) |
| `diff` | `-f` | 0 1 2 3 | — | 0.11.3 | Compare two workbooks (-g <file2>) and report cell, style and structure differences |
| `lint` | `-f` | 0 1 2 3 | — | 0.15.0 | Validate the raw package against the Excel-repair classes: child order, r:id resolution, content-type coverage, over-max refs, data-table integrity, <f> canon, external refs, defined names, calc chain (read-only) |
| `eval` | `-s` | 0 2 3 | — | 0.4.2 | Evaluate a formula without modifying the sheet (--with overrides; no -f for constants) |
| `evala` | `-f` `-s` | 0 2 3 | — | 0.9.0 | Evaluate an array formula and display, or spill (--at), the result grid |
| `sheets` | `-f` `--stream` | 0 2 3 | — | 0.1.0 | List sheets with visibility state and dimension (--stats loads the book for counts) |
| `sheets hide` | `-f` `-o`/`-i` | 0 2 3 | — | 0.9.2 | Hide a sheet from the sheet tabs (--very for VBA-only) |
| `sheets show` | `-f` `-o`/`-i` | 0 2 3 | — | 0.9.2 | Show a hidden sheet |
| `names` | `-f` | 0 2 3 | — | 0.2.0 | List defined names (named ranges) |
| `bounds` | `-f` `-s` `--stream` | 0 2 3 | — | 0.9.0 | Show the used range (instant from the dimension element; --scan for an accurate scan) |
| `view` | `-f` `-s` `--stream` | 0 1 2 3 | — | 0.1.0 | View a range as markdown, json, csv, html, svg, png, jpeg, webp or pdf (--eval, --formulas) |
| `cell` | `-f` `-s` `--stream` | 0 2 3 | — | 0.1.0 | Get one cell's value, style, comment and direct dependencies |
| `search` | `-f` `--stream` | 0 2 3 | — | 0.1.0 | Search cells by regex (all sheets unless -s; the scan stops at --limit, --total counts every match) |
| `stats` | `-f` `-s` `--stream` | 0 2 3 | — | 0.2.0 | Statistics for the numeric values in a range |
| `filter` | `-f` `-s` `--stream` | 0 2 3 | — | 0.11.3 | Filter rows of the used range with a --where predicate (read-only) |
| `describe` | `-f` `--stream` | 0 2 3 | — | 0.20.0 | Orient in a workbook: sheets, defined names, date system (--full adds per-sheet counts) |
| `audit` | `-f` | 0 1 2 3 | — | 0.20.0 | Find every reason a number can be wrong, in one pass (--fail-on-findings exits 1) |
| `deps` | `-f` `-s` | 0 2 3 | — | 0.20.0 | Trace one cell's precedents and dependents, hop by hop (--direction, --depth) |
| `batch` | `-f` `-s` `-o`/`-i` `--stream` | 0 1 2 3 | — | 0.1.0 | Apply multiple operations atomically from JSON (--dry-run validates; --schema prints the op schema) |
| `put` | `-f` `-s` `-o`/`-i` `--stream` | 0 1 2 3 | `put` | 0.1.0 | Write value(s) to a cell or range with smart type detection (--no-detect, --csv) |
| `putf` | `-f` `-s` `-o`/`-i` `--stream` | 0 1 2 3 | `putf` | 0.1.0 | Write formula(s) to a cell or range; one formula over a range drags with $ anchoring |
| `style` | `-f` `-s` `-o`/`-i` `--stream` | 0 2 3 | `style` | 0.2.0 | Apply formatting to cells; styles merge by default (--replace to overwrite) |
| `row` | `-f` `-s` `-o`/`-i` | 0 2 3 | `rowheight` | 0.3.0 | Set row properties: height, hide/show |
| `col` | `-f` `-s` `-o`/`-i` | 0 2 3 | `colwidth` | 0.3.0 | Set column properties: width, hide/show, auto-fit (ranges like A:F) |
| `group-rows` | `-f` `-s` `-o`/`-i` | 0 2 3 | `group-rows` | 0.18.0 | Group rows into a collapsible outline (--level, --collapsed) |
| `group-cols` | `-f` `-s` `-o`/`-i` | 0 2 3 | `group-cols` | 0.18.0 | Group columns into a collapsible outline (--level, --collapsed) |
| `ungroup-rows` | `-f` `-s` `-o`/`-i` | 0 2 3 | `ungroup-rows` | 0.18.0 | Remove outline grouping from rows |
| `ungroup-cols` | `-f` `-s` `-o`/`-i` | 0 2 3 | `ungroup-cols` | 0.18.0 | Remove outline grouping from columns |
| `autofit` | `-f` `-s` `-o`/`-i` | 0 2 3 | `autofit` | 0.9.6 | Auto-fit column widths from content (--columns A:F) |
| `recalc` | `-f` `-o`/`-i` | 0 1 2 3 | — | 0.12.6 | Recalculate every formula and rewrite the cached values (--tables, --parallel) |
| `import` | `-f` `-s` `-o`/`-i` | 0 2 3 | — | 0.7.0 | Import CSV data with type detection (--new-sheet, --delimiter, --skip-header) |
| `import-md` | `-f` `-s` `-o`/`-i` | 0 2 3 | — | 0.11.3 | Import a GFM markdown table with type detection (--start, --new-sheet) |
| `add-sheet` | `-f` `-o`/`-i` | 0 2 3 | `add-sheet` | 0.2.0 | Add a new empty sheet (--after, --before) |
| `remove-sheet` | `-f` `-o`/`-i` | 0 2 3 | — | 0.2.0 | Remove a sheet |
| `rename-sheet` | `-f` `-o`/`-i` | 0 2 3 | `rename-sheet` | 0.2.0 | Rename a sheet and rewrite every formula, defined name and chart reference to it |
| `move-sheet` | `-f` `-o`/`-i` | 0 2 3 | — | 0.2.0 | Move a sheet to a new position (--to, --after, --before) |
| `copy-sheet` | `-f` `-o`/`-i` | 0 2 3 | — | 0.2.0 | Copy a sheet to a new name |
| `merge` | `-f` `-s` `-o`/`-i` | 0 2 3 | `merge` | 0.2.0 | Merge cells in a range |
| `unmerge` | `-f` `-s` `-o`/`-i` | 0 2 3 | `unmerge` | 0.2.0 | Unmerge cells in a range |
| `comment` | `-f` `-s` `-o`/`-i` | 0 2 3 | `comment` | 0.6.0 | Add a comment to a cell (--author) |
| `remove-comment` | `-f` `-s` `-o`/`-i` | 0 2 3 | `remove-comment` | 0.9.6 | Remove a cell's comment |
| `clear` | `-f` `-s` `-o`/`-i` | 0 2 3 | `clear` | 0.6.0 | Clear cell contents, styles or comments in a range (--all, --styles, --comments) |
| `fill` | `-f` `-s` `-o`/`-i` | 0 1 2 3 | — | 0.6.0 | Fill cells with a source value or formula, Excel Ctrl+D/Ctrl+R (--right) |
| `sort` | `-f` `-s` `-o`/`-i` | 0 2 3 | — | 0.6.0 | Sort rows in a range by one or more columns (--by, --then-by, --desc, --numeric, --header) |
| `freeze` | `-f` `-s` `-o`/`-i` | 0 2 3 | `freeze` | 0.10.0 | Freeze panes at a cell (rows above and columns left are locked) |
| `unfreeze` | `-f` `-s` `-o`/`-i` | 0 2 3 | `unfreeze` | 0.10.0 | Remove freeze panes |
| `copy` | `-f` `-s` `-o`/`-i` | 0 1 2 3 | `copy` | 0.10.0 | Copy a range to another location with formula adjustment (--values-only) |
| `name add` | `-f` `-o`/`-i` | 0 2 3 | — | 0.10.0 | Add or replace a workbook-scoped named range |
| `name rm` | `-f` `-o`/`-i` | 0 2 3 | — | 0.10.0 | Remove a named range |
| `insert-rows` | `-f` `-s` `-o`/`-i` | 0 1 2 3 | — | 0.10.0 | Insert rows; shifts cells and rewrites formulas |
| `delete-rows` | `-f` `-s` `-o`/`-i` | 0 1 2 3 | — | 0.10.0 | Delete rows; shifts cells and rewrites formulas (#REF! on loss) |
| `insert-cols` | `-f` `-s` `-o`/`-i` | 0 1 2 3 | — | 0.10.0 | Insert columns; shifts cells and rewrites formulas |
| `delete-cols` | `-f` `-s` `-o`/`-i` | 0 1 2 3 | — | 0.10.0 | Delete columns; shifts cells and rewrites formulas (#REF! on loss) |
| `chart add` | `-f` `-s` `-o`/`-i` | 0 2 3 | `chart` | 0.12.0 | Add a typed chart built from sheet data ranges (--type, --data, --at) |
| `add-image` | `-f` `-s` `-o`/`-i` | 0 2 3 | — | 0.12.0 | Embed an image (png/jpeg/gif/bmp/tiff/emf/wmf) at a cell or over a range |
| `sheet-view` | `-f` `-s` `-o`/`-i` | 0 2 3 | `sheet-view` | 0.13.0 | Set sheet view options: gridlines, zoom, tab selection |
| `tab-color` | `-f` `-s` `-o`/`-i` | 0 2 3 | `tab-color` | 0.13.0 | Set or clear the sheet tab color (named, #hex, rgb(), theme:accent1[:tint]) |
| `autofilter` | `-f` `-s` `-o`/`-i` | 0 2 3 | `autofilter` | 0.18.0 | Set or clear the sheet-level autoFilter range |
| `page-setup` | `-f` `-s` `-o`/`-i` | 0 2 3 | `page-setup` | 0.13.0 | Set print page setup: orientation, scale, fit-to-page |
| `header-footer` | `-f` `-s` `-o`/`-i` | 0 2 3 | `header-footer` | 0.13.0 | Set print header/footer text with Excel codes (&L/&C/&R, &P, &N, &D, &F, &A) |
| `cf add` | `-f` `-s` `-o`/`-i` | 0 2 3 | `cf` | 0.13.0 | Add a conditional-formatting rule (--range, --rule DSL, format flags) |
| `cf list` | `-f` `-s` | 0 2 3 | — | 0.13.0 | List conditional-formatting rules on the sheet (read-only) |
