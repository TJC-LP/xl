package com.tjclp.xl.cli

import java.nio.file.Path

/**
 * Sheets subcommand actions.
 *
 * Supports list, hide, and show operations on sheets.
 */
sealed trait SheetsAction derives CanEqual
object SheetsAction:
  /** List all sheets (instant by default, --stats for cell counts) */
  case class List(stats: Boolean) extends SheetsAction

  /** Hide a sheet from the sheet tabs */
  case class Hide(name: String, veryHide: Boolean) extends SheetsAction

  /** Show a hidden sheet (make it visible) */
  case class Show(name: String) extends SheetsAction

/** Named-range (defined name) operations: add/replace and remove. */
sealed trait NameAction derives CanEqual
object NameAction:
  /** Add or replace a workbook-scoped named range. */
  case class Add(name: String, refersTo: String) extends NameAction

  /** Remove a workbook-scoped named range. */
  case class Remove(name: String) extends NameAction

/**
 * Command ADT representing all CLI operations.
 *
 * Named CliCommand to avoid conflict with com.monovore.decline.Command.
 */
enum CliCommand:
  // Read-only (workbook-level)
  case Sheets(action: SheetsAction)
  case Names
  // Mutation (workbook-level)
  case Name(action: NameAction)
  // Read-only (sheet-level)
  case Bounds(scan: Boolean) // scan=false: instant (dimension element), scan=true: full scan
  case View(
    range: String,
    showFormulas: Boolean,
    evalFormulas: Boolean,
    strict: Boolean,
    limit: Int,
    format: ViewFormat,
    printScale: Boolean,
    showGridlines: Boolean,
    showLabels: Boolean,
    dpi: Int,
    quality: Int,
    rasterOutput: Option[Path],
    skipEmpty: Boolean,
    headerRow: Option[Int],
    rasterizer: Option[String]
  )
  case Cell(ref: String, noStyle: Boolean)
  case Search(pattern: String, limit: Int, sheetsFilter: Option[String])
  case Stats(ref: String)
  // Row filtering with a --where predicate (GH-134, phase 1 — read-only, in-memory)
  case Filter(
    where: String,
    columns: Option[String],
    limit: Int,
    format: FilterFormat,
    header: Boolean
  )
  // Analyze
  case Eval(formula: String, overrides: List[String])
  case EvalArray(formula: String, targetRef: Option[String], overrides: List[String])
  // Mutate (require -o)
  case Put(
    ref: String,
    values: List[String],
    csvSplit: Boolean = false,
    detect: Boolean = true
  )
  case PutFormula(ref: String, formulas: List[String])
  case Style(
    range: String,
    bold: Boolean,
    italic: Boolean,
    underline: Boolean,
    bg: Option[String],
    fg: Option[String],
    fontSize: Option[Double],
    fontName: Option[String],
    align: Option[String],
    valign: Option[String],
    wrap: Boolean,
    numFormat: Option[String],
    border: Option[String],
    borderTop: Option[String],
    borderRight: Option[String],
    borderBottom: Option[String],
    borderLeft: Option[String],
    borderColor: Option[String],
    replace: Boolean
  )
  case RowOp(row: Int, height: Option[Double], hide: Boolean, show: Boolean)
  case ColOp(col: String, width: Option[Double], hide: Boolean, show: Boolean, autoFit: Boolean)
  // Row/column outline grouping (GH-421, require -o)
  case GroupRows(rows: String, level: Int, collapsed: Boolean)
  case GroupCols(cols: String, level: Int, collapsed: Boolean)
  case UngroupRows(rows: String)
  case UngroupCols(cols: String)
  case Batch(source: String, dryRun: Boolean = false) // "-" for stdin or file path
  // Whole-workbook recalculation: cache every formula's value (GH-352).
  // `tables` additionally seeds data-table interior caches (GH-442); default stays pinned-cache.
  case Recalc(tables: Boolean)
  case Import(
    csvPath: String,
    startRef: Option[String],
    delimiter: Char,
    skipHeader: Boolean,
    encoding: String,
    newSheet: Option[String],
    noTypeInference: Boolean
  )
  case ImportMarkdown(
    mdPath: String, // "-" for stdin or file path
    startRef: Option[String],
    skipHeader: Boolean,
    newSheet: Option[String],
    noTypeInference: Boolean
  )
  // Sheet management
  case AddSheet(name: String, after: Option[String], before: Option[String])
  case RemoveSheet(name: String)
  case RenameSheet(oldName: String, newName: String)
  case MoveSheet(name: String, toIndex: Option[Int], after: Option[String], before: Option[String])
  case CopySheet(sourceName: String, targetName: String)
  // Cell operations
  case Merge(range: String)
  case Unmerge(range: String)
  case AddComment(ref: String, text: String, author: Option[String])
  case RemoveComment(ref: String)
  case Clear(range: String, all: Boolean, styles: Boolean, comments: Boolean)
  case Fill(source: String, target: String, direction: FillDirection)
  case AutoFit(columns: Option[String]) // None = all used columns, Some("A:F") = specific range
  case Sort(range: String, sortKeys: List[SortKey], hasHeader: Boolean)
  // Freeze pane operations (require -o)
  case Freeze(ref: String) // Freeze panes at ref (rows above + columns left)
  case Unfreeze // Remove freeze panes
  // Sheet appearance & print setup (GH-358, require -o)
  case SheetViewOp(gridlines: Option[Boolean], zoom: Option[Int], tabSelected: Option[Boolean])
  case TabColorOp(color: Option[String], clear: Boolean)
  // Sheet-level autoFilter authoring (GH-432, requires -o)
  case AutoFilterOp(range: Option[String], clear: Boolean)
  case PageSetupOp(
    orientation: Option[String],
    scale: Option[Int],
    fitToWidth: Option[Int],
    fitToHeight: Option[Int],
    fitToPage: Option[Boolean] // tri-state: None derives/preserves (GH-284)
  )
  case HeaderFooterOp(
    oddHeader: Option[String],
    oddFooter: Option[String],
    evenHeader: Option[String],
    evenFooter: Option[String],
    firstHeader: Option[String],
    firstFooter: Option[String],
    differentOddEven: Boolean,
    differentFirst: Boolean
  )
  // Conditional formatting (GH-324): add requires -o; list is read-only
  case CfAdd(
    range: String,
    rule: String,
    bold: Boolean,
    italic: Boolean,
    underline: Boolean,
    strike: Boolean,
    bg: Option[String],
    fg: Option[String]
  )
  case CfList
  // Drawing layer (GH-221/GH-222, require -o)
  case ChartAdd(
    chartType: String, // column | bar | line | pie
    grouping: Option[String], // clustered | stacked | percent-stacked (column/bar only)
    data: String, // values range, e.g. B2:D10 (qualified accepted)
    categories: Option[String], // categories vector, e.g. A2:A10
    seriesNames: Option[String], // comma-separated literal names, positional
    seriesColors: Option[String], // comma-separated colors (hex/named), positional (GH-407)
    title: Option[String],
    legend: Option[String], // right | left | top | bottom | top-right | none
    at: String // placement: range -> two-cell anchor, single cell -> default extent
  )
  case AddImage(
    imagePath: Path,
    at: String, // single cell (natural size or --size) or range (stretch over)
    size: Option[String] // WxH in pixels, e.g. 320x240
  )
  // Range operations (require -o)
  case Copy(source: String, target: String, valuesOnly: Boolean)
  // Structural editing: insert/delete rows & columns with formula rewriting (require -o)
  case InsertRows(at: Int, count: Int) // Insert `count` rows before 1-based row `at`
  case DeleteRows(at: Int, count: Int) // Delete `count` rows starting at 1-based row `at`
  case InsertColumns(col: String, count: Int) // Insert `count` columns before column `col`
  case DeleteColumns(col: String, count: Int) // Delete `count` columns starting at column `col`
  // Compare two workbooks (-f vs -g); exit code 0 = identical, 1 = differs, 2 = error
  case Diff(file2: Path, format: DiffFormat)
  // Validate package structure on the raw zip (GH-397); exit 0 = clean, 1 = findings, 2 = error
  case Lint(format: LintFormat)

/** Fill direction for the fill command */
enum FillDirection derives CanEqual:
  case Down // Fill downward (default)
  case Right // Fill rightward

/** Output format for the diff command */
enum DiffFormat derives CanEqual:
  case Markdown // Human-readable, grouped by sheet (default)
  case Json // Stable machine-readable schema

/** Output format for the lint command */
enum LintFormat derives CanEqual:
  case Text // Human-readable findings list (default)
  case Json // Stable machine-readable schema

/** Output format for the filter command */
enum FilterFormat derives CanEqual:
  case Markdown // Table with Row column (default)
  case Csv // RFC 4180 with a label header line
  case Json // Array of {row, cells} objects with typed values

/** Sort direction for the sort command */
enum SortDirection derives CanEqual:
  case Ascending
  case Descending

/** Sort mode for the sort command */
enum SortMode derives CanEqual:
  case Alphanumeric // Case-insensitive string comparison (default)
  case Numeric // Force numeric comparison

/** Sort key specifying column, direction, and mode */
final case class SortKey(column: String, direction: SortDirection, mode: SortMode) derives CanEqual
