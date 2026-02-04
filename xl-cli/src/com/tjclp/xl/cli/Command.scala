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

/**
 * Command ADT representing all CLI operations.
 *
 * Named CliCommand to avoid conflict with com.monovore.decline.Command.
 */
enum CliCommand:
  // Read-only (workbook-level)
  case Sheets(action: SheetsAction)
  case Names
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
  // Analyze
  case Eval(formula: String, overrides: List[String])
  case EvalArray(formula: String, targetRef: Option[String], overrides: List[String])
  // Mutate (require -o)
  case Put(ref: String, values: List[String])
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
  case Batch(source: String) // "-" for stdin or file path
  case Import(
    csvPath: String,
    startRef: Option[String],
    delimiter: Char,
    skipHeader: Boolean,
    encoding: String,
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

/** Fill direction for the fill command */
enum FillDirection derives CanEqual:
  case Down // Fill downward (default)
  case Right // Fill rightward

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
