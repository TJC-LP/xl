package com.tjclp.xl.cli

import java.nio.file.Path

/**
 * Command ADT representing all CLI operations.
 *
 * Named CliCommand to avoid conflict with com.monovore.decline.Command.
 */
enum CliCommand:
  // Read-only (workbook-level)
  case Sheets
  case Names
  // Read-only (sheet-level)
  case Bounds
  case View(
    range: String,
    showFormulas: Boolean,
    evalFormulas: Boolean,
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
  case ColOp(col: String, width: Option[Double], hide: Boolean, show: Boolean)
  case Batch(source: String) // "-" for stdin or file path
  case Import(
    csvPath: String,
    startRef: Option[String],
    delimiter: Char,
    hasHeader: Boolean,
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

/** Fill direction for the fill command */
enum FillDirection derives CanEqual:
  case Down // Fill downward (default)
  case Right // Fill rightward
