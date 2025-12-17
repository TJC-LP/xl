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
    useImageMagick: Boolean
  )
  case Cell(ref: String, noStyle: Boolean)
  case Search(pattern: String, limit: Int, sheetsFilter: Option[String])
  case Stats(ref: String)
  // Analyze
  case Eval(formula: String, overrides: List[String])
  // Mutate (require -o)
  case Put(ref: String, value: String)
  case PutFormula(ref: String, formula: String)
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
  // Sheet management
  case AddSheet(name: String, after: Option[String], before: Option[String])
  case RemoveSheet(name: String)
  case RenameSheet(oldName: String, newName: String)
  case MoveSheet(name: String, toIndex: Option[Int], after: Option[String], before: Option[String])
  case CopySheet(sourceName: String, targetName: String)
  // Cell operations
  case Merge(range: String)
  case Unmerge(range: String)
