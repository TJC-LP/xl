package com.tjclp.xl.cli.commands

import cats.effect.IO
import cats.implicits.*
import com.tjclp.xl.{Sheet, Workbook}
import com.tjclp.xl.addressing.{ARef, Column}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.FilterFormat
import com.tjclp.xl.cli.helpers.{FilterPredicate, SheetResolver}
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Row filtering for the filter command (GH-134, phase 1).
 *
 * Scans the sheet's used range row by row and keeps rows matching the --where predicate (see
 * [[FilterPredicate]] for the grammar). With --header, the first used row provides column names for
 * predicates and output labels and is excluded from matching. Loads the workbook in memory
 * (standard --max-size envelope); --stream is not supported.
 */
object FilterCommands:

  def filter(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    where: String,
    columnsSpec: Option[String],
    limit: Int,
    format: FilterFormat,
    hasHeader: Boolean
  ): IO[String] =
    for
      sheet <- SheetResolver.requireSheet(wb, sheetOpt, "filter")
      pred <- FilterPredicate.parse(where) match
        case Right(p) => IO.pure(p)
        case Left(err) => IO.raiseError(new Exception(s"Invalid filter predicate: $err"))
      result <- sheet.usedRange match
        case None =>
          IO.pure(renderNoData(format, Vector.empty))
        case Some(range) =>
          runFilter(sheet, range, pred, columnsSpec, limit, format, hasHeader)
    yield result

  private def runFilter(
    sheet: Sheet,
    range: com.tjclp.xl.addressing.CellRange,
    pred: FilterPredicate.Pred,
    columnsSpec: Option[String],
    limit: Int,
    format: FilterFormat,
    hasHeader: Boolean
  ): IO[String] =
    val usedCols = (range.start.col.index0 to range.end.col.index0).toVector
    val headerRowIdx = range.start.row.index0

    // Header names from the first used row (original casing; matched case-insensitively)
    val headerNames: Vector[(String, Int)] =
      if hasHeader then
        usedCols.flatMap { col =>
          cellValue(sheet, col, headerRowIdx)
            .flatMap(textOf)
            .map(_.trim)
            .filter(_.nonEmpty)
            .map(_ -> col)
        }
      else Vector.empty
    val headers: Map[String, Int] = headerNames.map((name, col) => name.toLowerCase -> col).toMap

    def resolve(name: String): Option[Int] =
      headers
        .get(name.toLowerCase)
        .orElse(
          Column
            .fromLetter(name.toUpperCase)
            .toOption
            .map(_.index0)
            .filter(usedCols.contains)
        )

    // Validate every referenced column upfront for a proper error message
    val unresolved = FilterPredicate.columnRefs(pred).filter(resolve(_).isEmpty)

    if unresolved.nonEmpty then
      val available =
        if hasHeader then
          val names = headerNames.map(_._1).mkString(", ")
          s"available headers: $names; used columns: ${usedColumnsLabel(range)}"
        else s"used columns: ${usedColumnsLabel(range)}"
      IO.raiseError(
        new Exception(
          s"Unknown column(s) in predicate: ${unresolved.toVector.sorted.mkString(", ")} ($available)"
        )
      )
    else
      parseColumns(columnsSpec, usedCols) match
        case Left(err) => IO.raiseError(new Exception(err))
        case Right(selectedCols) =>
          IO.pure {
            val dataRows =
              (range.start.row.index0 to range.end.row.index0).toVector
                .filter(row => !(hasHeader && row == headerRowIdx))

            val matched = dataRows.filter { row =>
              FilterPredicate.evaluate(pred, resolve, col => cellValue(sheet, col, row))
            }
            val shown = matched.take(math.max(0, limit))

            val labels = selectedCols.map { col =>
              val letter = Column.from0(col).toLetter
              if hasHeader then
                cellValue(sheet, col, headerRowIdx)
                  .flatMap(textOf)
                  .map(_.trim)
                  .filter(_.nonEmpty)
                  .getOrElse(letter)
              else letter
            }

            format match
              case FilterFormat.Markdown =>
                renderMarkdown(sheet, shown, matched.length, selectedCols, labels)
              case FilterFormat.Csv => renderCsv(sheet, shown, selectedCols, labels)
              case FilterFormat.Json => renderJson(sheet, shown, selectedCols, labels)
          }

  // ==========================================================================
  // Column selection
  // ==========================================================================

  /** Parse "A,C:E" into 0-based column indices; None = all used columns. */
  private def parseColumns(
    spec: Option[String],
    usedCols: Vector[Int]
  ): Either[String, Vector[Int]] =
    spec match
      case None => Right(usedCols)
      case Some(s) =>
        s.split(",").toVector.map(_.trim).filter(_.nonEmpty).flatTraverse { token =>
          token.split(":", -1) match
            case Array(single) =>
              Column.fromLetter(single.trim.toUpperCase).map(c => Vector(c.index0))
            case Array(lo, hi) =>
              for
                l <- Column.fromLetter(lo.trim.toUpperCase)
                h <- Column.fromLetter(hi.trim.toUpperCase)
              yield (math.min(l.index0, h.index0) to math.max(l.index0, h.index0)).toVector
            case _ => Left(s"Invalid --columns token '$token' (use A or A:C)")
        }

  private def usedColumnsLabel(range: com.tjclp.xl.addressing.CellRange): String =
    s"${range.start.col.toLetter}:${range.end.col.toLetter}"

  // ==========================================================================
  // Cell access & display
  // ==========================================================================

  private def cellValue(sheet: Sheet, col: Int, row: Int): Option[CellValue] =
    sheet.cells.get(ARef.from0(col, row)).map(_.value)

  private def textOf(value: CellValue): Option[String] = value match
    case CellValue.Text(s) => Some(s)
    case CellValue.RichText(rt) => Some(rt.toPlainText)
    case _ => None

  /** NumFmt-aware display string (mirrors the view renderers; formulas show cached values). */
  private def displayCell(sheet: Sheet, col: Int, row: Int): String =
    sheet.cells
      .get(ARef.from0(col, row))
      .map { cell =>
        val numFmt = cell.styleId
          .flatMap(sheet.styleRegistry.get)
          .map(_.numFmt)
          .getOrElse(NumFmt.General)
        cell.value match
          case CellValue.Formula(expr, cached) =>
            val displayExpr = if expr.startsWith("=") then expr else s"=$expr"
            cached.map(cv => NumFmtFormatter.formatValue(cv, numFmt)).getOrElse(displayExpr)
          case CellValue.RichText(rt) => rt.toPlainText
          case CellValue.Bool(b) => if b then "TRUE" else "FALSE"
          case CellValue.Error(err) => err.toExcel
          case CellValue.Empty => ""
          case other => NumFmtFormatter.formatValue(other, numFmt)
      }
      .getOrElse("")

  // ==========================================================================
  // Rendering
  // ==========================================================================

  private def renderNoData(format: FilterFormat, labels: Vector[String]): String =
    format match
      case FilterFormat.Markdown => "No rows matched."
      case FilterFormat.Csv => ("row" +: labels).mkString(",")
      case FilterFormat.Json => "[]"

  private def renderMarkdown(
    sheet: Sheet,
    shown: Vector[Int],
    totalMatched: Int,
    cols: Vector[Int],
    labels: Vector[String]
  ): String =
    if totalMatched == 0 then "No rows matched."
    else
      val headers = "Row" +: labels
      val rows = shown.map { row =>
        (row + 1).toString +: cols.map(col => displayCell(sheet, col, row))
      }
      val widths = headers.indices.map { i =>
        val dataMax = rows.map(r => escapeMd(r(i)).length).maxOption.getOrElse(0)
        math.max(3, math.max(headers(i).length, dataMax))
      }
      val sb = new StringBuilder
      sb.append(headers.zip(widths).map((h, w) => s" ${h.padTo(w, ' ')} ").mkString("|", "|", "|"))
      sb.append("\n")
      sb.append(widths.map(w => "-" * (w + 2)).mkString("|", "|", "|"))
      sb.append("\n")
      rows.foreach { r =>
        sb.append(
          r.zip(widths).map((c, w) => s" ${escapeMd(c).padTo(w, ' ')} ").mkString("|", "|", "|")
        )
        sb.append("\n")
      }
      if totalMatched > shown.length then
        sb.append(s"\nMatched $totalMatched row(s); showing first ${shown.length} (--limit).\n")
      else sb.append(s"\n$totalMatched row(s) matched.\n")
      sb.toString

  private def escapeMd(s: String): String =
    s.replace("|", "\\|").replace("\n", " ").replace("\r", "")

  private def renderCsv(
    sheet: Sheet,
    shown: Vector[Int],
    cols: Vector[Int],
    labels: Vector[String]
  ): String =
    val header = ("row" +: labels).map(escapeCsv).mkString(",")
    val lines = shown.map { row =>
      ((row + 1).toString +: cols.map(col => displayCell(sheet, col, row)))
        .map(escapeCsv)
        .mkString(",")
    }
    (header +: lines).mkString("\n")

  /** RFC 4180: quote when the value contains comma, quote, or newline; double inner quotes. */
  private def escapeCsv(s: String): String =
    val needsQuoting = s.contains(',') || s.contains('"') || s.contains('\n') || s.contains('\r')
    if needsQuoting then s"\"${s.replace("\"", "\"\"")}\"" else s

  private def renderJson(
    sheet: Sheet,
    shown: Vector[Int],
    cols: Vector[Int],
    labels: Vector[String]
  ): String =
    val rows = shown.map { row =>
      val cells = ujson.Obj()
      cols.zip(labels).foreach { (col, label) =>
        cells(label) = jsonValue(cellValue(sheet, col, row))
      }
      ujson.Obj("row" -> ujson.Num((row + 1).toDouble), "cells" -> cells)
    }
    ujson.write(ujson.Arr.from(rows), indent = 2)

  private def jsonValue(value: Option[CellValue]): ujson.Value =
    value match
      case None => ujson.Null
      case Some(v) =>
        v match
          case CellValue.Number(n) => ujson.Num(n.toDouble)
          case CellValue.Text(s) => ujson.Str(s)
          case CellValue.Bool(b) => ujson.Bool(b)
          case CellValue.DateTime(dt) => ujson.Str(dt.toString)
          case CellValue.Error(err) => ujson.Str(err.toExcel)
          case CellValue.RichText(rt) => ujson.Str(rt.toPlainText)
          case CellValue.Empty => ujson.Null
          case CellValue.Formula(expr, cached) =>
            cached
              .map(c => jsonValue(Some(c)))
              .getOrElse(ujson.Str(if expr.startsWith("=") then expr else s"=$expr"))
