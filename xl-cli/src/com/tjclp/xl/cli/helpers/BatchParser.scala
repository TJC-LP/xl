package com.tjclp.xl.cli.helpers

import cats.effect.{IO, Resource}
import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{CellRange, Column, RefType, Row, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.styles.CellStyle

/**
 * Batch operation parsing and execution for CLI.
 *
 * Handles parsing JSON batch input and applying operations to workbooks. Uses uPickle for robust
 * JSON parsing that handles all edge cases (nested braces, escaping, unicode).
 */
object BatchParser:

  /**
   * Style properties for batch style operations.
   *
   * All fields are optional - only specified properties are applied. When `replace` is false
   * (default), styles are merged with existing cell styles.
   */
  final case class StyleProps(
    bold: Boolean = false,
    italic: Boolean = false,
    underline: Boolean = false,
    bg: Option[String] = None,
    fg: Option[String] = None,
    fontSize: Option[Double] = None,
    fontName: Option[String] = None,
    align: Option[String] = None,
    valign: Option[String] = None,
    wrap: Boolean = false,
    numFormat: Option[String] = None,
    border: Option[String] = None,
    borderTop: Option[String] = None,
    borderRight: Option[String] = None,
    borderBottom: Option[String] = None,
    borderLeft: Option[String] = None,
    borderColor: Option[String] = None,
    replace: Boolean = false
  )

  /**
   * Batch operation ADT.
   */
  enum BatchOp:
    case Put(ref: String, value: String)
    case PutFormula(ref: String, formula: String)
    case Style(range: String, props: StyleProps)
    case Merge(range: String)
    case Unmerge(range: String)
    case ColWidth(col: String, width: Double)
    case RowHeight(row: Int, height: Double)

  /**
   * Read batch input from file or stdin.
   *
   * @param source
   *   File path or "-" for stdin
   * @return
   *   IO containing input string
   */
  def readBatchInput(source: String): IO[String] =
    if source == "-" then IO.blocking(scala.io.Source.stdin.mkString)
    else
      Resource
        .fromAutoCloseable(IO.blocking(scala.io.Source.fromFile(source)))
        .use(src => IO.blocking(src.mkString))

  /**
   * Parse batch JSON input. Expects format:
   * {{{
   * [
   *   {"op": "put", "ref": "A1", "value": "Hello"},
   *   {"op": "putf", "ref": "B1", "value": "=A1*2"}
   * ]
   * }}}
   *
   * @param input
   *   JSON string
   * @return
   *   IO containing parsed operations
   */
  def parseBatchOperations(input: String): IO[Vector[BatchOp]] =
    IO.fromEither {
      val trimmed = input.trim
      if !trimmed.startsWith("[") then Left(new Exception("Batch input must be a JSON array"))
      else parseBatchJson(trimmed)
    }

  /**
   * Parse batch JSON using uPickle.
   *
   * Handles all JSON edge cases: nested braces, escaping, unicode, multi-line values.
   *
   * Supported operations:
   *   - `put`: {"op": "put", "ref": "A1", "value": "Hello"}
   *   - `putf`: {"op": "putf", "ref": "A1", "value": "=SUM(A1:A10)"}
   *   - `style`: {"op": "style", "range": "A1:B2", "bold": true, "bg": "#FFFF00"}
   *   - `merge`: {"op": "merge", "range": "A1:D1"}
   *   - `unmerge`: {"op": "unmerge", "range": "A1:D1"}
   *   - `colwidth`: {"op": "colwidth", "col": "A", "width": 15.5}
   *   - `rowheight`: {"op": "rowheight", "row": 1, "height": 30}
   */
  def parseBatchJson(json: String): Either[Exception, Vector[BatchOp]] =
    try
      val parsed = ujson.read(json)
      val arr = parsed.arrOpt.getOrElse(
        throw new Exception("Batch input must be a JSON array")
      )

      val ops = arr.value.toVector.zipWithIndex.map { case (obj, idx) =>
        val objMap = obj.objOpt.getOrElse(
          throw new Exception(
            s"Object ${idx + 1}: Expected JSON object, got ${obj.getClass.getSimpleName}"
          )
        )

        val op = objMap
          .get("op")
          .flatMap(_.strOpt)
          .getOrElse(
            throw new Exception(
              s"Object ${idx + 1}: Missing or invalid 'op' field"
            )
          )

        op match
          case "put" =>
            val ref = requireString(objMap, "ref", idx)
            val value = requireValue(objMap, idx)
            BatchOp.Put(ref, value)

          case "putf" =>
            val ref = requireString(objMap, "ref", idx)
            val value = requireValue(objMap, idx)
            BatchOp.PutFormula(ref, value)

          case "style" =>
            val range = requireString(objMap, "range", idx)
            val props = parseStyleProps(objMap)
            BatchOp.Style(range, props)

          case "merge" =>
            val range = requireString(objMap, "range", idx)
            BatchOp.Merge(range)

          case "unmerge" =>
            val range = requireString(objMap, "range", idx)
            BatchOp.Unmerge(range)

          case "colwidth" =>
            val col = requireString(objMap, "col", idx)
            val width = requireNumber(objMap, "width", idx)
            BatchOp.ColWidth(col, width)

          case "rowheight" =>
            val row = requireInt(objMap, "row", idx)
            val height = requireNumber(objMap, "height", idx)
            BatchOp.RowHeight(row, height)

          case other =>
            throw new Exception(
              s"Object ${idx + 1}: Unknown operation '$other'. " +
                "Valid: put, putf, style, merge, unmerge, colwidth, rowheight"
            )
      }

      Right(ops)
    catch
      case e: ujson.ParseException =>
        Left(new Exception(s"JSON parse error: ${e.getMessage}"))
      case e: Exception =>
        Left(e)

  /** Type alias for uPickle's LinkedHashMap. */
  private type ObjMap = upickle.core.LinkedHashMap[String, ujson.Value]

  /** Extract required string field from JSON object. */
  private def requireString(objMap: ObjMap, field: String, idx: Int): String =
    objMap
      .get(field)
      .flatMap(_.strOpt)
      .getOrElse(
        throw new Exception(s"Object ${idx + 1}: Missing or invalid '$field' field")
      )

  /** Extract required numeric field from JSON object. */
  private def requireNumber(objMap: ObjMap, field: String, idx: Int): Double =
    objMap
      .get(field)
      .flatMap(_.numOpt)
      .getOrElse(
        throw new Exception(
          s"Object ${idx + 1}: Missing or invalid '$field' field (expected number)"
        )
      )

  /** Extract required integer field from JSON object. */
  private def requireInt(objMap: ObjMap, field: String, idx: Int): Int =
    objMap
      .get(field)
      .flatMap(_.numOpt)
      .map(_.toInt)
      .getOrElse(
        throw new Exception(
          s"Object ${idx + 1}: Missing or invalid '$field' field (expected integer)"
        )
      )

  /** Extract value field (string, number, boolean, or null). */
  private def requireValue(objMap: ObjMap, idx: Int): String =
    objMap
      .get("value")
      .map {
        case v if v.strOpt.isDefined => v.str
        case v if v.numOpt.isDefined => v.num.toString
        case v if v.boolOpt.isDefined => v.bool.toString
        case v if v.isNull => ""
        case _ => throw new Exception(s"Object ${idx + 1}: Unsupported value type for 'value'")
      }
      .getOrElse(
        throw new Exception(s"Object ${idx + 1}: Missing 'value' field")
      )

  /** Parse style properties from JSON object. */
  private def parseStyleProps(objMap: ObjMap): StyleProps =
    StyleProps(
      bold = objMap.get("bold").flatMap(_.boolOpt).getOrElse(false),
      italic = objMap.get("italic").flatMap(_.boolOpt).getOrElse(false),
      underline = objMap.get("underline").flatMap(_.boolOpt).getOrElse(false),
      bg = objMap.get("bg").flatMap(_.strOpt),
      fg = objMap.get("fg").flatMap(_.strOpt),
      fontSize = objMap.get("fontSize").flatMap(_.numOpt),
      fontName = objMap.get("fontName").flatMap(_.strOpt),
      align = objMap.get("align").flatMap(_.strOpt),
      valign = objMap.get("valign").flatMap(_.strOpt),
      wrap = objMap.get("wrap").flatMap(_.boolOpt).getOrElse(false),
      numFormat = objMap.get("numFormat").flatMap(_.strOpt),
      border = objMap.get("border").flatMap(_.strOpt),
      borderTop = objMap.get("borderTop").flatMap(_.strOpt),
      borderRight = objMap.get("borderRight").flatMap(_.strOpt),
      borderBottom = objMap.get("borderBottom").flatMap(_.strOpt),
      borderLeft = objMap.get("borderLeft").flatMap(_.strOpt),
      borderColor = objMap.get("borderColor").flatMap(_.strOpt),
      replace = objMap.get("replace").flatMap(_.boolOpt).getOrElse(false)
    )

  /**
   * Apply batch operations to a workbook.
   *
   * Operations are applied **in order** for deterministic results. This ensures that sequences like
   * put → style → merge work correctly.
   *
   * @param wb
   *   Workbook to modify
   * @param defaultSheetOpt
   *   Default sheet for unqualified refs
   * @param ops
   *   Operations to apply
   * @return
   *   IO containing updated workbook
   */
  def applyBatchOperations(
    wb: Workbook,
    defaultSheetOpt: Option[Sheet],
    ops: Vector[BatchOp]
  ): IO[Workbook] =
    val defaultSheetName = defaultSheetOpt.map(_.name)

    ops.foldLeft(IO.pure(wb)) { (wbIO, op) =>
      wbIO.flatMap { currentWb =>
        op match
          case BatchOp.Put(refStr, valueStr) =>
            applyPut(currentWb, defaultSheetName, refStr, valueStr, isFormula = false)

          case BatchOp.PutFormula(refStr, formula) =>
            applyPut(currentWb, defaultSheetName, refStr, formula, isFormula = true)

          case BatchOp.Style(rangeStr, props) =>
            applyStyle(currentWb, defaultSheetName, rangeStr, props)

          case BatchOp.Merge(rangeStr) =>
            applyMerge(currentWb, defaultSheetName, rangeStr)

          case BatchOp.Unmerge(rangeStr) =>
            applyUnmerge(currentWb, defaultSheetName, rangeStr)

          case BatchOp.ColWidth(colStr, width) =>
            applyColWidth(currentWb, defaultSheetName, colStr, width)

          case BatchOp.RowHeight(row, height) =>
            applyRowHeight(currentWb, defaultSheetName, row, height)
      }
    }

  // ========== Operation Helpers ==========

  /** Apply a put or putf operation. */
  private def applyPut(
    wb: Workbook,
    defaultSheetName: Option[SheetName],
    refStr: String,
    valueStr: String,
    isFormula: Boolean
  ): IO[Workbook] =
    val value =
      if isFormula then
        val formula = if valueStr.startsWith("=") then valueStr.drop(1) else valueStr
        CellValue.Formula(formula, None)
      else ValueParser.parseValue(valueStr)

    IO.fromEither(RefType.parse(refStr).left.map(e => new Exception(e))).flatMap {
      case RefType.Cell(ref) =>
        defaultSheetName match
          case Some(sheetName) =>
            updateSheet(wb, sheetName)(_.put(ref -> value))
          case None =>
            IO.raiseError(new Exception(s"batch requires --sheet for unqualified ref '$refStr'"))

      case RefType.QualifiedCell(sheetName, ref) =>
        updateSheet(wb, sheetName)(_.put(ref -> value))

      case RefType.Range(_) | RefType.QualifiedRange(_, _) =>
        IO.raiseError(new Exception(s"batch put requires single cell ref, not range: $refStr"))
    }

  /** Apply a style operation. */
  private def applyStyle(
    wb: Workbook,
    defaultSheetName: Option[SheetName],
    rangeStr: String,
    props: StyleProps
  ): IO[Workbook] =
    for
      rangeRef <- parseRangeRef(rangeStr, defaultSheetName)
      (sheetName, range) = rangeRef
      cellStyle <- StyleBuilder.buildCellStyle(
        bold = props.bold,
        italic = props.italic,
        underline = props.underline,
        bg = props.bg,
        fg = props.fg,
        fontSize = props.fontSize,
        fontName = props.fontName,
        align = props.align,
        valign = props.valign,
        wrap = props.wrap,
        numFormat = props.numFormat,
        border = props.border,
        borderTop = props.borderTop,
        borderRight = props.borderRight,
        borderBottom = props.borderBottom,
        borderLeft = props.borderLeft,
        borderColor = props.borderColor
      )
      result <- updateSheet(wb, sheetName) { sheet =>
        if props.replace then
          // Replace mode: apply style directly to all cells in range
          sheet.style(range, cellStyle)
        else
          // Merge mode: merge with existing styles
          range.cells.foldLeft(sheet) { (s, ref) =>
            val existing = s.getCellStyle(ref).getOrElse(CellStyle.default)
            val merged = StyleBuilder.mergeStyles(existing, cellStyle)
            s.style(ref, merged)
          }
      }
    yield result

  /** Apply a merge operation. */
  private def applyMerge(
    wb: Workbook,
    defaultSheetName: Option[SheetName],
    rangeStr: String
  ): IO[Workbook] =
    for
      rangeRef <- parseRangeRef(rangeStr, defaultSheetName)
      (sheetName, range) = rangeRef
      result <- updateSheet(wb, sheetName)(_.merge(range))
    yield result

  /** Apply an unmerge operation. */
  private def applyUnmerge(
    wb: Workbook,
    defaultSheetName: Option[SheetName],
    rangeStr: String
  ): IO[Workbook] =
    for
      rangeRef <- parseRangeRef(rangeStr, defaultSheetName)
      (sheetName, range) = rangeRef
      result <- updateSheet(wb, sheetName)(_.unmerge(range))
    yield result

  /** Apply a column width operation. */
  private def applyColWidth(
    wb: Workbook,
    defaultSheetName: Option[SheetName],
    colStr: String,
    width: Double
  ): IO[Workbook] =
    for
      col <- IO.fromEither(Column.fromLetter(colStr).left.map(e => new Exception(e)))
      sheetName <- defaultSheetName match
        case Some(name) => IO.pure(name)
        case None => IO.raiseError(new Exception(s"batch colwidth requires --sheet"))
      result <- updateSheet(wb, sheetName) { sheet =>
        val props = sheet.getColumnProperties(col).copy(width = Some(width))
        sheet.setColumnProperties(col, props)
      }
    yield result

  /** Apply a row height operation. */
  private def applyRowHeight(
    wb: Workbook,
    defaultSheetName: Option[SheetName],
    rowNum: Int,
    height: Double
  ): IO[Workbook] =
    for
      row <- IO.pure(Row.from1(rowNum))
      sheetName <- defaultSheetName match
        case Some(name) => IO.pure(name)
        case None => IO.raiseError(new Exception(s"batch rowheight requires --sheet"))
      result <- updateSheet(wb, sheetName) { sheet =>
        val props = sheet.getRowProperties(row).copy(height = Some(height))
        sheet.setRowProperties(row, props)
      }
    yield result

  // ========== Utilities ==========

  /** Parse a range reference (possibly qualified with sheet name). */
  private def parseRangeRef(
    rangeStr: String,
    defaultSheetName: Option[SheetName]
  ): IO[(SheetName, CellRange)] =
    IO.fromEither(RefType.parse(rangeStr).left.map(e => new Exception(e))).flatMap {
      case RefType.Range(range) =>
        defaultSheetName match
          case Some(name) => IO.pure((name, range))
          case None =>
            IO.raiseError(
              new Exception(s"batch requires --sheet for unqualified range '$rangeStr'")
            )

      case RefType.QualifiedRange(sheetName, range) =>
        IO.pure((sheetName, range))

      case RefType.Cell(ref) =>
        // Single cell treated as 1x1 range
        defaultSheetName match
          case Some(name) => IO.pure((name, CellRange(ref, ref)))
          case None =>
            IO.raiseError(new Exception(s"batch requires --sheet for unqualified ref '$rangeStr'"))

      case RefType.QualifiedCell(sheetName, ref) =>
        IO.pure((sheetName, CellRange(ref, ref)))
    }

  /** Update a sheet in the workbook, raising error if sheet not found. */
  private def updateSheet(wb: Workbook, sheetName: SheetName)(
    f: Sheet => Sheet
  ): IO[Workbook] =
    wb.sheets.find(_.name == sheetName) match
      case None =>
        IO.raiseError(
          new Exception(
            s"Sheet '${sheetName.value}' not found. " +
              s"Available: ${wb.sheetNames.map(_.value).mkString(", ")}"
          )
        )
      case Some(sheet) =>
        IO.pure(wb.put(f(sheet)))
