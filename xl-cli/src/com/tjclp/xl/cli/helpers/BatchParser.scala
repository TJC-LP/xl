package com.tjclp.xl.cli.helpers

import cats.effect.{IO, Resource}
import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, Column, RefType, Row, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formatted.{Formatted, FormattedParsers}
import com.tjclp.xl.formula.{
  FormulaParser,
  FormulaPrinter,
  FormulaShifter,
  ParseError,
  SheetEvaluator
}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

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
   * Parsed value with optional explicit format.
   *
   * Used for typed JSON values (native numbers, booleans) and smart-detected strings (currency,
   * percent, dates).
   */
  final case class ParsedValue(cellValue: CellValue, format: Option[NumFmt])

  /**
   * Batch operation ADT.
   */
  enum BatchOp:
    /** Put a single value to a cell with optional format */
    case Put(ref: String, value: CellValue, format: Option[NumFmt])

    /** Put a formula to a single cell */
    case PutFormula(ref: String, formula: String)

    /** Put a formula to a range with dragging (from anchor cell) */
    case PutFormulaDragging(range: String, formula: String, from: String)

    /** Put explicit formulas to a range (no dragging) */
    case PutFormulas(range: String, formulas: Vector[String])

    /** Put explicit values to a range (row-major order) */
    case PutValues(range: String, values: Vector[ParsedValue])
    case Style(range: String, props: StyleProps)
    case Merge(range: String)
    case Unmerge(range: String)
    case ColWidth(col: String, width: Double)
    case RowHeight(row: Int, height: Double)

  /**
   * Result of batch parsing with optional warnings.
   *
   * @param ops
   *   Parsed batch operations
   * @param warnings
   *   Non-fatal warnings (e.g., unknown properties ignored)
   */
  final case class ParseResult(ops: Vector[BatchOp], warnings: Vector[String])

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
   *   IO containing parsed result with operations and warnings
   */
  def parseBatchOperations(input: String): IO[ParseResult] =
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
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  def parseBatchJson(json: String): Either[Exception, ParseResult] =
    try
      val parsed = ujson.read(json)
      val arr = parsed.arrOpt.getOrElse(
        throw new Exception("Batch input must be a JSON array")
      )

      // Collect warnings during parsing
      val warnings = scala.collection.mutable.ListBuffer[String]()

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
            collectUnknownPropsWarning(objMap, knownPutProps, "put", idx).foreach(warnings += _)
            val ref = requireString(objMap, "ref", idx)
            // detect defaults to true; set to false to disable smart detection
            val detect = objMap.get("detect").flatMap(_.boolOpt).getOrElse(true)
            // Check for explicit values array first (like putf's "values" support)
            objMap.get("values") match
              case Some(arr) if arr.arrOpt.isDefined =>
                val values = arr.arr.toVector.zipWithIndex.map { case (v, i) =>
                  parseJsonValue(v, idx, i, detect)
                }
                BatchOp.PutValues(ref, values)
              case _ =>
                val parsed = parseTypedValue(objMap, idx, detect)
                BatchOp.Put(ref, parsed.cellValue, parsed.format)

          case "putf" =>
            collectUnknownPropsWarning(objMap, knownPutfProps, "putf", idx).foreach(warnings += _)
            val ref = requireString(objMap, "ref", idx)
            // Check for explicit formulas array first
            objMap.get("values") match
              case Some(arr) if arr.arrOpt.isDefined =>
                val formulas = arr.arr.toVector.zipWithIndex.map { case (v, i) =>
                  v.strOpt.getOrElse(
                    throw new Exception(
                      s"Object ${idx + 1}: 'values[$i]' must be a string formula"
                    )
                  )
                }
                BatchOp.PutFormulas(ref, formulas)
              case _ =>
                val formula = requireStringValue(objMap, idx)
                // Check for 'from' field for formula dragging
                objMap.get("from").flatMap(_.strOpt) match
                  case Some(fromRef) => BatchOp.PutFormulaDragging(ref, formula, fromRef)
                  case None => BatchOp.PutFormula(ref, formula)

          case "style" =>
            collectUnknownPropsWarning(objMap, knownStyleProps, "style", idx).foreach(warnings += _)
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

      Right(ParseResult(ops, warnings.toVector))
    catch
      case e: ujson.ParseException =>
        Left(new Exception(s"JSON parse error: ${e.getMessage}"))
      case e: Exception =>
        Left(e)

  /** Type alias for uPickle's LinkedHashMap. */
  private type ObjMap = upickle.core.LinkedHashMap[String, ujson.Value]

  // ========== Known Properties for Validation ==========

  /** Known properties for 'put' operation */
  private val knownPutProps = Set("op", "ref", "value", "values", "format", "detect")

  /** Known properties for 'putf' operation */
  private val knownPutfProps = Set("op", "ref", "value", "values", "from")

  /** Known properties for 'style' operation */
  private val knownStyleProps = Set(
    "op",
    "range",
    "bold",
    "italic",
    "underline",
    "bg",
    "fg",
    "fontSize",
    "fontName",
    "align",
    "valign",
    "wrap",
    "numFormat",
    "border",
    "borderTop",
    "borderRight",
    "borderBottom",
    "borderLeft",
    "borderColor",
    "replace"
  )

  /** Collect warning about unknown properties in a batch operation (if any) */
  private def collectUnknownPropsWarning(
    objMap: ObjMap,
    known: Set[String],
    opType: String,
    idx: Int
  ): Option[String] =
    val keys = objMap.keys.toSet
    val unknown = keys -- known
    if unknown.nonEmpty then
      Some(
        s"Warning: Object ${idx + 1} ($opType): unknown properties ignored: ${unknown.mkString(", ")}"
      )
    else None

  // ========== Format Name Parsing ==========

  /**
   * Parse format name string to NumFmt.
   *
   * Supports named formats (currency, percent, date, etc.) and custom Excel format codes.
   */
  private def parseFormatName(name: String): Option[NumFmt] =
    name.toLowerCase match
      case "general" => Some(NumFmt.General)
      case "integer" => Some(NumFmt.Integer)
      case "decimal" | "number" => Some(NumFmt.Decimal)
      case "currency" => Some(NumFmt.Currency)
      case "percent" => Some(NumFmt.Percent)
      case "percent_decimal" => Some(NumFmt.PercentDecimal)
      case "date" => Some(NumFmt.Date)
      case "datetime" => Some(NumFmt.DateTime)
      case "time" => Some(NumFmt.Time)
      case "text" => Some(NumFmt.Text)
      case custom =>
        // Accept custom format codes that contain format characters
        if custom.contains("#") || custom.contains("0") || custom.contains("@") ||
          custom.toLowerCase.contains("yy") || custom.toLowerCase.contains("mm") ||
          custom.toLowerCase.contains("dd") || custom.toLowerCase.contains("hh")
        then Some(NumFmt.Custom(name)) // Preserve original case
        else None

  // ========== Typed Value Parsing ==========

  /**
   * Parse a JSON value with optional explicit format.
   *
   * Handles:
   *   - Native JSON numbers → Number CellValue
   *   - Native JSON booleans → Bool CellValue
   *   - JSON null → Empty CellValue
   *   - Strings with explicit format → parsed according to format
   *   - Strings without format → smart detection (currency, percent, date, number, text)
   *
   * @param detect
   *   If true (default), auto-detect currency/percent/date from strings. If false, treat strings as
   *   plain text unless explicit format is provided.
   */
  private def parseTypedValue(objMap: ObjMap, idx: Int, detect: Boolean = true): ParsedValue =
    val explicitFormat = objMap.get("format").flatMap(_.strOpt).flatMap(parseFormatName)

    objMap
      .get("value")
      .map { json =>
        // Handle native JSON types
        json.numOpt match
          case Some(num) =>
            // Native JSON number - use explicit format if provided
            ParsedValue(CellValue.Number(BigDecimal(num)), explicitFormat)
          case None =>
            json.boolOpt match
              case Some(b) =>
                // Native JSON boolean
                ParsedValue(CellValue.Bool(b), None)
              case None =>
                if json.isNull then
                  // JSON null → Empty
                  ParsedValue(CellValue.Empty, None)
                else
                  json.strOpt match
                    case Some(s) =>
                      // String value - use explicit format or smart detection
                      parseStringValue(s, explicitFormat, idx, detect)
                    case None =>
                      throw new Exception(
                        s"Object ${idx + 1}: 'value' must be string, number, boolean, or null"
                      )
      }
      .getOrElse(
        throw new Exception(s"Object ${idx + 1}: Missing 'value' field")
      )

  /**
   * Parse a string value with optional explicit format.
   *
   * With explicit format: parse according to format (currency strings become numbers, etc.) Without
   * explicit format: smart detection for currency, percent, dates (if detect=true).
   *
   * @param detect
   *   If true, auto-detect currency/percent/date patterns. If false, treat as plain text.
   */
  private def parseStringValue(
    s: String,
    explicitFormat: Option[NumFmt],
    idx: Int,
    detect: Boolean = true
  ): ParsedValue =
    explicitFormat match
      case Some(fmt) =>
        // Explicit format - parse string to appropriate type
        fmt match
          case NumFmt.Currency =>
            FormattedParsers
              .parseMoney(s)
              .orElse(FormattedParsers.parseAccounting(s))
              .map(f => ParsedValue(f.value, Some(f.numFmt)))
              .getOrElse {
                // If string doesn't parse as money, try as number with currency format
                scala.util
                  .Try(BigDecimal(s))
                  .toOption
                  .map(n => ParsedValue(CellValue.Number(n), Some(fmt)))
                  .getOrElse(ParsedValue(CellValue.Text(s), Some(fmt)))
              }
          case NumFmt.Percent | NumFmt.PercentDecimal =>
            FormattedParsers
              .parsePercent(s)
              .map(f => ParsedValue(f.value, Some(fmt)))
              .getOrElse {
                // Try as plain number (user wants percent display)
                scala.util
                  .Try(BigDecimal(s))
                  .toOption
                  .map(n => ParsedValue(CellValue.Number(n), Some(fmt)))
                  .getOrElse(ParsedValue(CellValue.Text(s), Some(fmt)))
              }
          case NumFmt.Date =>
            FormattedParsers
              .parseDate(s)
              .map(f => ParsedValue(f.value, Some(f.numFmt)))
              .getOrElse(ParsedValue(CellValue.Text(s), Some(fmt)))
          case NumFmt.DateTime | NumFmt.Time =>
            // Try date parse (includes time in DateTime)
            FormattedParsers
              .parseDate(s)
              .map(f => ParsedValue(f.value, Some(fmt)))
              .getOrElse(ParsedValue(CellValue.Text(s), Some(fmt)))
          case NumFmt.Integer | NumFmt.Decimal | NumFmt.General =>
            // Try to parse as number
            scala.util
              .Try(BigDecimal(s))
              .toOption
              .map(n => ParsedValue(CellValue.Number(n), Some(fmt)))
              .getOrElse(ParsedValue(CellValue.Text(s), Some(fmt)))
          case _ =>
            // Custom format - try number, fall back to text
            scala.util
              .Try(BigDecimal(s))
              .toOption
              .map(n => ParsedValue(CellValue.Number(n), Some(fmt)))
              .getOrElse(ParsedValue(CellValue.Text(s), Some(fmt)))

      case None =>
        // No explicit format
        if detect then
          // Smart detection enabled - detect currency, percent, dates
          detectAndParse(s)
        else
          // Smart detection disabled - treat as plain text
          ParsedValue(CellValue.Text(s), None)

  /**
   * Smart detection for string values without explicit format.
   *
   * Detects:
   *   - Currency: $1,234.56 or ($1,234.56)
   *   - Percent: 59.4%
   *   - ISO Date: 2025-11-10
   *   - Number: 123.45
   *   - Boolean: true/false
   *   - Text: everything else
   */
  private def detectAndParse(s: String): ParsedValue =
    val trimmed = s.trim

    // Currency: starts with $ or parenthesized with $
    if trimmed.startsWith("$") || (trimmed.startsWith("(") && trimmed.contains("$")) then
      FormattedParsers
        .parseMoney(trimmed)
        .orElse(FormattedParsers.parseAccounting(trimmed))
        .map(f => ParsedValue(f.value, Some(f.numFmt)))
        .getOrElse(ParsedValue(CellValue.Text(s), None))

    // Percent: ends with %
    else if trimmed.endsWith("%") then
      FormattedParsers
        .parsePercent(trimmed)
        .map(f => ParsedValue(f.value, Some(f.numFmt)))
        .getOrElse(ParsedValue(CellValue.Text(s), None))

    // ISO Date: YYYY-MM-DD pattern
    else if trimmed.matches("""\d{4}-\d{2}-\d{2}""") then
      FormattedParsers
        .parseDate(trimmed)
        .map(f => ParsedValue(f.value, Some(f.numFmt)))
        .getOrElse(ParsedValue(CellValue.Text(s), None))

    // Try number
    else
      scala.util.Try(BigDecimal(trimmed)).toOption match
        case Some(n) => ParsedValue(CellValue.Number(n), None)
        case None =>
          // Boolean detection
          trimmed.toLowerCase match
            case "true" => ParsedValue(CellValue.Bool(true), None)
            case "false" => ParsedValue(CellValue.Bool(false), None)
            case _ =>
              // Strip quotes if present
              val text =
                if trimmed.startsWith("\"") && trimmed.endsWith("\"") then
                  trimmed.drop(1).dropRight(1)
                else s
              ParsedValue(CellValue.Text(text), None)

  /**
   * Parse a single JSON value element (from a "values" array).
   *
   * Handles native JSON types and smart string detection, mirroring parseTypedValue but operating
   * on a bare ujson.Value instead of an ObjMap with a "value" key.
   */
  private def parseJsonValue(
    json: ujson.Value,
    objIdx: Int,
    elemIdx: Int,
    detect: Boolean
  ): ParsedValue =
    json.numOpt match
      case Some(num) =>
        ParsedValue(CellValue.Number(BigDecimal(num)), None)
      case None =>
        json.boolOpt match
          case Some(b) => ParsedValue(CellValue.Bool(b), None)
          case None =>
            if json.isNull then ParsedValue(CellValue.Empty, None)
            else
              json.strOpt match
                case Some(s) =>
                  parseStringValue(s, None, objIdx, detect)
                case None =>
                  throw new Exception(
                    s"Object ${objIdx + 1}: 'values[$elemIdx]' must be string, number, boolean, or null"
                  )

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

  /** Extract value field as string (for formulas) */
  private def requireStringValue(objMap: ObjMap, idx: Int): String =
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
          case BatchOp.Put(refStr, cellValue, format) =>
            applyPutTyped(currentWb, defaultSheetName, refStr, cellValue, format)

          case BatchOp.PutFormula(refStr, formula) =>
            applyPutFormula(currentWb, defaultSheetName, refStr, formula)

          case BatchOp.PutFormulaDragging(rangeStr, formula, fromRef) =>
            applyPutFormulaDragging(currentWb, defaultSheetName, rangeStr, formula, fromRef)

          case BatchOp.PutFormulas(rangeStr, formulas) =>
            applyPutFormulas(currentWb, defaultSheetName, rangeStr, formulas)

          case BatchOp.PutValues(rangeStr, values) =>
            applyPutValues(currentWb, defaultSheetName, rangeStr, values)

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

  /** Apply a typed put operation with optional format. */
  private def applyPutTyped(
    wb: Workbook,
    defaultSheetName: Option[SheetName],
    refStr: String,
    cellValue: CellValue,
    format: Option[NumFmt]
  ): IO[Workbook] =
    IO.fromEither(RefType.parse(refStr).left.map(e => new Exception(e))).flatMap {
      case RefType.Cell(ref) =>
        defaultSheetName match
          case Some(sheetName) =>
            updateSheetWithFormat(wb, sheetName, ref, cellValue, format)
          case None =>
            IO.raiseError(new Exception(s"batch requires --sheet for unqualified ref '$refStr'"))

      case RefType.QualifiedCell(sheetName, ref) =>
        updateSheetWithFormat(wb, sheetName, ref, cellValue, format)

      case RefType.Range(_) | RefType.QualifiedRange(_, _) =>
        IO.raiseError(new Exception(s"batch put requires single cell ref, not range: $refStr"))
    }

  /** Update sheet with value and optional format */
  private def updateSheetWithFormat(
    wb: Workbook,
    sheetName: SheetName,
    ref: ARef,
    cellValue: CellValue,
    format: Option[NumFmt]
  ): IO[Workbook] =
    format match
      case Some(numFmt) =>
        // Use Formatted to apply both value and style
        val formatted = Formatted(cellValue, numFmt)
        updateSheet(wb, sheetName)(_.put(ref, formatted))
      case None =>
        // No format - just put the value
        updateSheet(wb, sheetName)(_.put(ref, cellValue))

  /** Apply a single formula to a cell */
  private def applyPutFormula(
    wb: Workbook,
    defaultSheetName: Option[SheetName],
    refStr: String,
    formulaStr: String
  ): IO[Workbook] =
    val formula = if formulaStr.startsWith("=") then formulaStr.drop(1) else formulaStr
    val value = CellValue.Formula(formula, None)

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
        IO.raiseError(
          new Exception(
            s"batch putf with single formula requires single cell ref, not range: $refStr. " +
              "Use 'from' field for formula dragging or 'values' array for explicit formulas."
          )
        )
    }

  /** Apply formula with dragging to a range */
  private def applyPutFormulaDragging(
    wb: Workbook,
    defaultSheetName: Option[SheetName],
    rangeStr: String,
    formulaStr: String,
    fromRef: String
  ): IO[Workbook] =
    val formula = if formulaStr.startsWith("=") then formulaStr.drop(1) else formulaStr
    val fullFormula = s"=$formula"

    for
      rangeRef <- parseRangeRef(rangeStr, defaultSheetName)
      (sheetName, range) = rangeRef

      // Parse the 'from' reference
      fromARef <- IO.fromEither(
        ARef.parse(fromRef).left.map(e => new Exception(s"Invalid 'from' reference: $e"))
      )

      // Parse the formula
      parsedExpr <- IO.fromEither(
        FormulaParser.parse(fullFormula).left.map { e =>
          new Exception(ParseError.formatWithContext(e, fullFormula))
        }
      )

      // Apply formula with shifting
      result <- updateSheet(wb, sheetName) { sheet =>
        val startCol = Column.index0(fromARef.col)
        val startRow = Row.index0(fromARef.row)

        range.cells.foldLeft(sheet) { (s, targetRef) =>
          val colDelta = Column.index0(targetRef.col) - startCol
          val rowDelta = Row.index0(targetRef.row) - startRow
          val shiftedExpr = FormulaShifter.shift(parsedExpr, colDelta, rowDelta)
          val shiftedFormula = FormulaPrinter.print(shiftedExpr, includeEquals = false)
          val cachedValue =
            SheetEvaluator.evaluateFormula(s)(s"=$shiftedFormula", workbook = Some(wb)).toOption
          s.put(targetRef, CellValue.Formula(shiftedFormula, cachedValue))
        }
      }
    yield result

  /** Apply explicit formulas to a range (no dragging) */
  private def applyPutFormulas(
    wb: Workbook,
    defaultSheetName: Option[SheetName],
    rangeStr: String,
    formulas: Vector[String]
  ): IO[Workbook] =
    for
      rangeRef <- parseRangeRef(rangeStr, defaultSheetName)
      (sheetName, range) = rangeRef

      // Validate count matches
      cellCount = range.cellCount.toInt
      _ <-
        if cellCount != formulas.length then
          IO.raiseError(
            new Exception(
              s"Range ${range.toA1} has $cellCount cells but ${formulas.length} formulas provided."
            )
          )
        else IO.unit

      // Apply formulas
      result <- updateSheet(wb, sheetName) { sheet =>
        range.cellsRowMajor.zip(formulas.iterator).foldLeft(sheet) { case (s, (ref, formulaStr)) =>
          val formula = if formulaStr.startsWith("=") then formulaStr.drop(1) else formulaStr
          val cachedValue =
            SheetEvaluator.evaluateFormula(s)(s"=$formula", workbook = Some(wb)).toOption
          s.put(ref, CellValue.Formula(formula, cachedValue))
        }
      }
    yield result

  /** Apply explicit values to a range (row-major order, like applyPutFormulas). */
  private def applyPutValues(
    wb: Workbook,
    defaultSheetName: Option[SheetName],
    rangeStr: String,
    values: Vector[ParsedValue]
  ): IO[Workbook] =
    for
      rangeRef <- parseRangeRef(rangeStr, defaultSheetName)
      (sheetName, range) = rangeRef

      // Validate count matches
      cellCount = range.cellCount.toInt
      _ <-
        if cellCount != values.length then
          IO.raiseError(
            new Exception(
              s"Range ${range.toA1} has $cellCount cells but ${values.length} values provided."
            )
          )
        else IO.unit

      // Apply values
      result <- updateSheet(wb, sheetName) { sheet =>
        range.cellsRowMajor.zip(values.iterator).foldLeft(sheet) { case (s, (ref, pv)) =>
          pv.format match
            case Some(numFmt) =>
              val formatted = Formatted(pv.cellValue, numFmt)
              s.put(ref, formatted)
            case None =>
              s.put(ref, pv.cellValue)
        }
      }
    yield result

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
