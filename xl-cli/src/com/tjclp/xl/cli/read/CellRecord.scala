package com.tjclp.xl.cli.read

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue, FormulaKind}
import com.tjclp.xl.cli.output.{Escape, RendererCommon}
import com.tjclp.xl.display.NumFmtFormatter
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.alignment.Align
import com.tjclp.xl.styles.border.{Border, BorderStyle}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.{Font, Underline}
import com.tjclp.xl.styles.numfmt.NumFmt

/** What a cell holds, as the `type` of the `view --format json` cells has always named it. */
enum CellKind derives CanEqual:
  case Text, Number, Boolean, DateTime, Error, RichText, Empty, Formula

  /**
   * The JSON name: `text`, `number`, `boolean`, `datetime`, `error`, `richtext`, `empty`,
   * `formula`.
   */
  def name: String = this match
    case Text => "text"
    case Number => "number"
    case Boolean => "boolean"
    case DateTime => "datetime"
    case Error => "error"
    case RichText => "richtext"
    case Empty => "empty"
    case Formula => "formula"

object CellKind:
  /** The kind of a stored value. */
  def of(value: CellValue): CellKind = value match
    case CellValue.Text(_) => Text
    case CellValue.Number(_) => Number
    case CellValue.Bool(_) => Boolean
    case CellValue.DateTime(_) => DateTime
    case CellValue.Error(_) => Error
    case CellValue.RichText(_) => RichText
    case CellValue.Empty => Empty
    case CellValue.Formula(_, _, _) => Formula

/**
 * The formula side of a formula cell.
 *
 * @param expression
 *   the expression exactly as stored (with or without its leading `=`)
 * @param kind
 *   the CT_CellFormula record kind (normal, array, data table)
 * @param cached
 *   whether the file carried a cached value — when false the record's `value` is `Empty` and its
 *   `formatted` text is blank, and the expression stands in for the value in text renderings
 */
final case class FormulaInfo(expression: String, kind: FormulaKind, cached: Boolean)
    derives CanEqual:

  /** The expression with its leading `=`, as the JSON `formula` field carries it. */
  def text: String = if expression.startsWith("=") then expression else s"=$expression"

  /** The formula-bar spelling: `=…`, braced `{=…}` for array and data-table records. */
  def display: String = RendererCommon.formulaDisplay(expression, kind)

  /** The additive `formulaKind` name; `None` for a normal formula. */
  def kindName: Option[String] = kind match
    case _: FormulaKind.Normal => None
    case _: FormulaKind.ArrayFormula => Some("array")
    case _: FormulaKind.DataTable => Some("dataTable")

/**
 * One cell as every read verb sees it (W2.4): the single projection from which `view`, `cell`,
 * `search`, `stats` and `filter` render, whether the cell came from a loaded [[Sheet]] or from the
 * streaming reader. The scalar textualisers live here and nowhere else.
 *
 * @param value
 *   the scalar the cell shows: the stored value, or for a formula its cached (or evaluated) value —
 *   `Empty` when the formula carries none
 * @param formula
 *   the expression and record kind when the cell is a formula
 * @param hidden
 *   whether the cell's row or column is hidden; `None` from a source without the `hidden`
 *   capability (the streaming reader never sees row/column properties), rendered `null` — never an
 *   affirmative `false` the source cannot vouch for
 * @param mergedInto
 *   the merged range containing the cell, if any (`None` also when the source lacks `merges`)
 */
final case class CellRecord(
  ref: ARef,
  sheet: SheetName,
  kind: CellKind,
  value: CellValue,
  formula: Option[FormulaInfo],
  hidden: Option[Boolean],
  mergedInto: Option[CellRange],
  style: Option[CellStyle]
) derives CanEqual:

  /**
   * The display text of `value` under the cell's number format (General without a style), computed
   * on first use: `search` scans every occupied cell of a sheet but shows only the matches it
   * keeps, so the text of the rest is never built. An `Empty` value formats to the empty string.
   */
  lazy val formatted: String =
    NumFmtFormatter.formatValue(value, style.map(_.numFmt).getOrElse(NumFmt.General))

  /**
   * The scalar's canonical lexeme, the text `search` matches and `cell` prints as `Raw:`: a number
   * with every digit and no thousands or currency decoration, a date-time in ISO form, `TRUE` /
   * `FALSE`, an error token, text as is, and nothing for an empty cell.
   */
  def raw: String = CellRecord.raw(value)

  /** [[raw]], except that an uncached formula shows its expression (there is no value to show). */
  def searchText: String = formula match
    case Some(f) if !f.cached => f.display
    case _ => raw

  /**
   * The text a table cell shows: the formatted value; a formula shows its expression when
   * `showFormulas` asks for it or when it has no cached value to show.
   */
  def text(showFormulas: Boolean): String = formula match
    case Some(f) if showFormulas || !f.cached => f.display
    case _ => formatted

  /**
   * The JSON lexeme of the scalar: numbers and booleans bare, `null` for empty, strings escaped.
   */
  def rawJson: String = CellRecord.rawJson(value)

  /**
   * Whether the cell counts as empty for `--skip-empty`: no value, blank text, or a formula whose
   * cached value is one of those. An uncached formula is not empty — it has an expression to show.
   */
  def isEmpty: Boolean = formula match
    case Some(f) if !f.cached => false
    case _ =>
      value match
        case CellValue.Empty => true
        case CellValue.Text(s) if s.trim.isEmpty => true
        case _ => false

  /** The stored value this record projects: the formula with its cache, or the scalar itself. */
  def cellValue: CellValue = formula match
    case Some(f) =>
      CellValue.Formula(f.expression, Option.when(f.cached)(value), f.kind)
    case None => value

  /**
   * The record as a JSON object, as text so every number lexeme is exactly [[rawJson]].
   *
   * `legacyKeys = true` is the `view --format json` cell, byte for byte: `{"ref", "type"[,
   * "formula"[, "formulaKind"]], "value", "formatted"}`. `legacyKeys = false` is the typed record
   * the newer payloads carry: `{"ref", "sheet", "kind", "value", "formatted", "formula", "hidden",
   * "mergedInto"}` with `formula` an object or `null` and `hidden` a boolean or `null` (unknown) —
   * the style is rendered by [[CellDetail]] alone.
   */
  def toJson(legacyKeys: Boolean): String =
    if legacyKeys then
      val formulaField = formula.fold("") { f =>
        val kindField = f.kindName.fold("")(k => s""", "formulaKind": "$k"""")
        s""", "formula": ${Escape.json(f.text)}$kindField"""
      }
      s"""{"ref": "${ref.toA1}", "type": "${kind.name}"$formulaField, "value": $rawJson, "formatted": ${Escape
          .json(formatted)}}"""
    else s"{$typedFields}"

  /** The fields of the typed shape, without the braces (so `cell --json` can extend them). */
  private[read] def typedFields: String =
    val formulaField = formula.fold("null") { f =>
      s"""{"expression": ${Escape.json(f.text)}, "kind": "${f.kindName
          .getOrElse("normal")}", "cached": ${f.cached}}"""
    }
    val merged = mergedInto.fold("null")(r => s"\"${r.toA1}\"")
    val hiddenField = hidden.fold("null")(_.toString)
    s""""ref": "${ref.toA1}", "sheet": ${Escape.json(sheet.value)}, "kind": "${kind.name}", """ +
      s""""value": $rawJson, "formatted": ${Escape.json(formatted)}, "formula": $formulaField, """ +
      s""""hidden": $hiddenField, "mergedInto": $merged"""

object CellRecord:

  /**
   * Project a stored value: the one place a `CellValue` becomes kind, scalar and display text.
   * `style` supplies the number format (General without one).
   */
  def of(
    sheet: SheetName,
    ref: ARef,
    value: CellValue,
    style: Option[CellStyle],
    hidden: Option[Boolean],
    mergedInto: Option[CellRange]
  ): CellRecord =
    value match
      case CellValue.Formula(expr, cached, kind) =>
        CellRecord(
          ref,
          sheet,
          CellKind.Formula,
          cached.getOrElse(CellValue.Empty),
          Some(FormulaInfo(expr, kind, cached.isDefined)),
          hidden,
          mergedInto,
          style
        )
      case scalar =>
        CellRecord(ref, sheet, CellKind.of(scalar), scalar, None, hidden, mergedInto, style)

  /** An addressed position with no cell in it. */
  def empty(
    sheet: SheetName,
    ref: ARef,
    hidden: Option[Boolean],
    mergedInto: Option[CellRange]
  ): CellRecord =
    of(sheet, ref, CellValue.Empty, None, hidden, mergedInto)

  /** A loaded cell, or the empty record when the sheet has none at `ref`. */
  def project(
    sheet: SheetName,
    ref: ARef,
    cell: Option[Cell],
    style: Cell => Option[CellStyle],
    hidden: Option[Boolean],
    mergedInto: Option[CellRange]
  ): CellRecord =
    cell.fold(empty(sheet, ref, hidden, mergedInto))(c =>
      of(sheet, ref, c.value, style(c), hidden, mergedInto)
    )

  /**
   * A formula record whose value was just evaluated: the evaluated scalar in `value`, or — when the
   * evaluation failed — the Excel-style error token as both the value and its display text, the
   * kind staying `formula` and the expression still present.
   */
  def evaluated(record: CellRecord, result: Either[String, CellValue]): CellRecord =
    record.formula match
      case None => record
      case Some(f) =>
        // A text value formats to itself under every number format, so the token is its own text
        val value = result.fold(token => CellValue.Text(token), identity)
        record.copy(value = value, formula = Some(f.copy(cached = true)))

  /** [[CellRecord.raw]] of a scalar value. */
  def raw(value: CellValue): String = value match
    case CellValue.Text(s) => s
    case CellValue.Number(n) => numberLexeme(n)
    case CellValue.Bool(b) => if b then "TRUE" else "FALSE"
    case CellValue.DateTime(dt) => dt.toString
    case CellValue.Error(err) => err.toExcel
    case CellValue.RichText(rt) => rt.toPlainText
    case CellValue.Empty => ""
    case CellValue.Formula(expr, cached, kind) =>
      cached.fold(RendererCommon.formulaDisplay(expr, kind))(raw)

  /** [[CellRecord.rawJson]] of a scalar value. */
  def rawJson(value: CellValue): String = value match
    case CellValue.Number(n) => numberLexeme(n)
    case CellValue.Bool(b) => if b then "true" else "false"
    case CellValue.Empty => "null"
    case CellValue.Formula(_, cached, _) => cached.fold("null")(rawJson)
    case other => Escape.json(raw(other))

  /** A number with every digit: integers without a point, decimals without trailing zeros. */
  def numberLexeme(n: BigDecimal): String =
    if n.isWhole then n.toBigInt.toString
    else n.underlying.stripTrailingZeros.toPlainString

  /**
   * A cell style as a JSON object, as text: `{"font", "fill", "numFmt", "align", "border"}` — the
   * same facets `cell` prints in its `Style:` block, every one present.
   */
  def styleJson(style: CellStyle): String =
    def color(c: Color): String = c match
      case Color.Rgb(argb) => Escape.json(f"${argb & 0xffffff}%06X")
      case Color.Theme(slot, tint) =>
        val tintStr = if tint == 0.0 then "" else f" tint=$tint%.2f"
        Escape.json(s"$slot$tintStr")
    def side(style: BorderStyle): String = Escape.json(style.toString.toLowerCase)
    val font: Font = style.font
    val fontJson =
      s"""{"name": ${Escape.json(font.name)}, "size": ${numberLexeme(
          BigDecimal(font.sizePt)
        )}, """ +
        s""""bold": ${font.bold}, "italic": ${font.italic}, "underline": ${Escape.json(
            Underline.token(font.underline)
          )}, "color": ${font.color.fold("null")(color)}}"""
    val fillJson = style.fill match
      case Fill.Solid(c) => color(c)
      case Fill.Pattern(_, _, _) => Escape.json("pattern")
      case Fill.None => "null"
    val numFmtJson = style.numFmt match
      case NumFmt.Custom(code) => Escape.json(code)
      case other => Escape.json(other.toString)
    val align: Align = style.align
    val alignJson =
      s"""{"horizontal": ${Escape.json(align.horizontal.toString.toLowerCase)}, """ +
        s""""vertical": ${Escape.json(
            align.vertical.toString.toLowerCase
          )}, "wrap": ${align.wrapText}}"""
    val border: Border = style.border
    val borderJson =
      s"""{"top": ${side(border.top.style)}, "right": ${side(border.right.style)}, """ +
        s""""bottom": ${side(border.bottom.style)}, "left": ${side(border.left.style)}}"""
    s"""{"font": $fontJson, "fill": $fillJson, "numFmt": $numFmtJson, "align": $alignJson, "border": $borderJson}"""
