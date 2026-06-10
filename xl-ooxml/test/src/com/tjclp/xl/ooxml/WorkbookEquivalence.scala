package com.tjclp.xl.ooxml

import java.time.Duration

import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.border.{Border, BorderSide}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.Fill

/**
 * Round-trip equivalence (≈) between an authored workbook and the workbook read back from its
 * serialized bytes — the binding definition for the GH-240 generative round-trip law, reusable by
 * any spec that needs write→read comparison (streaming/in-memory parity, golden fixtures, ...).
 *
 * Two workbooks are ≈-equal when:
 *   - (a) same sheet names, in order
 *   - (b) per sheet: same cell refs; per cell:
 *     - CellValue equal, where:
 *       - Formula compares by formula TEXT only (cached values are write-only hints, the reader's
 *         cache may legitimately differ — e.g. whitespace in cached text)
 *       - Number compares by BigDecimal numeric value (scale-insensitive)
 *       - DateTime ≈ Number(serial): the OOXML cell model stores datetimes as serial numbers; a
 *         DateTime is equal to the serial that converts back to the same instant (±1s)
 *     - RESOLVED CellStyle equal: each side's styleId is resolved through its own StyleRegistry
 *       (unresolvable/dangling ids resolve to CellStyle.default, which is what the writer emits for
 *       them) and compared with numFmtId normalized to None — numFmtId is the raw OOXML format-id
 *       slot the reader fills in, not an authored property
 *     - hyperlink targets equal
 *   - comments equal by (plain text, author) per ref
 *   - merged ranges set-equal
 *   - (c) sheet-level: viewSettings equal; pageSetup equal (an authored PageSetup always
 *     round-trips when it has at least one visible field; all-default PageSetup serializes to no
 *     XML and reads back None by design — generators only produce visible ones)
 *
 * EXPLICITLY IGNORED (serialization noise or known write-only fields):
 *   - styleId numbering (dedup may renumber; only resolved styles matter)
 *   - SST ordering / inline-vs-shared string encoding
 *   - workbook metadata (docProps not written on save — GH-242) and activeSheetIndex (not
 *     serialized to bookViews)
 *   - the worksheet dimension element and derived defaults (defaultColWidth etc.)
 *   - freezePane (three-valued write-only override: None means "preserve", so the reader never
 *     populates it)
 *   - Cell.comment / rich-text comment formatting (Sheet.comments is the comment model; the
 *     author-prefix run the writer adds is presentation, so comments compare by plain text)
 *   - cached formula values (see Formula rule above)
 */
object WorkbookEquivalence:

  /** Tolerance for DateTime-vs-serial comparison (serial doubles lose sub-second precision). */
  private val dateTolerance = Duration.ofSeconds(1)

  /** All ≈ violations between `expected` (authored) and `actual` (read back); empty = ≈-equal. */
  def diff(expected: Workbook, actual: Workbook): List[String] =
    val nameDiffs =
      if expected.sheets.map(_.name.value) != actual.sheets.map(_.name.value) then
        List(
          s"sheet names mismatch: expected ${expected.sheets.map(_.name.value).mkString("[", ", ", "]")}, " +
            s"actual ${actual.sheets.map(_.name.value).mkString("[", ", ", "]")}"
        )
      else Nil
    val sheetDiffs =
      expected.sheets
        .zip(actual.sheets)
        .flatMap { case (exp, act) => sheetDiff(exp, act) }
        .toList
    nameDiffs ++ sheetDiffs

  /** True when `diff` reports no violations. */
  def equivalent(expected: Workbook, actual: Workbook): Boolean =
    diff(expected, actual).isEmpty

  private def sheetDiff(expected: Sheet, actual: Sheet): List[String] =
    val name = expected.name.value
    val refDiffs = keySetDiff(name, "cell refs", expected.cells.keySet, actual.cells.keySet)
    val cellDiffs =
      expected.cells.keySet
        .intersect(actual.cells.keySet)
        .toList
        .sortBy(_.toA1)
        .flatMap { ref =>
          val exp = expected.cells(ref)
          val act = actual.cells(ref)
          valueDiff(name, ref, exp.value, act.value).toList ++
            styleDiff(name, ref, expected, exp, actual, act).toList ++
            hyperlinkDiff(name, ref, exp, act).toList
        }
    val commentDiffs = commentsDiff(name, expected, actual)
    val mergeDiffs =
      if expected.mergedRanges != actual.mergedRanges then
        List(
          s"$name: merged ranges mismatch: expected ${render(expected.mergedRanges.map(_.toA1))}, " +
            s"actual ${render(actual.mergedRanges.map(_.toA1))}"
        )
      else Nil
    val viewDiffs =
      if expected.viewSettings != actual.viewSettings then
        List(
          s"$name: viewSettings mismatch: expected ${expected.viewSettings}, actual ${actual.viewSettings}"
        )
      else Nil
    val pageSetupDiffs =
      if expected.pageSetup != actual.pageSetup then
        List(
          s"$name: pageSetup mismatch: expected ${expected.pageSetup}, actual ${actual.pageSetup}"
        )
      else Nil
    refDiffs ++ cellDiffs ++ commentDiffs ++ mergeDiffs ++ viewDiffs ++ pageSetupDiffs

  private def keySetDiff(
    sheet: String,
    what: String,
    expected: Set[ARef],
    actual: Set[ARef]
  ): List[String] =
    val missing = expected.diff(actual)
    val extra = actual.diff(expected)
    val missingMsg =
      if missing.nonEmpty then List(s"$sheet: missing $what: ${render(missing.map(_.toA1))}")
      else Nil
    val extraMsg =
      if extra.nonEmpty then List(s"$sheet: extra $what: ${render(extra.map(_.toA1))}")
      else Nil
    missingMsg ++ extraMsg

  private def valueDiff(
    sheet: String,
    ref: ARef,
    expected: CellValue,
    actual: CellValue
  ): Option[String] =
    (expected, actual) match
      case (CellValue.Formula(expExpr, _), CellValue.Formula(actExpr, _)) =>
        Option.when(expExpr != actExpr)(
          s"$sheet!${ref.toA1}: formula text mismatch: expected '$expExpr', actual '$actExpr'"
        )
      case (CellValue.Number(exp), CellValue.Number(act)) =>
        // Scala BigDecimal == is scale-insensitive numeric compare, which is exactly the rule
        Option.when(exp != act)(
          s"$sheet!${ref.toA1}: number mismatch: expected $exp, actual $act"
        )
      case (CellValue.DateTime(exp), CellValue.Number(serial)) =>
        val actualDate = CellValue.excelSerialToDateTime(serial.toDouble)
        Option.when(Duration.between(exp, actualDate).abs().compareTo(dateTolerance) > 0)(
          s"$sheet!${ref.toA1}: datetime mismatch: expected $exp, actual serial $serial ($actualDate)"
        )
      case (CellValue.DateTime(exp), CellValue.DateTime(act)) =>
        Option.when(Duration.between(exp, act).abs().compareTo(dateTolerance) > 0)(
          s"$sheet!${ref.toA1}: datetime mismatch: expected $exp, actual $act"
        )
      case _ =>
        Option.when(expected != actual)(
          s"$sheet!${ref.toA1}: value mismatch: expected $expected, actual $actual"
        )

  /**
   * Canonicalize an ARGB color the way the style parser does: a 00 alpha byte means OPAQUE in the
   * wild (Excel/openpyxl write 00RRGGBB for opaque colors), so the reader deliberately rewrites
   * alpha 00 to FF on parse. Rgb(0x00RRGGBB) and Rgb(0xFFRRGGBB) denote the same color.
   */
  private def normalizeColor(color: Color): Color = color match
    case Color.Rgb(argb) if (argb >>> 24) == 0 => Color.Rgb(argb | 0xff000000)
    case other => other

  private def normalizeSide(side: BorderSide): BorderSide =
    side.copy(color = side.color.map(normalizeColor))

  /** Apply color canonicalization across every color slot of a style. */
  private def normalizeStyle(style: CellStyle): CellStyle =
    style.copy(
      font = style.font.copy(color = style.font.color.map(normalizeColor)),
      fill = style.fill match
        case Fill.None => Fill.None
        case Fill.Solid(c) => Fill.Solid(normalizeColor(c))
        case Fill.Pattern(fg, bg, p) => Fill.Pattern(normalizeColor(fg), normalizeColor(bg), p),
      border = Border(
        left = normalizeSide(style.border.left),
        right = normalizeSide(style.border.right),
        top = normalizeSide(style.border.top),
        bottom = normalizeSide(style.border.bottom)
      ),
      // numFmtId is the raw OOXML id slot (filled by the reader, ignored by canonicalKey) — not
      // an authored property
      numFmtId = None
    )

  /** Resolve a cell's style through its own sheet registry; dangling/absent ids are default. */
  private def resolvedStyle(sheet: Sheet, cell: Cell): CellStyle =
    normalizeStyle(
      cell.styleId
        .flatMap(sheet.styleRegistry.get)
        .getOrElse(CellStyle.default)
    )

  private def styleDiff(
    sheet: String,
    ref: ARef,
    expectedSheet: Sheet,
    expectedCell: Cell,
    actualSheet: Sheet,
    actualCell: Cell
  ): Option[String] =
    val exp = resolvedStyle(expectedSheet, expectedCell)
    val act = resolvedStyle(actualSheet, actualCell)
    Option.when(exp != act)(
      s"$sheet!${ref.toA1}: resolved style mismatch:\n  expected $exp\n  actual   $act"
    )

  private def hyperlinkDiff(
    sheet: String,
    ref: ARef,
    expected: Cell,
    actual: Cell
  ): Option[String] =
    Option.when(expected.hyperlink != actual.hyperlink)(
      s"$sheet!${ref.toA1}: hyperlink mismatch: expected ${expected.hyperlink}, actual ${actual.hyperlink}"
    )

  private def commentsDiff(sheet: String, expected: Sheet, actual: Sheet): List[String] =
    val refDiffs =
      keySetDiff(sheet, "comment refs", expected.comments.keySet, actual.comments.keySet)
    val bodyDiffs =
      expected.comments.keySet
        .intersect(actual.comments.keySet)
        .toList
        .sortBy(_.toA1)
        .flatMap { ref =>
          val exp = expected.comments(ref)
          val act = actual.comments(ref)
          val textMismatch = Option.when(exp.text.toPlainText != act.text.toPlainText)(
            s"$sheet!${ref.toA1}: comment text mismatch: expected '${exp.text.toPlainText}', " +
              s"actual '${act.text.toPlainText}'"
          )
          val authorMismatch = Option.when(exp.author != act.author)(
            s"$sheet!${ref.toA1}: comment author mismatch: expected ${exp.author}, actual ${act.author}"
          )
          textMismatch.toList ++ authorMismatch.toList
        }
    refDiffs ++ bodyDiffs

  private def render(items: Iterable[String]): String =
    items.toList.sorted.mkString("[", ", ", "]")
