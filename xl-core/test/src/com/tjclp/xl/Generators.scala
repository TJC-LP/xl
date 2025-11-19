package com.tjclp.xl

import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.cells.{Cell, CellError, CellValue}
import org.scalacheck.{Arbitrary, Gen}

import java.time.LocalDateTime
import com.tjclp.xl.styles.units.StyleId

/** ScalaCheck generators for XL types */
object Generators:

  /** Generate valid column index (0 to 16383 = A to XFD) */
  val genColumn: Gen[Column] =
    Gen.choose(0, 16383).map(Column.from0)

  /** Generate valid row index (0 to 1048575 = 1 to 1048576) */
  val genRow: Gen[Row] =
    Gen.choose(0, 1048575).map(Row.from0)

  /** Generate cell reference */
  val genARef: Gen[ARef] =
    for
      col <- genColumn
      row <- genRow
    yield ARef(col, row)

  /** Generate small cell reference (for testing ranges) */
  val genSmallARef: Gen[ARef] =
    for
      col <- Gen.choose(0, 25).map(Column.from0)
      row <- Gen.choose(0, 99).map(Row.from0)
    yield ARef(col, row)

  /** Generate valid sheets name */
  val genSheetName: Gen[SheetName] =
    Gen.identifier
      .map(s => s.take(31).filter(c => !Set(':', '\\', '/', '?', '*', '[', ']').contains(c)))
      .filter(_.nonEmpty)
      .map(SheetName.unsafe)

  /** Generate cell range */
  val genCellRange: Gen[CellRange] =
    for
      ref1 <- genSmallARef
      ref2 <- genSmallARef
    yield CellRange(ref1, ref2)

  /** Generate RefType (all variants) */
  val genRefType: Gen[com.tjclp.xl.addressing.RefType] =
    import com.tjclp.xl.addressing.RefType
    Gen.oneOf(
      genARef.map(RefType.Cell.apply),
      genCellRange.map(RefType.Range.apply),
      for
        sheet <- genSheetName
        ref <- genARef
      yield RefType.QualifiedCell(sheet, ref),
      for
        sheet <- genSheetName
        range <- genCellRange
      yield RefType.QualifiedRange(sheet, range)
    )

  /** Generate cell errors */
  val genCellError: Gen[CellError] =
    Gen.oneOf(
      CellError.Div0,
      CellError.NA,
      CellError.Name,
      CellError.Null,
      CellError.Num,
      CellError.Ref,
      CellError.Value
    )

  /** Generate cell value */
  val genCellValue: Gen[CellValue] =
    Gen.oneOf(
      Gen.alphaNumStr.map(CellValue.Text.apply),
      Gen.choose(-1000.0, 1000.0).map(d => CellValue.Number(BigDecimal(d))),
      Gen.oneOf(true, false).map(CellValue.Bool.apply),
      Gen.const(CellValue.Empty),
      genCellError.map(CellValue.Error.apply)
    )

  /** Generate cell */
  val genCell: Gen[Cell] =
    for
      ref <- genARef
      value <- genCellValue
      styleId <- Gen.option(Gen.choose(0, 100).map(StyleId.apply))
      comment <- Gen.option(Gen.alphaNumStr)
      hyperlink <- Gen.option(Gen.alphaNumStr.map(s => s"https://example.com/$s"))
    yield Cell(ref, value, styleId, comment, hyperlink)

  /** Generate column properties */
  val genColumnProperties: Gen[ColumnProperties] =
    for
      width <- Gen.option(Gen.choose(1.0, 255.0))
      hidden <- Gen.oneOf(true, false)
      styleId <- Gen.option(Gen.choose(0, 100).map(StyleId.apply))
    yield ColumnProperties(width, hidden, styleId)

  /** Generate row properties */
  val genRowProperties: Gen[RowProperties] =
    for
      height <- Gen.option(Gen.choose(1.0, 409.0))
      hidden <- Gen.oneOf(true, false)
      styleId <- Gen.option(Gen.choose(0, 100).map(StyleId.apply))
    yield RowProperties(height, hidden, styleId)

  /** Generate sheets with small number of cells */
  val genSheet: Gen[Sheet] =
    for
      name <- genSheetName
      numCells <- Gen.choose(0, 20)
      cells <- Gen.listOfN(numCells, genCell)
      numMerged <- Gen.choose(0, 3)
      merged <- Gen.listOfN(numMerged, genCellRange)
    yield
      Sheet(
        name = name,
        cells = cells.map(c => c.ref -> c).toMap,
        mergedRanges = merged.toSet
      )

  /** Generate workbooks metadata */
  val genWorkbookMetadata: Gen[WorkbookMetadata] =
    for
      creator <- Gen.option(Gen.alphaNumStr)
      app <- Gen.option(Gen.alphaNumStr)
    yield WorkbookMetadata(
      creator = creator,
      application = app
    )

  /** Generate workbooks with 1-3 sheets */
  val genWorkbook: Gen[Workbook] =
    for
      numSheets <- Gen.choose(1, 3)
      sheets <- Gen.listOfN(numSheets, genSheet)
      // Make sheets names unique
      uniqueSheets = sheets.zipWithIndex.map { case (sheet, i) =>
        sheet.copy(name = SheetName.unsafe(s"Sheet${i + 1}"))
      }
      metadata <- genWorkbookMetadata
      activeIndex <- Gen.choose(0, uniqueSheets.size - 1)
    yield Workbook(
      sheets = uniqueSheets.toVector,
      metadata = metadata,
      activeSheetIndex = activeIndex
    )

  // Arbitrary instances for property-based testing

  given Arbitrary[Column] = Arbitrary(genColumn)
  given Arbitrary[Row] = Arbitrary(genRow)
  given Arbitrary[ARef] = Arbitrary(genARef)
  given Arbitrary[SheetName] = Arbitrary(genSheetName)
  given Arbitrary[CellRange] = Arbitrary(genCellRange)
  given Arbitrary[CellError] = Arbitrary(genCellError)
  given Arbitrary[CellValue] = Arbitrary(genCellValue)
  given Arbitrary[Cell] = Arbitrary(genCell)
  given Arbitrary[ColumnProperties] = Arbitrary(genColumnProperties)
  given Arbitrary[RowProperties] = Arbitrary(genRowProperties)
  given Arbitrary[Sheet] = Arbitrary(genSheet)
  given Arbitrary[WorkbookMetadata] = Arbitrary(genWorkbookMetadata)
  given Arbitrary[Workbook] = Arbitrary(genWorkbook)

  // ===== String generators for runtime parsing tests =====

  /** Valid money strings */
  val genMoneyString: Gen[String] = Gen.oneOf(
    Gen.posNum[Double].map(n => f"$$$n%.2f"),
    Gen.posNum[Int].map(n => f"$$$n%,d.00"),
    Gen.const("$1,234.56"),
    Gen.const("999.99")
  )

  /** Valid percent strings */
  val genPercentString: Gen[String] =
    Gen.choose(0.0, 100.0).map(n => f"$n%.2f%%")

  /** Valid date strings (ISO format) */
  val genDateString: Gen[String] =
    Gen.choose(2000, 2030).flatMap { year =>
      Gen.choose(1, 12).flatMap { month =>
        Gen.choose(1, 28).map { day =>
          f"$year%04d-$month%02d-$day%02d"
        }
      }
    }

  /** Valid formula strings */
  val genFormulaString: Gen[String] = Gen.oneOf(
    Gen.const("=SUM(A1:A10)"),
    Gen.const("=IF(A1>0,B1,C1)"),
    Gen.const("=AVERAGE(B2:B100)"),
    Gen.const("=COUNT(C1:C50)")
  )

  // Invalid strings for negative tests
  val genInvalidMoney: Gen[String] = Gen.oneOf("$ABC", "1.2.3", "$$$$", "")
  val genInvalidPercent: Gen[String] = Gen.oneOf("ABC%", "1%%", "%", "")
  val genInvalidDate: Gen[String] = Gen.oneOf("2025-13-01", "not-a-date", "2025/11/10", "")

  /** Generate ModificationTracker for property-based testing */
  val genModificationTracker: Gen[ModificationTracker] =
    for
      modifiedSheets <- Gen.containerOf[Set, Int](Gen.choose(0, 20))
      deletedSheets <- Gen.containerOf[Set, Int](Gen.choose(0, 20))
      reorderedSheets <- Gen.oneOf(true, false)
      modifiedMetadata <- Gen.oneOf(true, false)
    yield ModificationTracker(modifiedSheets, deletedSheets, reorderedSheets, modifiedMetadata)

  given Arbitrary[ModificationTracker] = Arbitrary(genModificationTracker)
