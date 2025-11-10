package com.tjclp.xl

import org.scalacheck.{Arbitrary, Gen}
import java.time.LocalDateTime

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

  /** Generate valid sheet name */
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

  /** Generate cell error */
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

  /** Generate sheet with small number of cells */
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

  /** Generate workbook metadata */
  val genWorkbookMetadata: Gen[WorkbookMetadata] =
    for
      creator <- Gen.option(Gen.alphaNumStr)
      app <- Gen.option(Gen.alphaNumStr)
    yield WorkbookMetadata(
      creator = creator,
      application = app
    )

  /** Generate workbook with 1-3 sheets */
  val genWorkbook: Gen[Workbook] =
    for
      numSheets <- Gen.choose(1, 3)
      sheets <- Gen.listOfN(numSheets, genSheet)
      // Make sheet names unique
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
