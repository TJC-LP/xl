package com.tjclp.xl.ops

import com.tjclp.xl.Generators
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.cells.{CellValue, Comment}
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.sheets.styleSyntax.withCellStyle
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.styles.units.StyleId
import com.tjclp.xl.workbooks.Workbook
import org.scalacheck.Gen

/**
 * Generators for the algebra laws: a fixed two-sheet workbook (`Data`, `Other`) with content in the
 * A1:H12 grid, a `genEdit` that reaches EVERY `Edit` case, and a `genPatch` over the `Patch`
 * kernel. Targets stay inside the grid so most edits apply; the workbook-level cases mix hits and
 * misses so the fail-fast paths are exercised too.
 */
object EditGenerators:

  val data: SheetName = SheetName.unsafe("Data")
  val other: SheetName = SheetName.unsafe("Other")
  val missing: SheetName = SheetName.unsafe("Nope")

  private def cell(col: Int, row: Int): ARef = ARef.from0(col, row)

  /** A deterministic base sheet: numbers, text, formulas, a styled cell, a comment, a merge. */
  val baseData: Sheet =
    Sheet(data)
      .put(cell(0, 0), CellValue.Text("Item"))
      .put(cell(1, 0), CellValue.Text("Qty"))
      .put(cell(2, 0), CellValue.Text("Price"))
      .put(cell(0, 1), CellValue.Text("apple"))
      .put(cell(1, 1), CellValue.Number(BigDecimal(3)))
      .put(cell(2, 1), CellValue.Formula("B2*2", Some(CellValue.Number(BigDecimal(6)))))
      .put(cell(0, 2), CellValue.Text("pear"))
      .put(cell(1, 2), CellValue.Number(BigDecimal(5)))
      .put(cell(2, 2), CellValue.Formula("B3*2", None))
      .put(cell(0, 3), CellValue.Text("fig"))
      .put(cell(1, 3), CellValue.Number(BigDecimal(1)))
      .put(cell(3, 5), CellValue.Bool(true))
      .withCellStyle(cell(0, 0), CellStyle.default.withNumFmt(NumFmt.Currency))
      .comment(cell(0, 1), Comment.plainText("first fruit", Some("gen")))
      .merge(CellRange(cell(4, 4), cell(5, 5)))

  val baseOther: Sheet =
    Sheet(other)
      .put(cell(0, 0), CellValue.Number(BigDecimal(42)))
      .put(cell(1, 0), CellValue.Formula("Data!B2+A1", None))

  /** The base workbook every law runs on, with one workbook-scoped defined name. */
  val baseWorkbook: Workbook =
    Workbook(baseData, baseOther).withDefinedName("Total", "Data!$B$2:$B$4")

  val genSheetOpt: Gen[Option[SheetName]] =
    Gen.frequency(3 -> Gen.const(None), 3 -> Gen.const(Some(data)), 1 -> Gen.const(Some(other)))

  val genGridRef: Gen[ARef] = Generators.genGridRef

  /** Grid ranges, one in four inside E4:G7 so the branches over the fixture's E5:F6 merge fire. */
  val genGridRange: Gen[CellRange] =
    Gen.frequency(
      3 -> Generators.genGridRange,
      1 -> (for
        c1 <- Gen.choose(4, 6)
        r1 <- Gen.choose(3, 6)
        c2 <- Gen.choose(4, 6)
        r2 <- Gen.choose(3, 6)
      yield CellRange(cell(c1, r1), cell(c2, r2)))
    )
  val genLoc: Gen[Loc] = for s <- genSheetOpt; r <- genGridRef yield Loc(s, r)
  val genArea: Gen[Area] = for s <- genSheetOpt; r <- genGridRange yield Area(s, r)

  val genHint: Gen[Option[FormatHint]] =
    Gen.option(
      Gen.oneOf(
        Generators.genNumFmt.map(FormatHint.Inferred.apply),
        Generators.genNumFmt.map(FormatHint.Explicit.apply)
      )
    )

  val genValue: Gen[CellValue] = Generators.genCellValue

  val genFormulaText: Gen[String] =
    Gen.oneOf("=A1+1", "SUM(A1:B2)", "=$A$1*2", "=Data!B2", "=B2", "=IF(A1>1,\"y\",\"n\")")

  val genColSpan: Gen[ColSpan] =
    for a <- Gen.choose(0, 7); b <- Gen.choose(0, 7)
    yield ColSpan(Column.from0(a), Column.from0(b))

  val genRowSpan: Gen[RowSpan] =
    for a <- Gen.choose(0, 11); b <- Gen.choose(0, 11)
    yield RowSpan(Row.from0(a), Row.from0(b))

  val genOverlay: Gen[StyleOverlay] =
    for
      full <- Generators.genCellStyle
      mask <- Gen.listOfN(17, Gen.oneOf(true, false))
    yield
      val o = StyleOverlay.of(full)
      StyleOverlay(
        fontName = o.fontName.filter(_ => mask(0)),
        fontSize = o.fontSize.filter(_ => mask(1)),
        bold = o.bold.filter(_ => mask(2)),
        italic = o.italic.filter(_ => mask(3)),
        underline = o.underline.filter(_ => mask(4)),
        fontColor = o.fontColor.filter(_ => mask(5)),
        fill = o.fill.filter(_ => mask(6)),
        numFmt = o.numFmt.filter(_ => mask(7)),
        hAlign = o.hAlign.filter(_ => mask(8)),
        vAlign = o.vAlign.filter(_ => mask(9)),
        wrap = o.wrap.filter(_ => mask(10)),
        indent = o.indent.filter(_ => mask(11)),
        textRotation = o.textRotation.filter(_ => mask(12)),
        borderTop = o.borderTop.filter(_ => mask(13)),
        borderRight = o.borderRight.filter(_ => mask(14)),
        borderBottom = o.borderBottom.filter(_ => mask(15)),
        borderLeft = o.borderLeft.filter(_ => mask(16))
      )

  private val genSortKey: Gen[Edit.SortKeySpec] =
    for
      col <- Gen.choose(0, 7).map(Column.from0)
      dir <- Gen.oneOf(Edit.SortDir.Ascending, Edit.SortDir.Descending)
      mode <- Gen.oneOf(Edit.SortMode.Alphanumeric, Edit.SortMode.Numeric)
    yield Edit.SortKeySpec(col, dir, mode)

  private val genClearWhat: Gen[ClearWhat] =
    Gen.oneOf(
      ClearWhat.contents,
      ClearWhat.styles,
      ClearWhat.comments,
      ClearWhat.all,
      ClearWhat(contents = false, styles = true, comments = true)
    )

  private val genExistingSheet: Gen[SheetName] = Gen.frequency(4 -> data, 3 -> other, 1 -> missing)
  private val genNewSheet: Gen[SheetName] =
    Gen.frequency(4 -> SheetName.unsafe("New"), 1 -> data, 1 -> SheetName.unsafe("data"))

  private def small(n: Int): Gen[Int] = Gen.choose(1, n)

  /** One generator per `Edit` case, in declaration order. `EditLawsSpec` pins the coverage. */
  val caseGenerators: Vector[Gen[Edit]] = Vector(
    for at <- genLoc; v <- genValue; h <- genHint yield Edit.Put(at, v, h),
    for
      at <- genArea
      vs <- Gen.listOfN(at.range.cellCount.toInt, genValue)
      h <- genHint
    yield Edit.PutValues(at, vs.toVector, h),
    for at <- genLoc; f <- genFormulaText; h <- genHint yield Edit.PutFormula(at, f, h),
    for
      at <- genArea
      fs <- Gen.listOfN(at.range.cellCount.toInt, genFormulaText)
      h <- genHint
    yield Edit.PutFormulas(at, fs.toVector, h),
    for at <- genArea; f <- genFormulaText; a <- genGridRef; h <- genHint
    yield Edit.DragFormula(at, f, a, h),
    for
      s <- genSheetOpt
      col <- Gen.choose(0, 5)
      srcRows <- Gen.choose(0, 3)
      width <- Gen.choose(0, 2)
      dir <- Gen.oneOf(Edit.FillDir.Down, Edit.FillDir.Right)
      len <- Gen.choose(1, 6)
    yield
      val source = CellRange(cell(col, srcRows), cell(col + width, srcRows + 1))
      val target = dir match
        case Edit.FillDir.Down =>
          CellRange(cell(col, srcRows + 2), cell(col + width, srcRows + 2 + len))
        case Edit.FillDir.Right =>
          CellRange(cell(col + width + 1, srcRows), cell(col + width + 1 + len, srcRows + 1))
      Edit.Fill(Area(s, source), target, dir)
    ,
    for src <- genArea; tgt <- genLoc; v <- Gen.oneOf(true, false) yield Edit.Copy(src, tgt, v),
    for
      at <- genArea
      keys <- Gen.nonEmptyListOf(genSortKey).map(_.take(2).toVector)
      h <- Gen.oneOf(true, false)
    yield Edit.Sort(at, keys, h),
    for at <- genArea; w <- genClearWhat yield Edit.Clear(at, w),
    for at <- genArea; o <- genOverlay; m <- Gen.oneOf(StyleMode.Merge, StyleMode.Replace)
    yield Edit.Style(at, o, m),
    genArea.map(Edit.Merge.apply),
    genArea.map(Edit.Unmerge.apply),
    for s <- genSheetOpt; c <- genColSpan; w <- Gen.choose(1.0, 60.0) yield Edit.ColWidth(s, c, w),
    for s <- genSheetOpt; r <- genRowSpan; h <- Gen.choose(5.0, 120.0)
    yield Edit.RowHeight(s, r, h),
    for s <- genSheetOpt; c <- genColSpan yield Edit.HideCols(s, c),
    for s <- genSheetOpt; c <- genColSpan yield Edit.ShowCols(s, c),
    for s <- genSheetOpt; r <- genRowSpan yield Edit.HideRows(s, r),
    for s <- genSheetOpt; r <- genRowSpan yield Edit.ShowRows(s, r),
    for s <- genSheetOpt; r <- genRowSpan; l <- small(7); c <- Gen.oneOf(true, false)
    yield Edit.GroupRows(s, r, l, c),
    for s <- genSheetOpt; c <- genColSpan; l <- small(7); k <- Gen.oneOf(true, false)
    yield Edit.GroupCols(s, c, l, k),
    for s <- genSheetOpt; r <- genRowSpan yield Edit.UngroupRows(s, r),
    for s <- genSheetOpt; c <- genColSpan yield Edit.UngroupCols(s, c),
    for s <- genSheetOpt; c <- Gen.option(genColSpan) yield Edit.AutoFit(s, c),
    for at <- genLoc; c <- Generators.genComment yield Edit.SetComment(at, c),
    genLoc.map(Edit.RemoveComment.apply),
    for at <- genLoc; t <- Gen.option(Generators.genHyperlink) yield Edit.Hyperlink(at, t),
    for
      s <- genSheetOpt
      ranges <- Gen.nonEmptyListOf(genGridRange).map(_.take(2).toVector)
      rules <- Gen.nonEmptyListOf(Generators.genCfRule).map(_.take(2).toVector)
    yield Edit.AddConditionalFormat(s, ranges, rules),
    for s <- genSheetOpt; c <- Generators.genChart(data); a <- Generators.genDrawingAnchor
    yield Edit.AddChart(s, c, a),
    for s <- genSheetOpt; i <- Generators.genImageData; a <- Generators.genDrawingAnchor
    yield Edit.AddImage(s, i, a),
    for s <- genSheetOpt; c <- Gen.choose(0, 7); r <- Gen.choose(0, 11)
    yield Edit.Freeze(Loc(s, cell(c, r))),
    genSheetOpt.map(Edit.Unfreeze.apply),
    for
      s <- genSheetOpt
      g <- Gen.option(Gen.oneOf(true, false))
      z <- Gen.option(Gen.oneOf(10, 85, 100, 400, 5, 401))
      t <- Gen.option(Gen.oneOf(true, false))
    yield Edit.SetSheetView(s, g, z, t),
    for s <- genSheetOpt; c <- Gen.option(Generators.genColor) yield Edit.SetTabColor(s, c),
    for s <- genSheetOpt; r <- Gen.option(genGridRange) yield Edit.SetAutoFilter(s, r),
    for
      s <- genSheetOpt
      o <- Gen.option(Gen.oneOf("portrait", "landscape", "sideways"))
      sc <- Gen.option(Gen.oneOf(10, 75, 100, 400, 5))
      fw <- Gen.option(Gen.oneOf(0, 1, 2, -1))
      fh <- Gen.option(Gen.oneOf(0, 1, 3))
      fp <- Gen.option(Gen.oneOf(true, false))
    yield Edit.SetPageSetup(s, o, sc, fw, fh, fp),
    for
      s <- genSheetOpt
      oh <- Gen.option(Gen.const("&LLeft&RRight"))
      of <- Gen.option(Gen.const("Page &P of &N"))
      eh <- Gen.option(Gen.const("even"))
      ef <- Gen.option(Gen.const("even foot"))
      fh <- Gen.option(Gen.const("first"))
      ff <- Gen.option(Gen.const("first foot"))
      doe <- Gen.oneOf(true, false)
      df <- Gen.oneOf(true, false)
    yield Edit.SetHeaderFooter(s, oh, of, eh, ef, fh, ff, doe, df),
    for s <- genSheetOpt; at <- Gen.choose(0, 11).map(Row.from0); n <- small(3)
    yield Edit.InsertRows(s, at, n),
    for s <- genSheetOpt; at <- Gen.choose(0, 11).map(Row.from0); n <- small(3)
    yield Edit.DeleteRows(s, at, n),
    for s <- genSheetOpt; at <- Gen.choose(0, 7).map(Column.from0); n <- small(3)
    yield Edit.InsertCols(s, at, n),
    for s <- genSheetOpt; at <- Gen.choose(0, 7).map(Column.from0); n <- small(3)
    yield Edit.DeleteCols(s, at, n),
    for
      n <- genNewSheet
      after <- Gen.option(genExistingSheet)
      before <- Gen.frequency(3 -> Gen.const(None), 1 -> genExistingSheet.map(Some(_)))
    yield Edit.AddSheet(n, after, before),
    genExistingSheet.map(Edit.RemoveSheet.apply),
    for from <- genExistingSheet; to <- Gen.oneOf(SheetName.unsafe("Renamed"), other)
    yield Edit.RenameSheet(from, to),
    for
      n <- genExistingSheet
      choice <- Gen.choose(0, 2)
      idx <- Gen.choose(0, 2)
      rel <- genExistingSheet
    yield choice match
      case 0 => Edit.MoveSheet(n, Some(idx), None, None)
      case 1 => Edit.MoveSheet(n, None, Some(rel), None)
      case _ => Edit.MoveSheet(n, None, None, Some(rel)),
    for src <- genExistingSheet; tgt <- genNewSheet yield Edit.CopySheet(src, tgt),
    for n <- genExistingSheet; v <- Gen.oneOf(true, false) yield Edit.HideSheet(n, v),
    genExistingSheet.map(Edit.ShowSheet.apply),
    for
      n <- Gen.oneOf("Total", "Rate")
      r <- Gen.oneOf("Data!$A$1", "0.08", "Other!$A$1:$B$1")
      s <- Gen.frequency(2 -> Gen.const(None), 1 -> genExistingSheet.map(Some(_)))
    yield Edit.DefineName(n, r, s),
    for
      n <- Gen.oneOf("Total", "Rate")
      s <- Gen.frequency(2 -> Gen.const(None), 1 -> genExistingSheet.map(Some(_)))
    yield Edit.RemoveName(n, s)
  )

  /** Every case, uniformly. */
  val genEdit: Gen[Edit] = Gen.choose(0, caseGenerators.size - 1).flatMap(caseGenerators(_))

  /** The cases that stay on one sheet and never touch workbook structure (for the sheet laws). */
  val genSheetEdit: Gen[Edit] =
    genEdit.retryUntil(e => EditSchema.specOf(e).sheetScoped && !EditSchema.specOf(e).structural)

  val genEdits: Gen[Vector[Edit]] =
    Gen.choose(0, 4).flatMap(n => Gen.listOfN(n, genEdit)).map(_.toVector)

  // ----- Patch kernel -----

  private def genPatchLeaf(sheet: Sheet): Gen[Patch] =
    val registered: Gen[StyleId] =
      Gen.choose(0, math.max(0, sheet.styleRegistry.size - 1)).map(StyleId.apply)
    Gen.oneOf(
      for r <- genGridRef; v <- genValue yield Patch.Put(r, v),
      for r <- genGridRef; id <- registered yield Patch.SetStyle(r, id),
      for r <- genGridRef; s <- Generators.genCellStyle yield Patch.SetCellStyle(r, s),
      for rg <- genGridRange; s <- Generators.genCellStyle yield Patch.SetRangeStyle(rg, s),
      for r <- genGridRef; b <- Generators.genBorder yield Patch.MergeBorder(r, b),
      genGridRef.map(Patch.ClearStyle.apply),
      genGridRange.map(Patch.Merge.apply),
      genGridRange.map(Patch.Unmerge.apply),
      genGridRef.map(Patch.Remove.apply),
      genGridRange.map(Patch.RemoveRange.apply),
      for
        origin <- genGridRef
        rows <- Gen.choose(0, 3)
        cols <- Gen.choose(1, 3)
        values <- Gen.listOfN(rows, Gen.listOfN(cols, genValue).map(_.toVector)).map(_.toVector)
      yield Patch.PutArray(origin, values),
      for r <- genGridRef; c <- Generators.genComment yield Patch.SetComment(r, c),
      for
        ranges <- Gen.listOf(genGridRange).map(_.take(2).toVector)
        rules <- Gen.listOf(Generators.genCfRule).map(_.take(2).toVector)
      yield Patch.SetConditionalFormat(ranges, rules),
      for c <- Gen.choose(0, 7).map(Column.from0); p <- Generators.genColumnProperties
      yield Patch.SetColumnProperties(c, p),
      for r <- Gen.choose(0, 11).map(Row.from0); p <- Generators.genRowProperties
      yield Patch.SetRowProperties(r, p)
    )

  /** Leaves and small batches over the base sheet's registry. */
  def genPatch(sheet: Sheet): Gen[Patch] =
    Gen.frequency(
      4 -> genPatchLeaf(sheet),
      1 -> Gen
        .choose(0, 3)
        .flatMap(n => Gen.listOfN(n, genPatchLeaf(sheet)))
        .map(ps => Patch.Batch(ps.toVector))
    )
