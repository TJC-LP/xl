package com.tjclp.xl

import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.cells.{Cell, CellError, CellValue, Comment}
import com.tjclp.xl.cf.{CfOperator, CfPoint, CfRule, CfTextOp, Cfvo, ConditionalFormat}
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.context.ModificationTracker
import com.tjclp.xl.drawings.TestImages
import com.tjclp.xl.sheets.{FreezePane, HeaderFooter, PageMargins, PageSetup, SheetView}
import com.tjclp.xl.sheets.styleSyntax.*
import com.tjclp.xl.styles.{CellStyle, Dxf, DxfFont}
import com.tjclp.xl.styles.alignment.{Align, HAlign, VAlign}
import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.styles.color.{Color, ThemeSlot}
import com.tjclp.xl.styles.fill.{Fill, PatternType}
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt
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
      outlineLevel <- Gen.option(Gen.choose(0, 7))
      collapsed <- Gen.oneOf(true, false)
    yield ColumnProperties(width, hidden, styleId, outlineLevel, collapsed)

  /** Generate row properties */
  val genRowProperties: Gen[RowProperties] =
    for
      height <- Gen.option(Gen.choose(1.0, 409.0))
      hidden <- Gen.oneOf(true, false)
      styleId <- Gen.option(Gen.choose(0, 100).map(StyleId.apply))
      outlineLevel <- Gen.option(Gen.choose(0, 7))
      collapsed <- Gen.oneOf(true, false)
    yield RowProperties(height, hidden, styleId, outlineLevel, collapsed)

  /** Generate sheet with small number of cells */
  val genSheet: Gen[Sheet] =
    for
      name <- genSheetName
      numCells <- Gen.choose(0, 20)
      cells <- Gen.listOfN(numCells, genCell)
      numMerged <- Gen.choose(0, 3)
      merged <- Gen.listOfN(numMerged, genCellRange)
    yield Sheet(
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

  // =====================================================================
  // Rich generators for the OOXML generative round-trip law (GH-240a).
  //
  // These produce workbooks that exercise styles, comments, hyperlinks,
  // formulas, merges, sheet views, and page setup, while staying inside the
  // domain the OOXML layer can faithfully round-trip:
  //   - text avoids XML-illegal control chars (the writer strips them by
  //     design, GH-237); \r IS generated — the writer escapes it as _x000D_
  //     per ECMA-376 ST_Xstring (GH-288)
  //   - non-NFC text (decomposed accents) IS generated — the SST deduplicates
  //     exact strings, so NFC/NFD spellings round-trip byte-faithfully (GH-289)
  //   - Custom numFmt codes avoid the exact code strings NumFmt.parse maps
  //     back to built-in enum cases (those are the SAME format semantically,
  //     but would compare unequal as enum values)
  //   - degenerate style states the writer cannot represent are avoided:
  //     BorderSide(None, Some(color)) drops its color, Fill.Pattern with
  //     pattern None/Solid collapses to Fill.None/Fill.Solid
  //   - generated PageSetup/HeaderFooter always carry at least one visible
  //     (non-default) field; an all-default PageSetup serializes to nothing
  //     and reads back as None by design
  // =====================================================================

  /** Cell text that round-trips through OOXML XML (see constraints above) */
  val genXmlSafeText: Gen[String] =
    val safeChar: Gen[Char] = Gen.frequency(
      10 -> Gen.alphaNumChar,
      3 -> Gen.const(' '),
      2 -> Gen.oneOf('.', ',', '-', '_', '&', '<', '>', '"', '\'', '%', '$', '#', '(', ')', '/'),
      1 -> Gen.oneOf('é', 'ü', 'ß', '日', '本', '€', '£'),
      1 -> Gen.oneOf('\t', '\n', '\r')
    )
    val plain = Gen.chooseNum(0, 24).flatMap(n => Gen.listOfN(n, safeChar).map(_.mkString))
    // Adversarial suffixes: literal _xHHHH_ patterns must survive via _x005F_ protection (GH-288);
    // decomposed accents (NFD "é" = e + U+0301) must keep their exact codepoints (GH-289)
    Gen.frequency(
      14 -> plain,
      1 -> plain.map(_ + "_x000D_"),
      1 -> plain.map(_ + "\r\n"),
      1 -> plain.map(_ + "e\u0301")
    )

  /** Non-empty variant of [[genXmlSafeText]] */
  val genXmlSafeTextNonEmpty: Gen[String] =
    for
      head <- Gen.alphaNumChar
      tail <- genXmlSafeText
    yield s"$head$tail"

  /** Colors stable through the OOXML hex/theme round-trip */
  val genColor: Gen[Color] =
    Gen.frequency(
      6 -> Gen
        .oneOf(0xff000000, 0xffffffff, 0xffff0000, 0xff00b0f0, 0xff4472c4, 0xffffc000, 0xff70ad47)
        .map(Color.Rgb.apply),
      2 -> (for
        slot <- Gen.oneOf(ThemeSlot.values.toIndexedSeq)
        tint <- Gen.oneOf(-0.5, -0.25, 0.0, 0.25, 0.5)
      yield Color.Theme(slot, tint)),
      1 -> Gen.chooseNum(Int.MinValue, Int.MaxValue).map(Color.Rgb.apply)
    )

  /** Generate font (realistic names/sizes; all flag combinations) */
  val genFont: Gen[Font] =
    for
      name <- Gen.oneOf("Calibri", "Arial", "Times New Roman", "Courier New", "Aptos Narrow")
      size <- Gen.oneOf(8.0, 9.0, 10.0, 10.5, 11.0, 12.0, 14.0, 16.0, 22.0)
      bold <- Gen.oneOf(true, false)
      italic <- Gen.oneOf(true, false)
      underline <- Gen.oneOf(true, false)
      color <- Gen.option(genColor)
    yield Font(name, size, bold, italic, underline, color)

  /**
   * Generate fill. Fill.Pattern is only generated with pattern types other than None/Solid: the
   * dedicated enum cases Fill.None/Fill.Solid own those encodings.
   */
  val genFill: Gen[Fill] =
    val texturePatterns = PatternType.values.toIndexedSeq.filter {
      case PatternType.None | PatternType.Solid => false
      case _ => true
    }
    Gen.frequency(
      4 -> Gen.const(Fill.None),
      4 -> genColor.map(Fill.Solid.apply),
      2 -> (for
        fg <- genColor
        bg <- genColor
        pattern <- Gen.oneOf(texturePatterns)
      yield Fill.Pattern(fg, bg, pattern))
    )

  /**
   * Generate border side. A color is only attached when the style is not None — the writer cannot
   * represent a colored "no border" side (it serializes as a bare side element).
   */
  val genBorderSide: Gen[BorderSide] =
    for
      style <- Gen.oneOf(BorderStyle.values.toIndexedSeq)
      color <- if style == BorderStyle.None then Gen.const(None) else Gen.option(genColor)
    yield BorderSide(style, color)

  /** Generate border (independent sides) */
  val genBorder: Gen[Border] =
    for
      left <- genBorderSide
      right <- genBorderSide
      top <- genBorderSide
      bottom <- genBorderSide
    yield Border(left, right, top, bottom)

  /**
   * Representative number formats: every built-in enum case plus Custom codes (codes chosen to not
   * collide with the built-in code strings NumFmt.parse recognizes).
   */
  val genNumFmt: Gen[NumFmt] =
    Gen.frequency(
      6 -> Gen.oneOf(
        NumFmt.General,
        NumFmt.Integer,
        NumFmt.Decimal,
        NumFmt.ThousandsSeparator,
        NumFmt.ThousandsDecimal,
        NumFmt.Currency,
        NumFmt.Percent,
        NumFmt.PercentDecimal,
        NumFmt.Scientific,
        NumFmt.Fraction,
        NumFmt.Date,
        NumFmt.DateTime,
        NumFmt.Time,
        NumFmt.Text
      ),
      2 -> Gen
        .oneOf(
          "0.000",
          "#,##0.0",
          "yyyy-mm-dd",
          "mmm yyyy",
          "0.0%",
          "$#,##0.00;[Red]($#,##0.00)",
          "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
          "[$€-407] #,##0.00",
          "0.00 \"units\""
        )
        .map(NumFmt.Custom.apply)
    )

  /** Generate alignment incl. indent (GH-style: indent levels 0-15) */
  val genAlign: Gen[Align] =
    for
      h <- Gen.oneOf(HAlign.values.toIndexedSeq)
      v <- Gen.oneOf(VAlign.values.toIndexedSeq)
      wrap <- Gen.oneOf(true, false)
      indent <- Gen.frequency(3 -> Gen.const(0), 1 -> Gen.choose(1, 15))
    yield Align(h, v, wrap, indent)

  /** Generate complete cell style (numFmtId left writer-assigned) */
  val genCellStyle: Gen[CellStyle] =
    for
      font <- genFont
      fill <- genFill
      border <- genBorder
      numFmt <- genNumFmt
      align <- genAlign
    yield CellStyle(font, fill, border, numFmt, None, align)

  /**
   * Generate comment: plain text body plus optional author. Authors may carry edge whitespace or be
   * whitespace-only — the writer canonicalizes (trim; blank → unauthored, GH-290), and round-trip
   * equivalence compares canonical authors.
   */
  val genComment: Gen[Comment] =
    val genAuthor: Gen[String] =
      for
        base <- Gen.identifier.map(_.take(12))
        decorated <- Gen.frequency(
          6 -> Gen.const(base),
          1 -> Gen.const(s" $base "),
          1 -> Gen.const(s"$base  "),
          1 -> Gen.const("   ") // whitespace-only → canonicalizes to unauthored
        )
      yield decorated
    for
      text <- genXmlSafeTextNonEmpty
      author <- Gen.option(genAuthor)
    yield Comment.plainText(text, author)

  /** Generate hyperlink target: external URL/mailto or internal location */
  val genHyperlink: Gen[String] =
    Gen.frequency(
      5 -> Gen.identifier.map(s => s"https://example.com/${s.take(10)}"),
      2 -> Gen.identifier.map(s => s"mailto:${s.take(8)}@example.com"),
      3 -> Gen.oneOf("A1", "Sheet1!B2", "'Q1 Report'!C3")
    )

  /**
   * Formula text that round-trips verbatim (no surrounding whitespace — the reader trims the
   * formula element text). Formulas round-trip as TEXT; no evaluation happens.
   */
  val genFormulaExpr: Gen[String] =
    Gen.oneOf(
      "A1+1",
      "=A1*2",
      "SUM(A1:B2)",
      "IF(A1<5,1,0)",
      "AVERAGE(B1:B5)",
      "CONCATENATE(A1,\"x\")",
      "'Q1 Report'!A1+B2"
    )

  /** Numbers that survive plain-decimal serialization (incl. GH-238 magnitude extremes) */
  val genRoundTripNumber: Gen[BigDecimal] =
    Gen.frequency(
      5 -> Gen.chooseNum(-1000000.0, 1000000.0).map(BigDecimal.apply),
      2 -> Gen.chooseNum(-1000000L, 1000000L).map(BigDecimal.apply),
      1 -> Gen.oneOf(
        BigDecimal("1E20"),
        BigDecimal("1E-7"),
        BigDecimal("-1E18"),
        BigDecimal(0),
        BigDecimal("0.0001"),
        BigDecimal("123456789.123456789")
      )
    )

  /** DateTime within Excel's representable era, whole seconds */
  val genExcelDateTime: Gen[LocalDateTime] =
    for
      year <- Gen.choose(1950, 2099)
      month <- Gen.choose(1, 12)
      day <- Gen.choose(1, 28)
      hour <- Gen.choose(0, 23)
      minute <- Gen.choose(0, 59)
      second <- Gen.choose(0, 59)
    yield LocalDateTime.of(year, month, day, hour, minute, second)

  /** Formula cell value with optional cached value (cached values are write-only metadata) */
  val genFormulaCellValue: Gen[CellValue] =
    for
      expr <- genFormulaExpr
      cached <- Gen.option(
        Gen.oneOf(
          genRoundTripNumber.map(CellValue.Number.apply),
          Gen.oneOf(true, false).map(CellValue.Bool.apply),
          Gen.alphaNumStr.map(CellValue.Text.apply),
          genCellError.map(CellValue.Error.apply)
        )
      )
    yield CellValue.Formula(expr, cached)

  /** Cell values for round-trip testing (all OOXML-representable variants) */
  val genRichCellValue: Gen[CellValue] =
    Gen.frequency(
      4 -> genXmlSafeText.map(CellValue.Text.apply),
      3 -> genRoundTripNumber.map(CellValue.Number.apply),
      1 -> Gen.oneOf(true, false).map(CellValue.Bool.apply),
      1 -> genCellError.map(CellValue.Error.apply),
      1 -> genExcelDateTime.map(CellValue.DateTime.apply),
      2 -> genFormulaCellValue,
      1 -> Gen.const(CellValue.Empty)
    )

  /** Sheet view settings (zoom always within Excel's 10-400 bounds) */
  val genSheetView: Gen[SheetView] =
    for
      gridLines <- Gen.oneOf(true, false)
      zoom <- Gen.option(Gen.oneOf(25, 75, 85, 100, 150, 200, 400))
    yield SheetView(gridLines, zoom)

  /** Header/footer with at least one visible part or flag (all-default reads back as None) */
  val genHeaderFooter: Gen[HeaderFooter] =
    val part = Gen.oneOf(
      "&LTHE JORDAN COMPANY&RConfidential",
      "&CPage &P of &N",
      "&L&D &T",
      "Draft — &A",
      "&F"
    )
    for
      oddH <- Gen.option(part)
      oddF <- Gen.option(part)
      evenH <- Gen.option(part)
      evenF <- Gen.option(part)
      firstH <- Gen.option(part)
      firstF <- Gen.option(part)
      diffOddEven <- Gen.oneOf(true, false)
      diffFirst <- Gen.oneOf(true, false)
    yield
      val hf =
        HeaderFooter(oddH, oddF, evenH, evenF, firstH, firstF, diffOddEven, diffFirst)
      if hf == HeaderFooter() then HeaderFooter(oddHeader = Some("&CPage &P")) else hf

  /** Page margins (positive inches that round-trip through Double.toString) */
  val genPageMargins: Gen[PageMargins] =
    for
      left <- Gen.oneOf(0.25, 0.5, 0.7, 1.0)
      right <- Gen.oneOf(0.25, 0.5, 0.7, 1.0)
      top <- Gen.oneOf(0.5, 0.75, 1.0)
      bottom <- Gen.oneOf(0.5, 0.75, 1.0)
      header <- Gen.oneOf(0.25, 0.3, 0.5)
      footer <- Gen.oneOf(0.25, 0.3, 0.5)
    yield PageMargins(left, right, top, bottom, header, footer)

  /**
   * Page setup with at least one visible (non-default) field — an all-default PageSetup serializes
   * to no XML and intentionally reads back as None. printArea/repeatRows are exercised by the
   * dedicated PrintNames specs, not generated here.
   */
  val genPageSetup: Gen[PageSetup] =
    for
      scale <- Gen.frequency(3 -> Gen.const(100), 1 -> Gen.oneOf(50, 85, 120, 200))
      orientation <- Gen.option(Gen.oneOf("portrait", "landscape"))
      fitToWidth <- Gen.option(Gen.choose(0, 3))
      fitToHeight <- Gen.option(Gen.choose(0, 3))
      headerFooter <- Gen.option(genHeaderFooter)
      margins <- Gen.option(genPageMargins)
    yield
      val ps = PageSetup(scale, orientation, fitToWidth, fitToHeight, headerFooter, margins)
      if ps == PageSetup() then ps.copy(orientation = Some("landscape")) else ps

  /**
   * Freeze pane at a non-A1 cell (freezing at A1 is a no-op the writer elides). NOTE: freezePane
   * has write-only three-valued semantics (None = preserve) — the reader never populates it, so
   * generating it exercises the writer without participating in round-trip equality.
   */
  val genFreezePane: Gen[FreezePane] =
    for
      col <- Gen.choose(0, 3)
      row <- Gen.choose(0, 5)
    yield
      val ref = if col == 0 && row == 0 then ARef.from0(1, 1) else ARef.from0(col, row)
      FreezePane.At(ref)

  /** Cell reference within a compact grid (A1:H12) so merges/comments cluster realistically */
  val genGridRef: Gen[ARef] =
    for
      col <- Gen.choose(0, 7)
      row <- Gen.choose(0, 11)
    yield ARef.from0(col, row)

  /** Cell range within the compact grid */
  val genGridRange: Gen[CellRange] =
    for
      ref1 <- genGridRef
      ref2 <- genGridRef
    yield CellRange(ref1, ref2)

  // ===== Drawing generators (GH-221) =====

  /** Image payload drawn from fixed, valid byte templates (png/gif/jpeg/bmp, all 2x3 px). */
  val genImageData: Gen[ImageData] =
    Gen.oneOf(
      ImageData(TestImages.png2x3, ImageFormat.Png),
      ImageData(TestImages.gif2x3, ImageFormat.Gif),
      ImageData(TestImages.jpeg2x3, ImageFormat.Jpeg),
      ImageData(TestImages.bmp2x3, ImageFormat.Bmp)
    )

  /** EMU offset inside a cell (small, non-negative). */
  private val genEmuOffset: Gen[Emu] =
    Gen.oneOf(0L, 9525L, 95250L).map(Emu.apply)

  /** Drawing extent between ~1px and ~100px square-ish. */
  private val genExtent: Gen[Extent] =
    for
      cx <- Gen.choose(1L, 100L).map(_ * 9525L)
      cy <- Gen.choose(1L, 100L).map(_ * 9525L)
    yield Extent(Emu(cx), Emu(cy))

  /** All three anchor forms within the compact grid. */
  val genDrawingAnchor: Gen[DrawingAnchor] =
    val genPoint = for
      ref <- genGridRef
      dx <- genEmuOffset
      dy <- genEmuOffset
    yield AnchorPoint(ref, dx, dy)
    Gen.oneOf(
      for
        p <- genPoint
        e <- genExtent
      yield DrawingAnchor.OneCell(p, e),
      for
        range <- genGridRange
        editAs <- Gen.oneOf(EditAs.TwoCell, EditAs.OneCell, EditAs.Absolute)
      // markers built off the range corners; `over` gives the canonical one-past-end form
      yield DrawingAnchor.over(range, editAs),
      for
        x <- Gen.choose(0L, 1000L).map(v => Emu(v * 9525L))
        y <- Gen.choose(0L, 1000L).map(v => Emu(v * 9525L))
        e <- genExtent
      yield DrawingAnchor.Absolute(x, y, e)
    )

  /** Typed picture: varied anchors/formats; names sometimes empty (writer assigns a default). */
  val genPicture: Gen[Drawing.Picture] =
    for
      anchor <- genDrawingAnchor
      image <- genImageData
      name <- Gen.frequency(3 -> Gen.alphaNumStr.map(_.take(12)), 1 -> Gen.const(""))
      description <- Gen.frequency(1 -> Gen.alphaNumStr.map(_.take(20)), 2 -> Gen.const(""))
    yield Drawing.Picture(anchor, image, name, description)

  // ===== Chart generators (GH-222) =====

  /** Grid-constrained 1×N / N×1 vector range with Relative anchors (the chart-range contract). */
  val genVectorGridRange: Gen[CellRange] =
    for
      column <- Gen.oneOf(true, false)
      range <-
        if column then
          for
            col <- Gen.choose(0, 7)
            r1 <- Gen.choose(0, 11)
            r2 <- Gen.choose(0, 11)
          yield CellRange(ARef.from0(col, r1), ARef.from0(col, r2))
        else
          for
            row <- Gen.choose(0, 11)
            c1 <- Gen.choose(0, 7)
            c2 <- Gen.choose(0, 7)
          yield CellRange(ARef.from0(c1, row), ARef.from0(c2, row))
    yield range

  /** Sheet-qualified vector data range on `sheet`. */
  def genDataRef(sheet: SheetName): Gen[com.tjclp.xl.charts.DataRef] =
    genVectorGridRange.map(com.tjclp.xl.charts.DataRef(sheet, _))

  /** Chart series referencing `sheet` (self-referential — the genRichWorkbook rename remaps). */
  def genSeries(sheet: SheetName): Gen[com.tjclp.xl.charts.Series] =
    import com.tjclp.xl.charts.{Series, SeriesName}
    for
      values <- genDataRef(sheet)
      categories <- Gen.frequency(1 -> Gen.const(None), 2 -> genDataRef(sheet).map(Some.apply))
      name <- Gen.frequency(
        1 -> Gen.const(None),
        1 -> Gen.alphaNumStr.map(s => Some(SeriesName.Literal(s.take(12)))),
        1 -> genGridRef.map(ref => Some(SeriesName.FromCell(sheet, ref)))
      )
    yield Series(values, categories, name)

  /** Typed chart in the validated shape: ≥1 series, pie forced to exactly one. */
  def genChart(sheet: SheetName): Gen[com.tjclp.xl.charts.Chart] =
    import com.tjclp.xl.charts.*
    val genChartType: Gen[ChartType] = Gen
      .oneOf(
        Gen
          .zip(
            Gen.oneOf(BarDirection.Col, BarDirection.Bar),
            Gen.oneOf(BarGrouping.Clustered, BarGrouping.Stacked, BarGrouping.PercentStacked)
          )
          .map(ChartType.Bar(_, _)),
        Gen.const(ChartType.Line),
        Gen.const(ChartType.Pie)
      )
      .flatMap(identity)
    for
      chartType <- genChartType
      numSeries <- if chartType == ChartType.Pie then Gen.const(1) else Gen.choose(1, 3)
      series <- Gen.listOfN(numSeries, genSeries(sheet)).map(_.toVector)
      title <- Gen.frequency(
        1 -> Gen.const(None),
        2 -> Gen.alphaNumStr.map(s => Some(s.take(20)))
      )
      legend <- Gen.frequency(
        1 -> Gen.const(None),
        3 -> (for
          pos <- Gen.oneOf(
            LegendPosition.Right,
            LegendPosition.Left,
            LegendPosition.Top,
            LegendPosition.Bottom,
            LegendPosition.TopRight
          )
          overlay <- Gen.oneOf(true, false)
        yield Some(Legend(pos, overlay)))
      )
    yield Chart(chartType, series, title, legend)

  /** Typed chart frame on `sheet`; names sometimes empty (writer assigns "Chart {ordinal}"). */
  def genChartFrame(sheet: SheetName): Gen[Drawing.ChartFrame] =
    for
      anchor <- genDrawingAnchor
      chart <- genChart(sheet)
      name <- Gen.frequency(1 -> Gen.alphaNumStr.map(_.take(12)), 2 -> Gen.const(""))
    yield Drawing.ChartFrame(anchor, chart, name)

  // ===== Conditional-format generators (GH-136) =====
  // Constraints for the round-trip law (mirroring the style-generator notes above):
  //   - colors come from the stable pool/theme subset (an alpha-00 Rgb canonicalizes to FF on
  //     parse, which plain cf equality would flag)
  //   - dxf numFmts avoid NumFmt.Currency (no format-code string parses back to that case) and
  //     Custom codes that collide with built-in code strings
  //   - priorities are stamped concrete (1..n in document order) — the model always holds final
  //     priorities, so equivalence is plain equality with no canonicalization clause

  /** Colors stable under the OOXML color parse (alpha FF or theme+tint). */
  private val genCfColor: Gen[Color] =
    Gen.frequency(
      6 -> Gen
        .oneOf(0xff000000, 0xffffffff, 0xff9c0006, 0xffffc7ce, 0xff638ec6, 0xff00b050, 0xffffc000)
        .map(Color.Rgb.apply),
      2 -> (for
        slot <- Gen.oneOf(ThemeSlot.values.toIndexedSeq)
        tint <- Gen.oneOf(-0.25, 0.0, 0.25)
      yield Color.Theme(slot, tint))
    )

  /** Differential font: at least the generated deltas, Some(false) force-offs included. */
  val genDxfFont: Gen[DxfFont] =
    val flag = Gen.oneOf(None, Some(true), Some(false))
    for
      bold <- flag
      italic <- flag
      strike <- flag
      underline <- flag
      color <- Gen.option(genCfColor)
    yield DxfFont(bold, italic, strike, underline, color)

  /** Dxf numFmts stable through code-string round-trip (see constraints above). */
  private val genDxfNumFmt: Gen[NumFmt] =
    Gen.oneOf(
      NumFmt.General,
      NumFmt.Integer,
      NumFmt.Decimal,
      NumFmt.Percent,
      NumFmt.PercentDecimal,
      NumFmt.Date,
      NumFmt.Custom("0.000"),
      NumFmt.Custom("#,##0.0"),
      NumFmt.Custom("yyyy-mm-dd")
    )

  /** Differential format within the modeled subset (font/fill/border/numFmt deltas). */
  val genDxf: Gen[Dxf] =
    val side =
      for
        style <- Gen.oneOf(
          BorderStyle.None,
          BorderStyle.Thin,
          BorderStyle.Medium,
          BorderStyle.Double
        )
        color <- if style == BorderStyle.None then Gen.const(None) else Gen.option(genCfColor)
      yield BorderSide(style, color)
    val border = for l <- side; r <- side; t <- side; b <- side
    yield Border(l, r, t, b)
    for
      font <- Gen.option(genDxfFont)
      fill <- Gen.option(genCfColor.map(c => Fill.Solid(c): Fill))
      bdr <- Gen.frequency(3 -> Gen.const(None), 1 -> border.map(Some.apply))
      numFmt <- Gen.frequency(3 -> Gen.const(None), 1 -> genDxfNumFmt.map(Some.apply))
    yield Dxf(font, fill, bdr, numFmt)

  /** Cfvo points; formula text round-trips verbatim (never re-parsed on read). */
  val genCfvo: Gen[Cfvo] =
    Gen.frequency(
      3 -> Gen.oneOf(Cfvo.Min, Cfvo.Max),
      2 -> Gen.oneOf(BigDecimal(0), BigDecimal(10), BigDecimal("1.5")).map(Cfvo.Num.apply),
      2 -> Gen.oneOf(BigDecimal(10), BigDecimal(50), BigDecimal(90)).map(Cfvo.Percent.apply),
      1 -> Gen.oneOf(BigDecimal(25), BigDecimal(75)).map(Cfvo.Percentile.apply),
      1 -> Gen.oneOf("MAX($A$1:$A$9)", "AVERAGE($B$1:$B$9)+1").map(Cfvo.Formula.apply)
    )

  /** Color-scale point. */
  val genCfPoint: Gen[CfPoint] =
    for
      cfvo <- genCfvo
      color <- genCfColor
    yield CfPoint(cfvo, color)

  private val genCfFormula: Gen[String] =
    Gen.oneOf("100", "0", "$B$1", "AVERAGE($A$1:$A$9)", "LEN(A1)>3", "\"x\"")

  private val genCfText: Gen[String] =
    Gen.oneOf("todo", "Q1 total", "x", "say \"hi\"", "100%")

  /**
   * Typed rule families with the auto-priority sentinel (callers stamp concrete priorities).
   * Preserved rules are reader-constructed and deliberately not generated.
   */
  val genCfRule: Gen[CfRule] =
    val dxfOpt = Gen.frequency(1 -> Gen.const(None), 3 -> genDxf.map(Some.apply))
    val stop = Gen.frequency(3 -> Gen.const(false), 1 -> Gen.const(true))
    Gen.oneOf(
      for
        op <- Gen.oneOf(
          CfOperator.LessThan,
          CfOperator.LessThanOrEqual,
          CfOperator.Equal,
          CfOperator.NotEqual,
          CfOperator.GreaterThanOrEqual,
          CfOperator.GreaterThan
        )
        f <- genCfFormula
        dxf <- dxfOpt
        s <- stop
      yield CfRule.CellIs(op, f, None, dxf, CfRule.AutoPriority, s),
      for
        op <- Gen.oneOf(CfOperator.Between, CfOperator.NotBetween)
        lo <- Gen.oneOf("1", "0", "$A$1")
        hi <- Gen.oneOf("9", "100")
        dxf <- dxfOpt
        s <- stop
      yield CfRule.CellIs(op, lo, Some(hi), dxf, CfRule.AutoPriority, s),
      for
        f <- Gen.oneOf("$B1>$C1", "MOD(ROW(),2)=0", "A1>AVERAGE($A$1:$A$9)")
        dxf <- dxfOpt
        s <- stop
      yield CfRule.Expression(f, dxf, CfRule.AutoPriority, s),
      for
        min <- genCfPoint
        mid <- Gen.option(genCfPoint)
        max <- genCfPoint
      yield CfRule.ColorScale(min, mid, max, CfRule.AutoPriority),
      for
        min <- genCfvo
        max <- genCfvo
        color <- genCfColor
        show <- Gen.oneOf(true, false)
      yield CfRule.DataBar(min, max, color, show, CfRule.AutoPriority),
      for
        rank <- Gen.choose(1, 50)
        percent <- Gen.oneOf(true, false)
        bottom <- Gen.oneOf(true, false)
        dxf <- dxfOpt
        s <- stop
      yield CfRule.Top10(rank, percent, bottom, dxf, CfRule.AutoPriority, s),
      for
        op <- Gen
          .oneOf(CfTextOp.Contains, CfTextOp.NotContains, CfTextOp.BeginsWith, CfTextOp.EndsWith)
        text <- genCfText
        dxf <- dxfOpt
        s <- stop
      yield CfRule.Text(op, text, dxf, CfRule.AutoPriority, s)
    )

  /** One typed block: 1-2 in-grid ranges, 1-3 rules (priorities stamped by the caller). */
  private val genCfBlock: Gen[ConditionalFormat.Rules] =
    for
      numRanges <- Gen.choose(1, 2)
      ranges <- Gen.listOfN(numRanges, genGridRange)
      numRules <- Gen.choose(1, 3)
      rules <- Gen.listOfN(numRules, genCfRule)
    yield ConditionalFormat.Rules(ranges.toVector.distinct, rules.toVector)

  /**
   * ≤3 typed blocks with CONCRETE priorities 1..n stamped in document order — the shape
   * `Sheet.conditionalFormat` produces, so round-trip equivalence is plain equality.
   */
  val genConditionalFormats: Gen[Vector[ConditionalFormat]] =
    for
      numBlocks <- Gen.choose(1, 3)
      blocks <- Gen.listOfN(numBlocks, genCfBlock)
    yield
      val (stamped, _) = blocks.foldLeft((Vector.empty[ConditionalFormat], 0)) {
        case ((acc, n), block) =>
          val (rules, next) = block.rules.foldLeft((Vector.empty[CfRule], n)) {
            case ((rs, k), rule) => (rs :+ CfRule.withPriority(rule, k + 1), k + 1)
          }
          (acc :+ block.copy(rules = rules), next)
      }
      stamped

  /**
   * Remap a sheet's typed-chart references from `oldName` to `newName` — the genRichWorkbook rename
   * fixup (TRAP-3): generated charts self-reference their sheet, so the post-generation rename must
   * follow or every DataRef dangles and caches never get exercised.
   */
  private def remapChartSheet(sheet: Sheet, oldName: SheetName, newName: SheetName): Sheet =
    import com.tjclp.xl.charts.{DataRef, SeriesName}
    def remap(ref: DataRef): DataRef =
      if ref.sheet.value.equalsIgnoreCase(oldName.value) then ref.copy(sheet = newName) else ref
    val drawings = sheet.drawings.map {
      case Drawing.ChartFrame(anchor, chart, n) =>
        val series = chart.series.map { s =>
          s.copy(
            values = remap(s.values),
            categories = s.categories.map(remap),
            name = s.name.map {
              case SeriesName.FromCell(sn, ref) if sn.value.equalsIgnoreCase(oldName.value) =>
                SeriesName.FromCell(newName, ref)
              case other => other
            }
          )
        }
        Drawing.ChartFrame(anchor, chart.copy(series = series), n)
      case other => other
    }
    sheet.copy(drawings = drawings)

  /**
   * Rich sheet for the round-trip law: ≤30 cells with values/styles/comments/hyperlinks, merges,
   * optional view settings, page setup, and freeze panes. Styles are registered through the sheet's
   * StyleRegistry (withCellStyle) so styleIds always resolve.
   */
  val genRichSheet: Gen[Sheet] =
    val genEntry =
      for
        ref <- genGridRef
        value <- genRichCellValue
        style <- Gen.frequency(2 -> Gen.const(None), 2 -> genCellStyle.map(Some.apply))
        comment <- Gen.frequency(3 -> Gen.const(None), 1 -> genComment.map(Some.apply))
        hyperlink <- Gen.frequency(4 -> Gen.const(None), 1 -> genHyperlink.map(Some.apply))
      yield (ref, value, style, comment, hyperlink)
    for
      name <- genSheetName
      numCells <- Gen.choose(0, 30)
      entries <- Gen.listOfN(numCells, genEntry)
      numMerges <- Gen.choose(0, 3)
      merges <- Gen.listOfN(numMerges, genGridRange)
      view <- Gen.option(genSheetView)
      pageSetup <- Gen.option(genPageSetup)
      freeze <- Gen.frequency(3 -> Gen.const(None), 1 -> genFreezePane.map(Some.apply))
      // GH-221/GH-222: pictures at comment-like frequency, charts rarer still
      drawings <- Gen.frequency(
        3 -> Gen.const(Vector.empty[Drawing]),
        1 -> Gen.choose(1, 2).flatMap(n => Gen.listOfN(n, genPicture).map(_.toVector)),
        1 -> genChartFrame(name).map(Vector(_))
      )
      // GH-136: typed conditional formats at picture-like frequency
      condFmts <- Gen.frequency(
        3 -> Gen.const(Vector.empty[ConditionalFormat]),
        1 -> genConditionalFormats
      )
    yield
      val withCells = entries.foldLeft(Sheet(name)) {
        case (sheet, (ref, value, style, comment, hyperlink)) =>
          val withCell = sheet.put(Cell(ref, value, None, None, hyperlink))
          val styled = style.fold(withCell)(s => withCell.withCellStyle(ref, s))
          comment.fold(styled)(c => styled.comment(ref, c))
      }
      val withMerges = merges.foldLeft(withCells)(_.merge(_))
      withMerges.copy(
        viewSettings = view,
        pageSetup = pageSetup,
        freezePane = freeze,
        drawings = drawings,
        conditionalFormats = condFmts
      )

  /**
   * Rich workbook for the round-trip law: 1-3 rich sheets with unique, realistic names (spaces,
   * accents, apostrophes, ampersands all legal in sheet names).
   */
  val genRichWorkbook: Gen[Workbook] =
    val namePool = Vector("Data", "Q1 Report", "Summary", "Détails", "O'Brien", "P&L")
    for
      numSheets <- Gen.choose(1, 3)
      sheets <- Gen.listOfN(numSheets, genRichSheet)
      offset <- Gen.choose(0, namePool.size - 1)
      metadata <- genWorkbookMetadata
      activeIndex <- Gen.choose(0, numSheets - 1)
    yield
      val unique = sheets.zipWithIndex.map { case (sheet, i) =>
        val newName = SheetName.unsafe(s"${namePool((offset + i) % namePool.size)} ${i + 1}")
        // TRAP-3 (GH-222): chart refs are self-referential — remap them with the rename
        remapChartSheet(sheet.copy(name = newName), sheet.name, newName)
      }
      Workbook(
        sheets = unique.toVector,
        metadata = metadata,
        activeSheetIndex = activeIndex
      )

  /** Generate ModificationTracker for property-based testing */
  val genModificationTracker: Gen[ModificationTracker] =
    for
      modifiedSheets <- Gen.containerOf[Set, Int](Gen.choose(0, 20))
      deletedSheets <- Gen.containerOf[Set, Int](Gen.choose(0, 20))
      reorderedSheets <- Gen.oneOf(true, false)
      modifiedMetadata <- Gen.oneOf(true, false)
    yield ModificationTracker(modifiedSheets, deletedSheets, reorderedSheets, modifiedMetadata)

  given Arbitrary[ModificationTracker] = Arbitrary(genModificationTracker)
