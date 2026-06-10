package com.tjclp.xl.ooxml

import munit.ScalaCheckSuite
import org.scalacheck.Prop
import org.scalacheck.Prop.forAll

import com.tjclp.xl.Generators
import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{Cell, CellValue, Comment}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.{FreezePane, HeaderFooter, PageMargins, PageSetup, SheetView}
import com.tjclp.xl.sheets.styleSyntax.*
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.alignment.{Align, HAlign, VAlign}
import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.styles.color.{Color, ThemeSlot}
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * GH-240 (part a): the GENERATIVE round-trip law.
 *
 * `readFromBytes(writeToBytes(wb)) ≈ wb` for arbitrary generated workbooks — unlike
 * OoxmlRoundTripSpec, the inputs are NOT hand-picked values the writer is known to handle; they are
 * drawn from the full generator distribution (styles, comments, hyperlinks, formulas, merges, sheet
 * views, page setup, unicode text, magnitude-extreme numbers).
 *
 * The binding definition of ≈ lives in [[WorkbookEquivalence]] (reusable by parity and fixture
 * specs). In short: same sheet names in order; same cell refs; CellValue equal (formulas by text,
 * numbers by BigDecimal compare, DateTime ≈ stored serial); RESOLVED cell styles equal (registry
 * indices may be renumbered by dedup, numFmtId is reader-filled metadata); comments equal by (plain
 * text, author); hyperlinks equal; merges set-equal; viewSettings/pageSetup equal. Explicitly
 * ignored serialization noise: styleId numbering, SST ordering, workbook metadata +
 * activeSheetIndex (not serialized — GH-242), dimension/defaultColWidth-style derived props,
 * write-only freezePane, cached formula values, and Rgb alpha-00 ≡ alpha-FF (the parser
 * deliberately canonicalizes 00 alpha to opaque, matching Excel/openpyxl's 00RRGGBB convention).
 *
 * Determinism: the suite fixes the ScalaCheck initial seed so CI failures are reproducible.
 * Locally, crank the run count to shake out new bugs:
 * `XL_ROUNDTRIP_MIN_SUCCESS=1000 ./mill xl-ooxml.test.testOnly com.tjclp.xl.ooxml.OoxmlGenerativeRoundTripSpec`
 * (or `-Dxl.roundtrip.minSuccess=...` where your runner forwards JVM props).
 */
class OoxmlGenerativeRoundTripSpec extends ScalaCheckSuite:

  private val minSuccess: Int =
    sys.props
      .get("xl.roundtrip.minSuccess")
      .orElse(sys.env.get("XL_ROUNDTRIP_MIN_SUCCESS"))
      .flatMap(_.toIntOption)
      .getOrElse(100)

  override def scalaCheckTestParameters: org.scalacheck.Test.Parameters =
    super.scalaCheckTestParameters
      .withMinSuccessfulTests(minSuccess)
      // Fixed seed: CI failures replay exactly. The generator space was shaken out at 1000+
      // cases across multiple random seeds before pinning (see GH-240).
      .withInitialSeed(org.scalacheck.rng.Seed(20260610L))

  private def roundTripDiff(wb: Workbook): List[String] =
    XlsxWriter.writeToBytes(wb) match
      case Left(err) => List(s"write failed: ${err.message}")
      case Right(bytes) =>
        XlsxReader.readFromBytes(bytes) match
          case Left(err) => List(s"read failed: ${err.message}")
          case Right(readBack) => WorkbookEquivalence.diff(wb, readBack)

  private def roundTripProp(wb: Workbook): Prop =
    val diffs = roundTripDiff(wb)
    Prop(diffs.isEmpty) :| diffs.mkString("\n")

  property("LAW: read(write(wb)) ≈ wb — rich workbooks (styles/comments/links/formulas/views)") {
    forAll(Generators.genRichWorkbook)(roundTripProp)
  }

  property("LAW: read(write(wb)) ≈ wb — legacy distribution (full-grid refs, dangling styleIds)") {
    forAll(Generators.genWorkbook)(roundTripProp)
  }

  // A deterministic everything-at-once workbook: readable failure output for debugging the law,
  // and a regression anchor independent of the generator distribution.
  test("deterministic kitchen-sink workbook round-trips under ≈") {
    val style = CellStyle(
      font = Font(
        "Arial",
        10.5,
        bold = true,
        italic = false,
        underline = true,
        Some(Color.Rgb(0xff4472c4))
      ),
      fill = Fill.Solid(Color.Theme(ThemeSlot.Accent2, 0.25)),
      border = Border(
        left = BorderSide(BorderStyle.Thin, Some(Color.Rgb(0xff000000))),
        right = BorderSide(BorderStyle.MediumDashDot, None),
        top = BorderSide(BorderStyle.None, None),
        bottom = BorderSide(BorderStyle.Double, Some(Color.Theme(ThemeSlot.Dark1, 0.0)))
      ),
      numFmt = NumFmt.Custom("$#,##0.00;[Red]($#,##0.00)"),
      numFmtId = None,
      align = Align(HAlign.Center, VAlign.Middle, wrapText = true, indent = 2)
    )
    val sheet = Sheet(SheetName.unsafe("P&L 1"))
      .put(Cell(ref"A1", CellValue.Text("héllo <world> & \"friends\"")))
      .put(Cell(ref"B2", CellValue.Number(BigDecimal("1E20"))))
      .put(Cell(ref"C3", CellValue.Formula("SUM(A1:B2)", Some(CellValue.Number(BigDecimal(7))))))
      .put(Cell(ref"D4", CellValue.Bool(true), None, None, Some("https://example.com/x")))
      .put(Cell(ref"E5", CellValue.DateTime(java.time.LocalDateTime.of(2026, 6, 10, 12, 30, 0))))
      .put(Cell(ref"F6", CellValue.Empty))
      .withCellStyle(ref"A1", style)
      .comment(ref"B2", Comment.plainText("check this\nnumber", Some("Reviewer")))
      .comment(ref"F6", Comment.plainText("unauthored note", None))
      .merge(ref"A1:C1")
      .copy(
        viewSettings = Some(SheetView(showGridLines = false, zoomScale = Some(85))),
        pageSetup = Some(
          PageSetup(
            scale = 100,
            orientation = Some("landscape"),
            fitToWidth = Some(1),
            fitToHeight = None,
            headerFooter = Some(
              HeaderFooter(
                oddHeader = Some("&LTHE JORDAN COMPANY&RConfidential"),
                oddFooter = Some("&CPage &P of &N"),
                evenHeader = Some("&CEven"),
                firstHeader = Some("&CFirst"),
                differentOddEven = true,
                differentFirst = true
              )
            ),
            margins = Some(PageMargins(0.5, 0.5, 0.75, 0.75, 0.3, 0.3))
          )
        ),
        freezePane = Some(FreezePane.At(ref"B2"))
      )
    val wb = Workbook(Vector(sheet))
    val diffs = roundTripDiff(wb)
    assert(diffs.isEmpty, diffs.mkString("\n"))
  }
