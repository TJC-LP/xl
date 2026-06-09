package com.tjclp.xl

import com.tjclp.xl.{*, given}
import com.tjclp.xl.render.SvgRenderer
import com.tjclp.xl.sheets.{HeaderFooter, PageMargins, PageSetup, SheetView}
import munit.FunSuite

/**
 * GH-258/GH-259 model tests: SheetView (gridlines, zoom), PageSetup print extensions
 * (header/footer, margins, print area, repeat rows), and renderer gridline suppression.
 */
class SheetViewPageSetupSpec extends FunSuite:

  // ===== SheetView (GH-258) =====

  test("SheetView defaults match Excel (gridlines on, no explicit zoom)") {
    assertEquals(SheetView.default, SheetView(showGridLines = true, zoomScale = None))
  }

  test("SheetView accepts zoom bounds 10 and 400") {
    assertEquals(SheetView(zoomScale = Some(10)).zoomScale, Some(10))
    assertEquals(SheetView(zoomScale = Some(400)).zoomScale, Some(400))
  }

  test("SheetView rejects zoom outside 10-400") {
    intercept[IllegalArgumentException](SheetView(zoomScale = Some(9)))
    intercept[IllegalArgumentException](SheetView(zoomScale = Some(401)))
  }

  test("Sheet.withViewSettings sets viewSettings") {
    val view = SheetView(showGridLines = false, zoomScale = Some(85))
    val sheet = Sheet("S").withViewSettings(view)
    assertEquals(sheet.viewSettings, Some(view))
  }

  // ===== PageSetup extensions (GH-259) =====

  test("PageMargins defaults match Excel's Normal preset") {
    val m = PageMargins.default
    assertEquals((m.left, m.right, m.top, m.bottom, m.header, m.footer), (0.7, 0.7, 0.75, 0.75, 0.3, 0.3))
  }

  test("PageMargins rejects negative margins") {
    intercept[IllegalArgumentException](PageMargins(left = -0.1))
    intercept[IllegalArgumentException](PageMargins(footer = -1.0))
  }

  test("PageSetup accepts a valid repeat-rows span") {
    assertEquals(PageSetup(repeatRows = Some((1, 3))).repeatRows, Some((1, 3)))
    assertEquals(PageSetup(repeatRows = Some((5, 5))).repeatRows, Some((5, 5)))
  }

  test("PageSetup rejects invalid repeat-rows spans") {
    intercept[IllegalArgumentException](PageSetup(repeatRows = Some((0, 3)))) // 1-based
    intercept[IllegalArgumentException](PageSetup(repeatRows = Some((4, 2)))) // start > end
    intercept[IllegalArgumentException](PageSetup(repeatRows = Some((1, 1_048_577)))) // > max row
  }

  test("Sheet.withPageSetup sets pageSetup with all print extensions") {
    val setup = PageSetup(
      scale = 85,
      orientation = Some("landscape"),
      headerFooter = Some(HeaderFooter(oddFooter = Some("Page &P of &N"))),
      margins = Some(PageMargins(left = 1.0)),
      printArea = Some(ref"A1:D20"),
      repeatRows = Some((1, 2))
    )
    val sheet = Sheet("S").withPageSetup(setup)
    assertEquals(sheet.pageSetup, Some(setup))
  }

  // ===== Renderer integration (GH-258) =====

  test("SvgRenderer suppresses gridlines when sheet view disables them") {
    val sheet = Sheet("S")
      .put("A1" -> "X")
      .withViewSettings(SheetView(showGridLines = false))
    val svg = SvgRenderer.toSvg(sheet, ref"A1:B2", showGridlines = true)
    assert(!svg.contains("stroke=\"#D0D0D0\""), "sheet-level showGridLines=false must win")
  }

  test("SvgRenderer keeps caller-requested gridlines when sheet has no view settings") {
    val sheet = Sheet("S").put("A1" -> "X")
    val svg = SvgRenderer.toSvg(sheet, ref"A1:B2", showGridlines = true)
    assert(svg.contains("stroke=\"#D0D0D0\""), "caller-controlled gridlines unchanged by default")
  }
