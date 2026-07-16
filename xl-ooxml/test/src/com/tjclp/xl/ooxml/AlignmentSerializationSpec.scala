package com.tjclp.xl.ooxml

import scala.xml.XML

import munit.FunSuite
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.sheets.styleSyntax.*
import com.tjclp.xl.styles.{CellStyle, Align, HAlign, VAlign}
import com.tjclp.xl.styles.dsl.*
import com.tjclp.xl.workbooks.Workbook

/** Tests for alignment serialization to styles.xml */
@SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
class AlignmentSerializationSpec extends FunSuite:

  test("default alignment omits <alignment> element") {
    val style = CellStyle.default
    val index = StyleIndex(
      fonts = Vector.empty,
      fills = Vector.empty,
      borders = Vector.empty,
      numFmts = Vector.empty,
      cellStyles = Vector(style),
      styleToIndex = Map.empty
    )
    val styles = OoxmlStyles(index)
    val xml = styles.toXml

    val xfElems = xml \ "cellXfs" \ "xf"
    assert(xfElems.nonEmpty, "Should have xf elements")

    // Default style should have applyAlignment="0" and no <alignment> child
    val firstXf = xfElems(0)
    val applyAlignment = (firstXf \ "@applyAlignment").text
    assertEquals(applyAlignment, "0", "Default style should have applyAlignment='0'")

    val alignmentElems = firstXf \ "alignment"
    assert(alignmentElems.isEmpty, "Default style should not have <alignment> element")
  }

  test("non-default alignment includes applyAlignment=1") {
    val style = CellStyle.default.withAlign(Align(horizontal = HAlign.Center))
    val index = StyleIndex(
      fonts = Vector.empty,
      fills = Vector.empty,
      borders = Vector.empty,
      numFmts = Vector.empty,
      cellStyles = Vector(style),
      styleToIndex = Map.empty
    )
    val styles = OoxmlStyles(index)
    val xml = styles.toXml

    val xfElems = xml \ "cellXfs" \ "xf"
    val firstXf = xfElems(0)
    val applyAlignment = (firstXf \ "@applyAlignment").text
    assertEquals(applyAlignment, "1", "Non-default alignment should have applyAlignment='1'")
  }

  test("alignment serializes to <alignment> in cellXfs") {
    val style = CellStyle.default.withAlign(
      Align(
        horizontal = HAlign.Right,
        vertical = VAlign.Top,
        wrapText = true,
        indent = 2,
        textRotation = 90
      )
    )
    val index = StyleIndex(
      fonts = Vector.empty,
      fills = Vector.empty,
      borders = Vector.empty,
      numFmts = Vector.empty,
      cellStyles = Vector(style),
      styleToIndex = Map.empty
    )
    val styles = OoxmlStyles(index)
    val xml = styles.toXml

    val xfElems = xml \ "cellXfs" \ "xf"
    val firstXf = xfElems(0)
    val alignmentElem = firstXf \ "alignment"

    assert(alignmentElem.nonEmpty, "Should have <alignment> element for non-default alignment")

    // Check attributes
    val horizontal = (alignmentElem \ "@horizontal").text
    val vertical = (alignmentElem \ "@vertical").text
    val wrapText = (alignmentElem \ "@wrapText").text
    val indent = (alignmentElem \ "@indent").text
    val textRotation = (alignmentElem \ "@textRotation").text

    assertEquals(horizontal, "right", "horizontal should be 'right'")
    assertEquals(vertical, "top", "vertical should be 'top'")
    assertEquals(wrapText, "1", "wrapText should be '1'")
    assertEquals(indent, "2", "indent should be '2'")
    assertEquals(textRotation, "90", "textRotation should be '90'")
  }

  test("alignment with only horizontal changed emits only horizontal attribute") {
    val style = CellStyle.default.withAlign(Align(horizontal = HAlign.Center))
    val index = StyleIndex(
      fonts = Vector.empty,
      fills = Vector.empty,
      borders = Vector.empty,
      numFmts = Vector.empty,
      cellStyles = Vector(style),
      styleToIndex = Map.empty
    )
    val styles = OoxmlStyles(index)
    val xml = styles.toXml

    val alignmentElem = (xml \ "cellXfs" \ "xf" \ "alignment").head
    val horizontal = (alignmentElem \ "@horizontal").text
    assertEquals(horizontal, "center", "Should have horizontal='center'")

    // Verify other attributes are not present (using defaults)
    val vertical = (alignmentElem \ "@vertical").text
    val wrapText = (alignmentElem \ "@wrapText").text
    val indent = (alignmentElem \ "@indent").text
    val textRotation = (alignmentElem \ "@textRotation").text

    assertEquals(vertical, "", "vertical should not be present (default)")
    assertEquals(wrapText, "", "wrapText should not be present (default false)")
    assertEquals(indent, "", "indent should not be present (default 0)")
    assertEquals(textRotation, "", "textRotation should not be present (default 0)")
  }

  test("alignment round-trips through write/read cycle") {
    val originalAlign = Align(
      horizontal = HAlign.Justify,
      vertical = VAlign.Middle,
      wrapText = true,
      indent = 3,
      textRotation = 135
    )
    val style = CellStyle.default.withAlign(originalAlign)
    val index = StyleIndex(
      fonts = Vector.empty,
      fills = Vector.empty,
      borders = Vector.empty,
      numFmts = Vector.empty,
      cellStyles = Vector(style),
      styleToIndex = Map.empty
    )
    val styles = OoxmlStyles(index)
    val xml = styles.toXml

    // Parse it back
    val parsedStyles = WorkbookStyles.fromXml(xml)

    parsedStyles match
      case Right(workbookStyles) =>
        assert(workbookStyles.cellStyles.nonEmpty, "Should have parsed cell styles")
        val parsedStyle = workbookStyles.cellStyles(0)
        assertEquals(
          parsedStyle.align,
          originalAlign,
          "Alignment should round-trip correctly"
        )
      case Left(error) =>
        fail(s"Failed to parse styles: $error")
  }

  test("workbook with indented styles round-trips indent through write/read (GH-260)") {
    // House templates indent line items via the alignment indent attribute, not leading
    // spaces — the indent must survive a full workbook write → read cycle.
    val item = CellStyle.default.left
    val sub = CellStyle.default.left.indent(1)
    val subsub = CellStyle.default.left.indent(2)

    val sheet = Sheet("Indent")
      .put(ref"A1", CellValue.Text("Revenue"))
      .put(ref"A2", CellValue.Text("Product"))
      .put(ref"A3", CellValue.Text("Hardware"))
      .withCellStyle(ref"A1", item)
      .withCellStyle(ref"A2", sub)
      .withCellStyle(ref"A3", subsub)
    val wb = Workbook(Vector(sheet))

    val bytes = XlsxWriter.writeToBytes(wb) match
      case Right(b) => b
      case Left(err) => fail(s"Write failed: $err")
    val readWb = XlsxReader.readFromBytes(bytes) match
      case Right(w) => w
      case Left(err) => fail(s"Read failed: $err")

    val readSheet = readWb.sheets(0)
    def indentOf(r: ARef): Int =
      readSheet.getCellStyle(r).map(_.align.indent).getOrElse(-1)

    assertEquals(indentOf(ref"A1"), 0, "A1 should have no indent")
    assertEquals(indentOf(ref"A2"), 1, "A2 indent should round-trip")
    assertEquals(indentOf(ref"A3"), 2, "A3 indent should round-trip")
  }

  test("style dedup keeps styles differing only by indent distinct after round-trip (GH-260)") {
    // canonicalKey accounts for indent: two styles identical except for indent must map to
    // distinct xf entries in styles.xml, not collapse into one.
    val indented1 = CellStyle.default.left.indent(1)
    val indented2 = CellStyle.default.left.indent(2)

    val sheet = Sheet("Dedup")
      .put(ref"A1", CellValue.Text("one"))
      .put(ref"A2", CellValue.Text("two"))
      .withCellStyle(ref"A1", indented1)
      .withCellStyle(ref"A2", indented2)
    val wb = Workbook(Vector(sheet))

    val bytes = XlsxWriter.writeToBytes(wb) match
      case Right(b) => b
      case Left(err) => fail(s"Write failed: $err")
    val readWb = XlsxReader.readFromBytes(bytes) match
      case Right(w) => w
      case Left(err) => fail(s"Read failed: $err")

    val readSheet = readWb.sheets(0)
    val ids = Vector(ref"A1", ref"A2").flatMap(r => readSheet.cells.get(r).flatMap(_.styleId))
    assertEquals(ids.size, 2, "both cells should carry a style after round-trip")
    assertEquals(ids.distinct.size, 2, "indent-only style variants must not collapse")
    assertEquals(readSheet.getCellStyle(ref"A1").map(_.align.indent), Some(1))
    assertEquals(readSheet.getCellStyle(ref"A2").map(_.align.indent), Some(2))
  }

  test("workbook with rotated captions round-trips textRotation through write/read (GH-380)") {
    // Sensitivity-grid row-axis labels are 90-degree rotated text in a tall merged cell; the
    // rotation must survive a full workbook write -> read cycle.
    val up = CellStyle.default.center.rotated(90)
    val down = CellStyle.default.center.rotated(-45) // ST_TextRotation 135 (downward encoding)
    val stacked = CellStyle.default.center.rotated(255)

    val sheet = Sheet("Rotation")
      .put(ref"A1", CellValue.Text("Exit Multiple"))
      .put(ref"A2", CellValue.Text("Down"))
      .put(ref"A3", CellValue.Text("Stacked"))
      .withCellStyle(ref"A1", up)
      .withCellStyle(ref"A2", down)
      .withCellStyle(ref"A3", stacked)
    val wb = Workbook(Vector(sheet))

    val bytes = XlsxWriter.writeToBytes(wb) match
      case Right(b) => b
      case Left(err) => fail(s"Write failed: $err")
    val readWb = XlsxReader.readFromBytes(bytes) match
      case Right(w) => w
      case Left(err) => fail(s"Read failed: $err")

    val readSheet = readWb.sheets(0)
    def rotationOf(r: ARef): Int =
      readSheet.getCellStyle(r).map(_.align.textRotation).getOrElse(-1)

    assertEquals(rotationOf(ref"A1"), 90, "upward rotation should round-trip")
    assertEquals(rotationOf(ref"A2"), 135, "downward rotation should round-trip")
    assertEquals(rotationOf(ref"A3"), 255, "vertical stacked should round-trip")
  }

  test(
    "style dedup keeps styles differing only by textRotation distinct after round-trip (GH-380)"
  ) {
    // canonicalKey accounts for textRotation: two styles identical except for rotation must map
    // to distinct xf entries in styles.xml, not collapse into one.
    val rotated1 = CellStyle.default.left.rotated(45)
    val rotated2 = CellStyle.default.left.rotated(90)

    val sheet = Sheet("DedupRot")
      .put(ref"A1", CellValue.Text("one"))
      .put(ref"A2", CellValue.Text("two"))
      .withCellStyle(ref"A1", rotated1)
      .withCellStyle(ref"A2", rotated2)
    val wb = Workbook(Vector(sheet))

    val bytes = XlsxWriter.writeToBytes(wb) match
      case Right(b) => b
      case Left(err) => fail(s"Write failed: $err")
    val readWb = XlsxReader.readFromBytes(bytes) match
      case Right(w) => w
      case Left(err) => fail(s"Read failed: $err")

    val readSheet = readWb.sheets(0)
    val ids = Vector(ref"A1", ref"A2").flatMap(r => readSheet.cells.get(r).flatMap(_.styleId))
    assertEquals(ids.size, 2, "both cells should carry a style after round-trip")
    assertEquals(ids.distinct.size, 2, "rotation-only style variants must not collapse")
    assertEquals(readSheet.getCellStyle(ref"A1").map(_.align.textRotation), Some(45))
    assertEquals(readSheet.getCellStyle(ref"A2").map(_.align.textRotation), Some(90))
  }

  test("reader keeps valid textRotation values and drops invalid ones (GH-380)") {
    // ST_TextRotation admits 0-180 plus 255 (vertical stacked); anything else from a malformed
    // file is dropped so the parser stays total (Align requires the same range).
    def parsedRotation(attr: String): Int =
      val xml = XML.loadString(
        s"""<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
           |  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
           |  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
           |  <borders count="1"><border/></borders>
           |  <cellXfs count="1">
           |    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1">
           |      <alignment horizontal="left" textRotation="$attr"/>
           |    </xf>
           |  </cellXfs>
           |</styleSheet>""".stripMargin
      )
      WorkbookStyles.fromXml(xml) match
        case Right(styles) =>
          styles.cellStyles.headOption.map(_.align.textRotation).getOrElse(fail("no style parsed"))
        case Left(err) => fail(s"styles parse must not fail: $err")

    assertEquals(parsedRotation("90"), 90, "in-range rotation kept")
    assertEquals(parsedRotation("180"), 180, "boundary rotation kept")
    assertEquals(parsedRotation("255"), 255, "vertical stacked sentinel kept")
    assertEquals(parsedRotation("181"), 0, "181-254 gap is invalid and dropped")
    assertEquals(parsedRotation("200"), 0, "invalid rotation dropped")
    assertEquals(parsedRotation("-10"), 0, "negative rotation dropped")
    assertEquals(parsedRotation("junk"), 0, "unparseable rotation dropped")
  }

  test("all HAlign enum values serialize correctly") {
    import HAlign.*

    // Note: HAlign.Left is the default, so it won't be emitted unless another property is non-default
    // Test non-default values only
    val testCases = List(
      (Center, "center"),
      (Right, "right"),
      (Justify, "justify"),
      (Fill, "fill")
    )

    testCases.foreach { case (hAlign, expected) =>
      val style = CellStyle.default.withAlign(Align(horizontal = hAlign))
      val index = StyleIndex(
        fonts = Vector.empty,
        fills = Vector.empty,
        borders = Vector.empty,
        numFmts = Vector.empty,
        cellStyles = Vector(style),
        styleToIndex = Map.empty
      )
      val styles = OoxmlStyles(index)
      val xml = styles.toXml

      val horizontal = (xml \ "cellXfs" \ "xf" \ "alignment" \ "@horizontal").text
      assertEquals(horizontal, expected, s"HAlign.$hAlign should serialize as '$expected'")
    }
  }

  test("all VAlign enum values serialize correctly") {
    import VAlign.*

    // Note: VAlign.Bottom is the default, so only test non-default values
    val testCases = List(
      (Top, "top"),
      (Middle, "center"), // VAlign.Middle → "center" per OOXML spec
      (Justify, "justify"),
      (Distributed, "distributed")
    )

    testCases.foreach { case (vAlign, expected) =>
      val style = CellStyle.default.withAlign(Align(vertical = vAlign))
      val index = StyleIndex(
        fonts = Vector.empty,
        fills = Vector.empty,
        borders = Vector.empty,
        numFmts = Vector.empty,
        cellStyles = Vector(style),
        styleToIndex = Map.empty
      )
      val styles = OoxmlStyles(index)
      val xml = styles.toXml

      val vertical = (xml \ "cellXfs" \ "xf" \ "alignment" \ "@vertical").text
      assertEquals(vertical, expected, s"VAlign.$vAlign should serialize as '$expected'")
    }
  }
