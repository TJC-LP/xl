package com.tjclp.xl.ooxml

import munit.FunSuite
import com.tjclp.xl.style.{CellStyle, Align, HAlign, VAlign}

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
      Align(horizontal = HAlign.Right, vertical = VAlign.Top, wrapText = true, indent = 2)
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

    assertEquals(horizontal, "right", "horizontal should be 'right'")
    assertEquals(vertical, "top", "vertical should be 'top'")
    assertEquals(wrapText, "1", "wrapText should be '1'")
    assertEquals(indent, "2", "indent should be '2'")
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

    assertEquals(vertical, "", "vertical should not be present (default)")
    assertEquals(wrapText, "", "wrapText should not be present (default false)")
    assertEquals(indent, "", "indent should not be present (default 0)")
  }

  test("alignment round-trips through write/read cycle") {
    val originalAlign = Align(
      horizontal = HAlign.Justify,
      vertical = VAlign.Middle,
      wrapText = true,
      indent = 3
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
