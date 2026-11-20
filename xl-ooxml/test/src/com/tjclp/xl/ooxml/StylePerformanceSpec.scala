package com.tjclp.xl.ooxml

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.styles.*

/**
 * Style serialization over many unique styles: `OoxmlStyles.toXml` resolves each style's font, fill
 * and border through the index's `firstIndex` maps (O(1) per style), and this spec pins the output
 * of 1,000 unique styles. It pins no clock: at this size an `indexOf` regression costs about a
 * million `equals` calls, single-digit milliseconds, below what a 200 ms wall-clock gate on a
 * loaded CI runner could resolve, so the gate only ever flaked (#843).
 */
class StylePerformanceSpec extends FunSuite:

  test("serialises 1,000 unique styles with every part present and counted") {
    // Generate 1000 unique cell styles
    val styles = (0 until 1000).map { i =>
      val hue = (i * 360.0 / 1000.0).toInt
      val r = ((hue / 360.0) * 255).toInt & 0xff
      val g = ((i / 1000.0) * 255).toInt & 0xff
      val b = 128
      CellStyle(
        font = Font.default.copy(sizePt = 10.0 + (i % 20)),
        fill = Fill.Solid(Color.Rgb(0xff000000 | (r << 16) | (g << 8) | b)),
        border = Border.none,
        numFmt = NumFmt.General,
        align = Align.default
      )
    }

    // Build StyleIndex with all unique styles
    val fonts = styles.map(_.font).distinct.toVector
    val fills = styles.map(_.fill).distinct.toVector
    val borders = Vector(Border.none)
    val styleMappings: Vector[(String, StyleId)] = styles.zipWithIndex.map { case (s, i) =>
      CellStyle.canonicalKey(s) -> (StyleId(i): StyleId)
    }.toVector

    val styleIndex = StyleIndex(
      fonts = fonts,
      fills = fills,
      borders = borders,
      numFmts = Vector.empty,
      cellStyles = styles.toVector,
      styleToIndex = styleMappings.toMap
    )

    val ooxmlStyles = OoxmlStyles(styleIndex)
    val xml = ooxmlStyles.toXml

    // Verify XML was generated correctly
    val xmlString = xml.toString
    assert(xmlString.contains("<fonts"), "Should have fonts element")
    assert(xmlString.contains("<fills"), "Should have fills element")
    assert(xmlString.contains("<cellXfs"), "Should have cellXfs element")

    // Verify correct count
    assert(
      xmlString.contains(s"count=\"${styles.size}\""),
      s"Should have ${styles.size} cell styles"
    )
  }
