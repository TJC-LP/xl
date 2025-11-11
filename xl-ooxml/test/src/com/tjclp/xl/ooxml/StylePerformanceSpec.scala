package com.tjclp.xl.ooxml

import munit.FunSuite
import com.tjclp.xl.*
import com.tjclp.xl.style.*

/**
 * Performance tests for style serialization
 *
 * Verifies that O(n²) → O(1) optimization in Styles.toXml scales linearly
 */
class StylePerformanceSpec extends FunSuite:

  test("style serialization scales linearly with 1000+ unique styles") {
    // Generate 1000 unique cell styles
    val styles = (0 until 1000).map { i =>
      val hue = (i * 360.0 / 1000.0).toInt
      val r = ((hue / 360.0) * 255).toInt & 0xFF
      val g = ((i / 1000.0) * 255).toInt & 0xFF
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
    val styleIndex = StyleIndex(
      fonts = fonts,
      fills = fills,
      borders = borders,
      numFmts = Vector.empty,
      cellStyles = styles.toVector,
      styleToIndex = styles.zipWithIndex.map { case (s, i) =>
        CellStyle.canonicalKey(s) -> StyleId(i)
      }.toMap
    )

    val ooxmlStyles = OoxmlStyles(styleIndex)

    // Measure serialization time
    val start = System.nanoTime()
    val xml = ooxmlStyles.toXml
    val elapsed = (System.nanoTime() - start) / 1_000_000 // Convert to milliseconds

    // Should complete in under 100ms (O(n) performance)
    // With O(n²), 1000 styles could take 1000ms+
    assert(elapsed < 100, s"Style serialization took ${elapsed}ms, expected <100ms for O(n) performance")

    // Verify XML was generated correctly
    val xmlString = xml.toString
    assert(xmlString.contains("<fonts"), "Should have fonts element")
    assert(xmlString.contains("<fills"), "Should have fills element")
    assert(xmlString.contains("<cellXfs"), "Should have cellXfs element")

    // Verify correct count
    assert(xmlString.contains(s"count=\"${styles.size}\""), s"Should have ${styles.size} cell styles")
  }

  test("performance comparison: 100 vs 1000 styles scales sub-quadratically") {
    def measureSerialization(styleCount: Int): Long =
      val styles = (0 until styleCount).map { i =>
        CellStyle(
          font = Font.default.copy(sizePt = 10.0 + (i % 20)),
          fill = Fill.Solid(Color.Rgb((0xff000000 | (i * 1000)) & 0xFFFFFFFF)),
          border = Border.none,
          numFmt = NumFmt.General,
          align = Align.default
        )
      }

      val styleIndex = StyleIndex(
        fonts = styles.map(_.font).distinct.toVector,
        fills = styles.map(_.fill).distinct.toVector,
        borders = Vector(Border.none),
        numFmts = Vector.empty,
        cellStyles = styles.toVector,
        styleToIndex = styles.zipWithIndex.map { case (s, i) =>
          CellStyle.canonicalKey(s) -> StyleId(i)
        }.toMap
      )

      val start = System.nanoTime()
      OoxmlStyles(styleIndex).toXml
      System.nanoTime() - start

    // Warm up JVM
    measureSerialization(10)
    measureSerialization(10)

    // Measure
    val time100 = measureSerialization(100)
    val time1000 = measureSerialization(1000)

    val ratio = time1000.toDouble / time100.toDouble

    // With O(n²), ratio would be ~100 (10x styles → 100x time)
    // With O(n), ratio should be ~10 (10x styles → 10x time)
    // Allow up to 20x for variance
    assert(ratio < 20.0, s"Performance ratio ${ratio} suggests O(n²) behavior (expected <20 for O(n))")
  }
