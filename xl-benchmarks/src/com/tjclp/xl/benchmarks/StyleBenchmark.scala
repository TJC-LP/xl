package com.tjclp.xl.benchmarks

import com.tjclp.xl.*
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.{CellStyle, StyleRegistry, Color, Font, Fill}
import com.tjclp.xl.styles.units.StyleId
import org.openjdk.jmh.annotations.*
import java.util.concurrent.TimeUnit
import scala.compiletime.uninitialized

/**
 * Benchmarks for style system operations.
 *
 * Tests:
 *   - StyleRegistry.register() deduplication
 *   - CellStyle.canonicalKey computation
 *   - Style application to cell ranges
 *   - Unique vs duplicate style registration
 */
@State(Scope.Benchmark)
@BenchmarkMode(Array(Mode.AverageTime))
@OutputTimeUnit(TimeUnit.NANOSECONDS)
@Warmup(iterations = 3, time = 1, timeUnit = TimeUnit.SECONDS)
@Measurement(iterations = 5, time = 2, timeUnit = TimeUnit.SECONDS)
@Fork(value = 1)
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.Null"))
class StyleBenchmark {

  var baseSheet: Sheet = uninitialized
  var registry: StyleRegistry = uninitialized
  var uniqueStyles: Seq[CellStyle] = uninitialized
  var duplicateStyles: Seq[CellStyle] = uninitialized
  var testStyle: CellStyle = uninitialized
  var largeRange: CellRange = uninitialized

  @Setup(Level.Trial)
  def setup(): Unit = {
    // Create base sheet
    baseSheet = BenchmarkUtils.generateSheet("Test", 1000, styled = false)
    registry = StyleRegistry.default

    // Generate unique styles (1000 different styles)
    uniqueStyles = BenchmarkUtils.generateStyles(1000)

    // Generate duplicate styles (same style repeated 1000 times)
    val singleStyle = CellStyle.default
      .withFont(Font.default.copy(bold = true))
      .withFill(Fill.Solid(Color.Rgb(0xff0000)))
    duplicateStyles = Seq.fill(1000)(singleStyle)

    // Test style for canonicalKey
    testStyle = CellStyle.default
      .withFont(Font.default.copy(bold = true, italic = true, sizePt = 14.0))
      .withFill(Fill.Solid(Color.Rgb(0x4472c4)))

    // Large range for style application tests
    largeRange = CellRange.parse("A1:Z100").getOrElse(CellRange(ref"A1", ref"A1"))
  }

  @Benchmark
  def canonicalKeyComputation(): String = {
    // Measure time to compute canonical key
    testStyle.canonicalKey
  }

  @Benchmark
  def registerUniqueStyle(): (StyleRegistry, StyleId) = {
    // Measure registration of a new unique style (deduplication miss)
    val newStyle = CellStyle.default.withFont(Font.default.copy(sizePt = 99.0))
    registry.register(newStyle)
  }

  @Benchmark
  def registerDuplicateStyle(): (StyleRegistry, StyleId) = {
    // Measure registration of existing style (deduplication hit)
    val (reg1, _) = registry.register(testStyle)
    reg1.register(testStyle) // Second registration should be faster
  }

  @Benchmark
  def registerManyUniqueStyles(): StyleRegistry = {
    // Measure overhead of registering many unique styles
    var reg = StyleRegistry.default
    uniqueStyles.foreach { style =>
      val (updatedReg, _) = reg.register(style)
      reg = updatedReg
    }
    reg
  }

  @Benchmark
  def registerManyDuplicateStyles(): StyleRegistry = {
    // Measure overhead with perfect deduplication (all same style)
    var reg = StyleRegistry.default
    duplicateStyles.foreach { style =>
      val (updatedReg, _) = reg.register(style)
      reg = updatedReg
    }
    reg
  }

  @Benchmark
  def styleIndexLookup(): Option[StyleId] = {
    // Measure indexOf performance (O(n) vector search)
    val (reg, _) = registry.register(testStyle)
    reg.indexOf(testStyle)
  }

  @Benchmark
  def applyStyleToRange(): Sheet = {
    // Measure overhead of applying style to large range (2600 cells)
    import com.tjclp.xl.sheets.styleSyntax.*
    baseSheet.withRangeStyle(largeRange, testStyle)
  }
}
