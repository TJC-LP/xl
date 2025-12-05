//> using scala 3.7.3
//> using dep com.tjclp::xl-core:0.1.5-SNAPSHOT
//> using repository ivy2Local

/**
 * XL Row/Column Sizing Demo
 *
 * Demonstrates setting row heights and column widths, then exporting to
 * HTML and SVG with accurate dimensions.
 *
 * Features shown:
 * - Setting column widths with the DSL
 * - Setting row heights with the DSL
 * - Hiding rows and columns
 * - HTML export with colgroup and table-layout: fixed
 * - SVG export with accurate viewBox dimensions
 *
 * Run with: scala-cli run examples/resize-demo.sc
 */

import com.tjclp.xl.{*, given}
import com.tjclp.xl.macros.{ref, col}
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties}
import com.tjclp.xl.addressing.{Column, Row}
import com.tjclp.xl.dsl.RowColumnDsl.{row, width, hidden, height}

// ============================================================================
// STEP 1: Create a Sheet with Data
// ============================================================================

println("=" * 70)
println("STEP 1: Create Sheet with Data")
println("=" * 70)

val sheet = Sheet("Report")
  .put(
    ref"A1" -> "Product",
    ref"B1" -> "Q1 Sales",
    ref"C1" -> "Q2 Sales",
    ref"D1" -> "Notes",
    ref"A2" -> "Widget A",
    ref"B2" -> BigDecimal("125000"),
    ref"C2" -> BigDecimal("142500"),
    ref"D2" -> "Strong growth",
    ref"A3" -> "Widget B",
    ref"B3" -> BigDecimal("98000"),
    ref"C3" -> BigDecimal("87000"),
    ref"D3" -> "Declining",
    ref"A4" -> "Widget C",
    ref"B4" -> BigDecimal("210000"),
    ref"C4" -> BigDecimal("235000"),
    ref"D4" -> "Best seller"
  )

println("Created sheet with 4x4 data grid")
println()

// ============================================================================
// STEP 2: Set Column Widths Using DSL
// ============================================================================

println("=" * 70)
println("STEP 2: Set Column Widths Using DSL")
println("=" * 70)

// Column widths are in Excel "character units"
// Formula: pixels = (width * 7 + 5)
val sizingPatches = Vector(
  col"A".width(15.0).toPatch,   // Product column: 15 chars = 110px
  col"B".width(12.0).toPatch,   // Q1 Sales: 12 chars = 89px
  col"C".width(12.0).toPatch,   // Q2 Sales: 12 chars = 89px
  col"D".width(25.0).toPatch,   // Notes (wide): 25 chars = 180px
  row(0).height(30.0).toPatch,  // Header row: 30 points = 40px
  row(1).height(20.0).toPatch,  // Data rows: 20 points = 26px
  row(2).height(20.0).toPatch,
  row(3).height(20.0).toPatch
)
val sizingPatch = sizingPatches.reduce(Patch.combine)

val sizedSheet = Patch.applyPatch(sheet, sizingPatch)

println("Applied sizing:")
println("  Column A (Product): 15 chars = 110px")
println("  Column B (Q1 Sales): 12 chars = 89px")
println("  Column C (Q2 Sales): 12 chars = 89px")
println("  Column D (Notes): 25 chars = 180px")
println("  Row 0 (Header): 30pt = 40px")
println("  Rows 1-3 (Data): 20pt = 26px each")
println()

// ============================================================================
// STEP 3: Export to HTML
// ============================================================================

println("=" * 70)
println("STEP 3: Export to HTML")
println("=" * 70)

val html = sizedSheet.toHtml(ref"A1:D4")

println("HTML output (first 600 chars):")
println("-" * 40)
println(html.take(600))
println("...")
println("-" * 40)

// Verify sizing is in HTML
val hasColgroup = html.contains("<colgroup>")
val hasFixedLayout = html.contains("table-layout: fixed")
val hasRowHeight = html.contains("height:")

println()
println("Verification:")
println(s"  Has <colgroup>: $hasColgroup")
println(s"  Has table-layout: fixed: $hasFixedLayout")
println(s"  Has row heights: $hasRowHeight")
println()

// ============================================================================
// STEP 4: Export to SVG
// ============================================================================

println("=" * 70)
println("STEP 4: Export to SVG")
println("=" * 70)

val svg = sizedSheet.toSvg(ref"A1:D4")

// Extract viewBox dimensions
val viewBoxMatch = """viewBox="0 0 (\d+) (\d+)"""".r.findFirstMatchIn(svg)
viewBoxMatch match
  case Some(m) =>
    val width = m.group(1)
    val height = m.group(2)
    println(s"SVG viewBox dimensions: ${width}x${height} pixels")
    println()
    // Expected: HeaderWidth(40) + 110 + 89 + 89 + 180 = 508 width
    // Expected: HeaderHeight(24) + 40 + 26*3 = 142 height
  case None =>
    println("Could not extract viewBox dimensions")

println("SVG output (first 500 chars):")
println("-" * 40)
println(svg.take(500))
println("...")
println("-" * 40)
println()

// ============================================================================
// STEP 5: Hidden Rows/Columns Demo
// ============================================================================

println("=" * 70)
println("STEP 5: Hidden Rows/Columns Demo")
println("=" * 70)

// Hide column C (Q2 Sales) and row 3 (Widget B)
val hiddenPatch = Patch.combine(
  col"C".hidden.toPatch,
  row(2).hidden.toPatch
)

val hiddenSheet = Patch.applyPatch(sizedSheet, hiddenPatch)

println("Hidden column C (Q2 Sales) and row 3 (Widget B)")

val hiddenHtml = hiddenSheet.toHtml(ref"A1:D4")
val hiddenSvg = hiddenSheet.toSvg(ref"A1:D4")

// Verify hidden elements are 0px
val hasZeroWidth = hiddenHtml.contains("width: 0px")
val hasZeroHeight = hiddenHtml.contains("height: 0px")

println()
println("HTML verification:")
println(s"  Has zero-width column: $hasZeroWidth")
println(s"  Has zero-height row: $hasZeroHeight")

// Check SVG viewBox is smaller
val hiddenViewBoxMatch = """viewBox="0 0 (\d+) (\d+)"""".r.findFirstMatchIn(hiddenSvg)
hiddenViewBoxMatch match
  case Some(m) =>
    val width = m.group(1)
    val height = m.group(2)
    println(s"  SVG with hidden: ${width}x${height} pixels (smaller due to hidden row/col)")
  case None =>
    println("  Could not extract hidden viewBox dimensions")

println()

// ============================================================================
// STEP 6: Using Direct API (Alternative to DSL)
// ============================================================================

println("=" * 70)
println("STEP 6: Direct API Alternative")
println("=" * 70)

// You can also use the direct Sheet API
val directSheet = Sheet("Direct")
  .put(ref"A1" -> "Data")
  .setColumnProperties(Column.from0(0), ColumnProperties(width = Some(20.0)))
  .setRowProperties(Row.from0(0), RowProperties(height = Some(30.0)))

println("Created sheet using direct API:")
println("  sheet.setColumnProperties(Column.from0(0), ColumnProperties(width = Some(20.0)))")
println("  sheet.setRowProperties(Row.from0(0), RowProperties(height = Some(30.0)))")

val directSvg = directSheet.toSvg(ref"A1:A1")
val directViewBox = """viewBox="0 0 (\d+) (\d+)"""".r.findFirstMatchIn(directSvg)
directViewBox.foreach { m =>
  println(s"  SVG viewBox: ${m.group(1)}x${m.group(2)} pixels")
}

println()

// ============================================================================
// Summary
// ============================================================================

println("=" * 70)
println("SUMMARY")
println("=" * 70)
println("""
Row/Column Sizing Features:
  - col"A".width(15.0).toPatch      Set column width (character units)
  - row(0).height(30.0).toPatch     Set row height (points)
  - col"C".hidden.toPatch           Hide a column
  - row(2).hidden.toPatch           Hide a row

Unit Conversions:
  - Column width: pixels = (chars * 7 + 5)
  - Row height: pixels = (points * 4/3)

HTML Export:
  - Uses <colgroup> with explicit widths
  - Uses table-layout: fixed for accurate sizing
  - Hidden rows/columns render as 0px

SVG Export:
  - viewBox reflects actual dimensions
  - Column widths and row heights affect layout
  - Hidden elements render as 0px
""")
