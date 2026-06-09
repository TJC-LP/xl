#!/usr/bin/env -S scala-cli shebang
//> using file project.scala

// Canonical tour of the scripting prelude — ONE import gives a script everything.
// This file is compile-verified by scripts/test-examples.sh and is the source of truth
// for snippets in the xl-scripting skill (plugin/skills/xl-scripting/).
//
// Run: ./mill __.publishLocal && scala-cli run examples/scripting_tour.sc

import com.tjclp.xl.scripting.{*, given}
import java.time.LocalDate

println("🧭 XL Scripting Tour\n")

// ========== 1. Build a sheet: compile-time literals are infallible ==========
// ref"A1" / fx"..." / money"..." are validated at compile time — typos fail the build.

val header = CellStyle.default.bold.size(12.0).center

val sales = Sheet("Sales")
  .put(ref"A1", "Product")
  .put(ref"B1", "Units")
  .put(ref"C1", "Price")
  .put(ref"D1", "Revenue")
  .style(ref"A1:D1", header)

// ========== 2. Bulk generation: fold data into a Patch (monoid composition) ==========
val products = List(("Widget", 150, 19.99), ("Gadget", 75, 29.99), ("Doohickey", 25, 49.99))

val rows = products.zipWithIndex.foldLeft(Patch.empty) { case (acc, ((name, units, price), i)) =>
  val r = ref"A2".shift(0, i) // total navigation — no Either in the loop
  // fx"" with literals validates at compile time and returns CellValue directly; with runtime
  // interpolation ($i) it validates at runtime and returns Either — unwrap at the boundary.
  acc ++
    (r := name) ++
    (r.shift(1, 0) := units) ++
    (r.shift(2, 0) := price) ++
    (r.shift(3, 0) := fx"=B${i + 2}*C${i + 2}".unsafe)
}

val filled = sales.put(rows)
println(s"  ✓ Built ${filled.cells.size} cells from ${products.size} records")

// ========== 3. Typed values: codecs infer formats (dates, decimals) ==========
val stamped = filled
  .put(ref"F1", "As of")
  .put(ref"G1", LocalDate.of(2026, 6, 9))

// ========== 4. Evaluate formulas: whole-workbook recalc with cached results ==========
val wb = Workbook(stamped).withCachedFormulas()
val sheet = wb.sheets.headOption.getOrElse(sys.exit(1))

// Single-formula evaluation returns XLResult — compose with for-comprehensions
val total = sheet.evaluateFormula("=SUM(D2:D4)")
println(s"  ✓ Total revenue: $total")

// ========== 5. Typed extraction: readTyped returns Either[CodecError, Option[A]] ==========
val firstUnits = sheet.readTyped[Int](ref"B2")
println(s"  ✓ First product units: $firstUnits")

// ========== 6. Display: excel interpolator formats through NumFmt ==========
given Sheet = sheet
println(excel"  ✓ A2 = ${ref"A2"}, B2 = ${ref"B2"}")

// ========== 7. IO at the edge: sync Excel facade, ONE .unsafe boundary ==========
val out = "/tmp/scripting-tour.xlsx"
Excel.write(wb, out)
val loaded = Excel.read(out)
println(s"  ✓ Round-trip: ${loaded.sheets.size} sheet(s), ${loaded.sheets.headOption.map(_.cells.size).getOrElse(0)} cells")

// Runtime strings return XLResult — unwrap once, at the edge, explicitly
val updated = loaded
  .update("Sales", _.put(ref"F2", "verified"))
  .unsafe
println(s"  ✓ Updated via XLResult + .unsafe boundary")

println("\n✨ One import. Compile-time refs. Total loops. Either at the edges.")
