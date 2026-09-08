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

// ========== 3. Range fill: Excel Ctrl+Enter semantics, := is total ==========
val zeroed = filled.put(ref"E2:E4" := 0) // every cell in E2:E4 gets 0

// ========== 4. Typed values: codecs infer formats; smart detection for raw strings ==========
val stamped = zeroed
  .put(ref"F1", "As of")
  .put(ref"G1", LocalDate.of(2026, 6, 9))
  .put(ref"F2", "$1,234.56".toFormatted) // detected: Currency (also: percents, ISO dates)

// ========== 5. Recalculate: total, per-cell errors, cross-sheet aware ==========
val recalc = Workbook(stamped).recalculate()
if !recalc.isClean then recalc.errors.foreach(e => println(s"  ⚠ ${e.render}"))
val wb = recalc.workbook
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

// ========== 8. Records: a case class is a row (derives RowCodec) ==========
// Field order = column order, field names = header row, Option fields = empty cells.
final case class Order(id: Int, customer: String, qty: Int, price: BigDecimal, shipped: Option[LocalDate])
  derives RowCodec

val orders = Vector(
  Order(1, "Acme", 3, BigDecimal("9.99"), Some(LocalDate.of(2026, 1, 15))),
  Order(2, "Globex", 1, BigDecimal("120.00"), None)
)
val placed = Sheet("Orders").putTable(ref"A1", orders, "Orders").unsafe // header + rows + Excel table
val back = placed.sheet.readRowsByHeader[Order](Row.from1(1))            // Either[RowCodecError, Vector[Order]]
println(s"  ✓ Records: wrote ${placed.count} rows over ${placed.range.map(_.toA1).getOrElse("-")}, read back ${back.map(_.size).getOrElse(-1)}")

// ========== 9. Writing a model: writeChecked fills in the uncached formulas ==========
// `stamped` still holds the fx"" cells exactly as authored — no cached values. Excel.write would
// ship them blank to every cached-value consumer (openpyxl data_only, pandas, previewers, Excel
// before its first recalc); writeChecked computes exactly those cells, writes, and reports.
val checkedOut = "/tmp/scripting-tour-checked.xlsx"
val checked = Excel.writeChecked(Workbook(stamped), checkedOut)
println(s"  ✓ ${checked.summary}") // "Recalculated 3 formulas" — the same line `xl recalc` prints

// One sheet straight from disk (a typo throws an XLException naming the candidate sheets), and
// the workbook's shape without loading a cell.
val salesOnDisk = Excel.readSheet(checkedOut, "Sales")
println(s"  ✓ D2 on disk: ${salesOnDisk.readTypedOr[BigDecimal](ref"D2", BigDecimal(0))}")
val meta = Excel.readMetadata(checkedOut)
println(s"  ✓ Sheets: ${meta.sheets.map(_.name.value).mkString(", ")}")

// orExit: the value, or "error:/code:/hint:" on stderr and exit 1 — how a script fails like `xl`
val audited = orExit(salesOnDisk.putAt("F3", "audited"))
println(s"  ✓ orExit unwrapped a ${audited.cells.size}-cell sheet")

println("\n✨ One import. Compile-time refs. Total loops. Either at the edges.")
