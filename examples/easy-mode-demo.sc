//> using scala 3.7.3
//> using dep com.tjclp::xl-core:0.1.0-SNAPSHOT
//> using dep com.tjclp::xl-cats-effect:0.1.0-SNAPSHOT
//> using repository ivy2Local

// Standalone demo script - run with:
//   1. Publish locally: ./mill xl-core.publishLocal && ./mill xl-cats-effect.publishLocal
//   2. Run script: scala-cli run examples/easy-mode-demo.sc

import com.tjclp.xl.*
import com.tjclp.xl.io.EasyExcel as Excel
import com.tjclp.xl.error.XLException
import com.tjclp.xl.unsafe.*
import java.time.LocalDate

println("🚀 XL Easy Mode API Demo\n")

// ========== Example 1: String-Based References ==========
println("📋 Example 1: String-Based References (Easy Mode)")

val headerStyle = CellStyle.default.bold.size(12.0).center
val titleStyle = CellStyle.default.bold.size(14.0).bgBlue.white

val report = Sheet("Q1 Report")
  .applyStyle("A1:D1", titleStyle)      // Chainable XLResult[Sheet]
  .applyStyle("A2:D2", headerStyle)
  .put("A1", "Q1 2025 Sales Report")    // Chains on XLResult[Sheet]
  .put("A2", "Product")
  .put("B2", "Units")
  .put("A3", "Widget")
  .put("B3", 150)
  .unsafe  // Single unwrap at end

println(s"  ✓ Used string refs for ${report.cells.size} cells")
println()

// ========== Example 2: Inline Styling ==========
println("🎨 Example 2: Inline Styling")

val quickSheet = Sheet("Quick")
  .put("A1", "Alert", CellStyle.default.bold.red)
  .put("A2", "Success", CellStyle.default.bold.green)
  .unsafe

println(s"  ✓ Applied inline styles")
println()

// ========== Example 3: Rich Text ==========
println("💬 Example 3: Rich Text")

val richSheet = Sheet("RichText")
  .put("A1", "Status: ".bold + "ACTIVE".green.bold)
  .put("A2", "Error: ".red.bold + "Fix immediately!")
  .unsafe

println(s"  ✓ Created rich text cells")
println()

// ========== Example 4: Safe Lookups ==========
println("🔍 Example 4: Safe Lookups")

val value = report.getCell("A3")
val cells = report.getCells("A3:B3")

println(s"  ✓ Looked up cell: ${value.map(_.value)}")
println(s"  ✓ Found ${cells.size} cells in range")
println()

// ========== Example 5: Excel IO ==========
println("💾 Example 5: Excel IO (EasyExcel)")

val workbook = Workbook.empty
  .addSheet(report)
  .addSheet(quickSheet)
  .addSheet(richSheet)
  .unsafe  // Single unwrap at the end!

Excel.write(workbook, "/tmp/easy-mode-demo.xlsx")
println(s"  ✓ Wrote /tmp/easy-mode-demo.xlsx")

val loaded = Excel.read("/tmp/easy-mode-demo.xlsx")
println(s"  ✓ Read back (${loaded.sheets.size} sheets)")

// Demonstrate in-memory modification
val firstSheet = loaded.sheets.headOption.getOrElse(throw new Exception("No sheets"))
val modifiedSheet = firstSheet.put("A5", "Updated: " + LocalDate.now.toString).unsafe
println(s"  ✓ Modified sheet in-memory (${modifiedSheet.cells.size} cells)")
println()

// ========== Example 6: Error Handling ==========
println("⚠️  Example 6: Structured Errors")

try {
  Sheet("Test").unsafe.put("INVALID!!!!", "fail")
} catch {
  case ex: XLException =>
    println(s"  ✓ Caught: ${ex.getMessage.take(50)}...")
    println(s"  ✓ Error type: ${ex.error.getClass.getSimpleName}")
}
println()

println("✨ Easy Mode API: String refs, template styling, simplified IO!")
