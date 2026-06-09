#!/usr/bin/env -S scala-cli shebang
//> using file project.scala


// Standalone demo script - run with:
//   1. Publish locally: ./mill xl-core.publishLocal && ./mill xl-cats-effect.publishLocal
//   2. Run script: scala-cli run examples/easy_mode_demo.sc

import com.tjclp.xl.scripting.{*, given} // One-import scripting prelude (core + IO + evaluator + unsafe)
import java.time.LocalDate

println("🚀 XL Easy Mode API Demo\n")

// ========== Example 1: String-Based References ==========
println("📋 Example 1: String-Based References (Easy Mode)")

val headerStyle = CellStyle.default.bold.size(12.0).center
val titleStyle = CellStyle.default.bold.size(14.0).bgBlue.white

val report = Sheet("Q1 Report")
  .style("A1:D1", titleStyle)           // ✨ Clean chainable API!
  .style("A2:D2", headerStyle)
  .put("A1", "Q1 2025 Sales Report")
  .put("A2", "Product")
  .put("B2", "Units")
  .put("A3", "Widget")
  .put("B3", 150)

println(s"  ✓ Used string refs for ${report.cells.size} cells")
println()

// ========== Example 2: Inline Styling ==========
println("🎨 Example 2: Inline Styling")

val quickSheet = Sheet("Quick")
  .put("A1", "Alert", CellStyle.default.bold.red)
  .put("A2", "Success", CellStyle.default.bold.green)

println(s"  ✓ Applied inline styles")
println()

// ========== Example 3: Rich Text ==========
println("💬 Example 3: Rich Text")

val richSheet = Sheet("RichText")
  .put("A1", "Status: ".bold + "ACTIVE".green.bold)
  .put("A2", "Error: ".red.bold + "Fix immediately!")

println(s"  ✓ Created rich text cells")
println()

// ========== Example 4: Patch DSL (Declarative Alternative) ==========
println("🔧 Example 4: Patch DSL (Declarative Approach)")

val boldStyle = CellStyle.default.bold
val patchSheet = Sheet("Patch Demo")
  .put(
    (ref"A1" := "Product Report") ++
    ref"A1".styled(headerStyle) ++
    ref"A1:C1".merge ++
    (ref"A3" := "Product") ++
    (ref"B3" := "Price") ++
    (ref"C3" := "Quantity") ++
    (ref"A4" := "Widget") ++
    (ref"B4" := 19.99) ++
    (ref"C4" := 100)
  )

println(s"  ✓ Built sheet with Patch DSL (${patchSheet.cells.size} cells, ${patchSheet.mergedRanges.size} merge)")
println()

// ========== Example 5: Dynamic Patch with String Interpolation ==========
println("🔥 Example 5: Dynamic Patch Generation (Runtime Refs)")

// Sample data: products to populate
val products = List(
  ("Widget", 19.99, 100),
  ("Gadget", 29.99, 50),
  ("Doohickey", 39.99, 25)
)

// Build patch by folding over data with interpolated refs + conditional styling
val dynamicPatch = products.zipWithIndex.foldLeft(Patch.empty) { case (acc, ((name, price, qty), idx)) =>
  val row = (idx + 4).toString  // Start at row 4

  // Conditional styling: green if price > 25, red otherwise
  val priceStyle = if (price > 25.0)
    CellStyle.default.bold.green
  else
    CellStyle.default.bold.red

  // Runtime interpolation returns Either[XLError, RefType]
  // RefType now supports := operator directly!
  (for {
    nameRef <- ref"A$row"
    priceRef <- ref"B$row"
    qtyRef <- ref"C$row"
  } yield
    acc ++
    (nameRef := name) ++      // Works directly with RefType!
    (priceRef := price) ++
    (priceRef.styled(priceStyle)) ++  // Conditional styling!
    (qtyRef := qty)
  ).getOrElse(acc)  // Graceful fallback on parse error
}

val dynamicSheet = Sheet("Dynamic")
  .put(ref"A3" := "Product")
  .put(ref"B3" := "Price")
  .put(ref"C3" := "Quantity")
  .put(dynamicPatch)

println(s"  ✓ Generated ${products.size} rows dynamically with interpolated refs")
println(s"  ✓ Applied conditional styling (green if price > $$25, red otherwise)")
println(s"  ✓ Total cells: ${dynamicSheet.cells.size}")

// Export to HTML to visualize the conditional styling
val htmlOutput = dynamicSheet.toHtml(ref"A3:C6")
println(s"  ✓ Exported to HTML (${htmlOutput.length} chars)")
println(s"\nHTML Preview (with conditional styling):")
println(htmlOutput)
println()

// ========== Example 6: Comments & Annotations ==========
println("💬 Example 6: Comments & Annotations")

val commentedSheet = Sheet("Commented Data")
  .put("A1", "Q4 Revenue")
  .put("B1", 125000)
  .put("A2", "Q4 Expenses")
  .put("B2", 87500)
  .put("A3", "Net Profit")
  .put("B3", "=B1-B2")  // Formula
  // Add plain text comments
  .comment(ref"B1", Comment.plainText("Revenue increased by 15% vs Q3", Some("Finance Team")))
  .comment(ref"B2", Comment.plainText("Includes one-time marketing spend", Some("CFO")))
  // Rich text comment with formatting
  .comment(ref"B3", Comment(
    text = "Net profit: ".bold + "$37,500".green.bold + " (30% margin)".italic,
    author = Some("CEO")
  ))

println(s"  ✓ Created sheet with ${commentedSheet.comments.size} comments")
commentedSheet.comments.foreach { (ref, comment) =>
  val author = comment.author.map(a => s"[$a]").getOrElse("[Anonymous]")
  println(s"  ✓ ${ref.toA1} $author: ${comment.text.toPlainText.take(40)}...")
}

// Export to HTML with comments as tooltips
val htmlWithComments = commentedSheet.toHtml(ref"A1:B3", includeComments = true)
println(s"\n  ✓ HTML export with comments (${htmlWithComments.length} chars)")
println(s"  ℹ️  Hover over cells in browser to see comment tooltips!")
println("\nHTML Preview (with comment tooltips):")
println(htmlWithComments)
println()

// ========== Example 7: Safe Lookups ==========
println("🔍 Example 7: Safe Lookups")

val value = report.cell("A3")          // ✨ Clean lookup!
val range = report.range("A3:B3")      // ✨ Get cells in range!

println(s"  ✓ Looked up cell: ${value.map(_.value)}")
println(s"  ✓ Found ${range.size} cells in range")
println()

// ========== Example 8: Excel IO ==========
println("💾 Example 8: Excel IO")

val workbook = Workbook.empty
  .put(report)
  .put(quickSheet)
  .put(richSheet)
  .put(patchSheet)
  .put(dynamicSheet)
  .put(commentedSheet)
  .remove("Sheet1")  // Sheet1 is always created by default per Excel standards
  .unsafe  // Single unwrap at the end!

Excel.write(workbook, "/tmp/easy-mode-demo.xlsx")
println(s"  ✓ Wrote /tmp/easy-mode-demo.xlsx")

val loaded = Excel.read("/tmp/easy-mode-demo.xlsx")
println(s"  ✓ Read back (${loaded.sheets.size} sheets)")

// Demonstrate in-memory modification
val firstSheet = loaded.sheets.headOption.getOrElse(throw new Exception("No sheets"))
val modifiedSheet = firstSheet.put("A5", "Updated: " + LocalDate.now.toString)
println(s"  ✓ Modified sheet in-memory (${modifiedSheet.cells.size} cells)")
println()

// ========== Example 9: Error Handling ==========
println("⚠️  Example 9: Structured Errors")

// Note: Invalid literals like "INVALID!!!!" now fail at compile time!
// To demonstrate runtime error handling, we use a runtime string:
val invalidRef = "INVALID!!!!"  // Runtime string bypasses compile-time check
Sheet("Test").put(invalidRef, "fail") match {
  case Left(error) =>
    println(s"  ✓ Caught: ${error.message.take(50)}...")
    println(s"  ✓ Error type: ${error.getClass.getSimpleName}")
  case Right(_) =>
    println("  ✗ Should have failed!")
}
println()

println("✨ Easy Mode API: Chainable + Declarative + Dynamic (with interpolation)!")
