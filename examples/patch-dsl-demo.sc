//> using scala 3.7.3
//> using dep com.tjclp::xl-core:0.1.0-SNAPSHOT
//> using dep com.tjclp::xl-cats-effect:0.1.0-SNAPSHOT
//> using repository ivy2Local

// Standalone demo script - run with:
//   1. Publish locally: ./mill xl-core.publishLocal xl-cats-effect.publishLocal
//   2. Run script: scala-cli run examples/patch-dsl-demo.sc

import com.tjclp.xl.*  // Unified import includes DSL, macros, domain model
import com.tjclp.xl.io.ExcelIO
import cats.effect.{IO, IOApp}
import cats.effect.unsafe.implicits.global  // Demo only - use IOApp in production
import java.nio.file.Paths

println("=== XL Patch DSL Demo ===\n")

// Create a base sheet
val sheetResult = Sheet("Demo")

sheetResult match
  case Right(sheet) =>
    println("✓ Created empty sheet\n")

    // === OLD WAY: Verbose with type ascription ===
    println("=== OLD WAY (Verbose, requires type ascription with Cats) ===")
    println("""import cats.syntax.all.*
      |
      |val patch =
      |  (Patch.Put(ref"A1", CellValue.Text("Hello")): Patch) |+|
      |  (Patch.Put(ref"B1", CellValue.Number(42)): Patch) |+|
      |  (Patch.Put(ref"C1", CellValue.Text("World")): Patch)
      |""".stripMargin)

    // === NEW WAY: Clean DSL syntax ===
    println("\n=== NEW WAY (Clean DSL syntax, no type ascription!) ===")
    val cleanPatch =
      (ref"A1" := "Hello") ++
      (ref"B1" := 42) ++
      (ref"C1" := "World")

    println("""import com.tjclp.xl.*
      |
      |val patch =
      |  (ref"A1" := "Hello") ++
      |  (ref"B1" := 42) ++
      |  (ref"C1" := "World")
      |""".stripMargin)

    // Apply the clean patch
    Patch.applyPatch(sheet, cleanPatch) match
      case Right(updated) =>
        println(s"\n✓ Applied patch! Sheet now has ${updated.cellCount} cells:")
        updated.nonEmptyCells.foreach { cell =>
          println(s"  ${cell.toA1}: ${cell.value}")
        }
      case Left(err) =>
        println(s"✗ Error: ${err.message}")

    // === Complex patch with styles and merges ===
    println("\n=== Complex Patch Example ===")

    // Create a header style with fluent DSL (compare to verbose constructor!)
    val headerStyle = CellStyle.default.bold.size(14.0).white.bgBlue.center.middle

    val complexPatch =
      (ref"A1" := "Product Report") ++
      (ref"A1".styled(headerStyle)) ++
      ref"A1:C1".merge ++
      (ref"A3" := "Product") ++
      (ref"B3" := "Price") ++
      (ref"C3" := "Quantity") ++
      (ref"A4" := "Widget") ++
      (ref"B4" := 19.99) ++
      (ref"C4" := 100)

    println("""val patch =
      |  (ref"A1" := "Product Report") ++
      |  (ref"A1".styled(headerStyle)) ++
      |  ref"A1:C1".merge ++
      |  (ref"A3" := "Product") ++
      |  (ref"B3" := "Price") ++
      |  (ref"C3" := "Quantity") ++
      |  (ref"A4" := "Widget") ++
      |  (ref"B4" := 19.99) ++
      |  (ref"C4" := 100)
      |""".stripMargin)

    val sheet2Result = for
      emptySheet <- Sheet("Report")
      populated <- Patch.applyPatch(emptySheet, complexPatch)
    yield populated

    sheet2Result match
      case Right(report) =>
        println(s"\n✓ Created report with ${report.cellCount} cells:")
        report.nonEmptyCells.foreach { cell =>
          val styleInfo = if cell.styleId.isDefined then " [styled]" else ""
          println(s"  ${cell.toA1}: ${cell.value}${styleInfo}")
        }

        val mergedRanges = report.mergedRanges
        if mergedRanges.nonEmpty then
          println(s"\nMerged ranges:")
          mergedRanges.foreach(r => println(s"  ${r.toA1}"))
      case Left(err) =>
        println(s"✗ Error: ${err.message}")

    // === Range operations ===
    println("\n=== Range Operations ===")

    val rangePatch =
      (ref"A10" := "Start") ++
      (ref"B10" := "End") ++
      ref"A10:B10".merge ++
      ref"A12:C14".remove

    println("""val patch =
      |  (ref"A10" := "Start") ++
      |  (ref"B10" := "End") ++
      |  ref"A10:B10".merge ++
      |  ref"A12:C14".remove  // Remove entire range
      |""".stripMargin)

    // === Batch operations ===
    println("\n=== Batch Patch Construction ===")

    val batchPatch = PatchBatch(
      ref"E1" := "Batch",
      ref"E2" := "Operations",
      ref"E3" := "Rock!",
      ref"E1:E3".merge
    )

    println("""val patch = PatchBatch(
      |  ref"E1" := "Batch",
      |  ref"E2" := "Operations",
      |  ref"E3" := "Rock!",
      |  ref"E1:E3".merge
      |)
      |""".stripMargin)

    println("\n✓ All patches composed without type ascription!")

    // === Write to Excel file ===
    println("\n=== Write to Excel File ===")

    // Build a complete workbook with the complex patch
    val finalWorkbookResult = for
      reportSheet <- Sheet("Sales Report")
      populatedSheet <- Patch.applyPatch(reportSheet, complexPatch)
    yield Workbook(Vector(populatedSheet))

    finalWorkbookResult match
      case Right(workbook) =>
        val outputPath = Paths.get("patch-demo-output.xlsx")
        val excel = ExcelIO.instance[IO]

        println(s"Writing workbook to: ${outputPath.toAbsolutePath}")

        // Write the file using Cats Effect IO
        val writeResult = excel.write(workbook, outputPath).attempt.unsafeRunSync()

        writeResult match
          case Right(_) =>
            println(s"✓ Successfully wrote Excel file!")
            println(s"  Location: ${outputPath.toAbsolutePath}")
            println(s"  Sheets: ${workbook.sheets.size}")
            println(s"  Cells: ${workbook.sheets.head.cellCount}")
          case Left(err) =>
            println(s"✗ Error writing file: ${err.getMessage}")

      case Left(err) =>
        println(s"✗ Error creating workbook: ${err.message}")

  case Left(err) =>
    println(s"✗ Error creating sheet: ${err.message}")

println("\n=== Key Takeaways ===")
println("1. Import com.tjclp.xl.* (unified import includes DSL, macros, domain model)")
println("2. Use := operator: ref\"A1\" := \"value\" (auto-converts types)")
println("3. Use ++ operator to combine patches (no type ascription needed!)")
println("4. Extension methods: .styled(style), .merge, .unmerge, .remove")
println("5. PatchBatch(patches*) for varargs construction")
println("6. ExcelIO.instance[IO] for file I/O (pure effect system)")
println("\n=== Demo Complete ===")
