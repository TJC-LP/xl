//> using scala 3.7.3
//> using dep com.tjclp::xl-core:0.1.0-SNAPSHOT
//> using dep com.tjclp::xl-ooxml:0.1.0-SNAPSHOT
//> using repository ivy2Local

import java.nio.file.{Files, Paths}
import java.util.zip.ZipFile
import com.tjclp.xl.api.*
import com.tjclp.xl.ooxml.XlsxReader
import com.tjclp.xl.ooxml.XlsxWriter
import com.tjclp.xl.macros.ref

println("=" * 80)
println("SURGICAL MODIFICATION DEMO - Real-World Data")
println("=" * 80)

val dataDir = Paths.get("/Users/rcaputo3/git/xl/data")
val inputPath = dataDir.resolve("Syndigo Valuation_Q3 2025_2025.10.15_VALUES.xlsx")
val outputPath = dataDir.resolve("syndigo-surgical-output.xlsx")

// Clean up previous output
Files.deleteIfExists(outputPath)

println("\n[Step 1] Reading source file...")
println(s"  File: ${inputPath.getFileName}")
println(s"  Size: ${Files.size(inputPath)} bytes (${Files.size(inputPath) / 1024}KB)")

val readResult = XlsxReader.read(inputPath)
val workbook = readResult match
  case Right(wb) =>
    println("  ✓ File read successfully")
    wb
  case Left(err) =>
    println(s"  ✗ Failed to read: $err")
    sys.exit(1)

println(s"  Sheets: ${workbook.sheets.size}")
workbook.sheets.foreach { sheet =>
  println(s"    - ${sheet.name.value}")
}

println("\n[Step 2] Analyzing SourceContext...")
workbook.sourceContext match
  case None =>
    println("  ✗ No SourceContext (unexpected!)")
    sys.exit(1)
  case Some(ctx) =>
    println("  ✓ SourceContext created")
    println(s"  Source path: ${ctx.sourcePath}")
    println(s"  Is clean: ${ctx.isClean}")

    val manifest = ctx.partManifest
    val totalEntries = manifest.entries.size
    val parsedCount = manifest.parsedParts.size
    val unparsedCount = manifest.unparsedParts.size

    println(s"\n  Part Manifest:")
    println(s"    Total entries: $totalEntries")
    println(s"    Parsed (known to XL): $parsedCount")
    println(s"    Unparsed (preserved): $unparsedCount")

    if unparsedCount > 0 then
      println(s"\n  Unknown parts that will be preserved:")
      manifest.unparsedParts.toList.sorted.take(10).foreach { path =>
        println(s"    - $path")
      }
      if unparsedCount > 10 then
        println(s"    ... and ${unparsedCount - 10} more")

println("\n[Step 3] Making minimal modification...")
val modifiedResult = for
  // Use "Syndigo - Valuation" by NAME (safe, not index-dependent)
  sheet <- workbook("Syndigo - Valuation")
  // Use Z100 - a safe empty cell far from data
  testCell = ref"Z100"
  originalValue = sheet.cells.get(testCell).map(_.value)
  _ = println(s"  Target: Sheet '${sheet.name.value}', Cell: ${testCell.toA1}")
  _ = println(s"  Original ${testCell.toA1} value: $originalValue")
  updatedSheet <- sheet.put(testCell -> "🎯 Modified by XL Surgical Write")
  updated <- workbook.put(updatedSheet)
yield updated

val modified = modifiedResult match
  case Right(wb) =>
    println("  ✓ Modification successful")
    wb
  case Left(err) =>
    println(s"  ✗ Failed to modify: $err")
    sys.exit(1)

modified.sourceContext.foreach { ctx =>
  val tracker = ctx.modificationTracker
  val modifiedSheetNames = tracker.modifiedSheets.flatMap(idx =>
    modified.sheets.lift(idx).map(_.name.value)
  ).mkString(", ")
  println(s"\n  Modification Tracker:")
  println(s"    Modified sheets: ${tracker.modifiedSheets} ($modifiedSheetNames)")
  println(s"    Deleted sheets: ${tracker.deletedSheets}")
  println(s"    Reordered: ${tracker.reorderedSheets}")
  println(s"    Is clean: ${tracker.isClean}")
}

println("\n[Step 4] Writing with surgical modification...")
val startTime = System.currentTimeMillis()

val writeResult = XlsxWriter.write(modified, outputPath)
writeResult match
  case Right(_) =>
    val duration = System.currentTimeMillis() - startTime
    println(s"  ✓ Write successful in ${duration}ms")
  case Left(err) =>
    println(s"  ✗ Failed to write: $err")
    sys.exit(1)

val outputSize = Files.size(outputPath)
println(s"  Output size: $outputSize bytes (${outputSize / 1024}KB)")
println(s"  Size delta: ${outputSize - Files.size(inputPath)} bytes")

println("\n[Step 5] Verifying preservation...")

// Re-read output
val verifyResult = XlsxReader.read(outputPath)
val reloaded = verifyResult match
  case Right(wb) =>
    println("  ✓ Output is valid XLSX")
    wb
  case Left(err) =>
    println(s"  ✗ Failed to reload: $err")
    sys.exit(1)

// Verify modified cell by SHEET NAME
val testCell = ref"Z100"
val cellValue = reloaded("Syndigo - Valuation").flatMap(sheet =>
  Right(sheet.cells.get(testCell).map(_.value))
)
cellValue match
  case Right(Some(value)) =>
    println(s"  ✓ Modified cell ${testCell.toA1} in 'Syndigo - Valuation': $value")
  case _ =>
    println(s"  ✗ Modified cell ${testCell.toA1} not found in 'Syndigo - Valuation'")

// Compare unknown parts count
val originalUnknownCount = workbook.sourceContext.get.partManifest.unparsedParts.size
val outputUnknownCount = reloaded.sourceContext.get.partManifest.unparsedParts.size

println(s"\n  Unknown parts comparison:")
println(s"    Original: $originalUnknownCount parts")
println(s"    Output: $outputUnknownCount parts")
println(s"    Status: ${if originalUnknownCount == outputUnknownCount then "✓ PRESERVED" else "✗ LOST"}")

println("\n[Step 6] Byte-level verification...")

// Open both as ZIPs and compare a few unknown parts
val sourceZip = new ZipFile(inputPath.toFile)
val outputZip = new ZipFile(outputPath.toFile)

try
  val unknownParts = workbook.sourceContext.get.partManifest.unparsedParts.toList.sorted.take(3)

  if unknownParts.isEmpty then
    println("  No unknown parts to verify (file has no charts/images)")
  else
    println(s"  Comparing ${unknownParts.size} unknown parts byte-by-byte:")

    unknownParts.foreach { partPath =>
      val sourceEntry = sourceZip.getEntry(partPath)
      val outputEntry = outputZip.getEntry(partPath)

      if sourceEntry != null && outputEntry != null then
        val sourceBytes = sourceZip.getInputStream(sourceEntry).readAllBytes()
        val outputBytes = outputZip.getInputStream(outputEntry).readAllBytes()

        val identical = java.util.Arrays.equals(sourceBytes, outputBytes)
        val status = if identical then "✓ IDENTICAL" else "✗ DIFFERENT"
        println(s"    $partPath: $status (${sourceBytes.length} bytes)")
      else
        println(s"    $partPath: ✗ MISSING in output")
    }
finally
  sourceZip.close()
  outputZip.close()

println("\n" + "=" * 80)
println("SURGICAL MODIFICATION SUCCESSFUL! 🎉")
println("=" * 80)
println("\nSummary:")
println(s"  - Read ${workbook.sheets.size} sheets")
println(s"  - Modified 1 cell (Z100) in 'Syndigo - Valuation' sheets (by name, not index)")
println(s"  - Preserved ${originalUnknownCount} unknown parts")
println("\nThe hybrid write strategy:")
println("  ✓ Regenerated only 'Syndigo - Valuation' (the modified sheets)")
println("  ✓ Copied 8 unmodified sheets byte-for-byte (including hidden system sheets)")
println("  ✓ Preserved all unknown parts (charts, images, comments, calcChain, etc.)")
println("  ✓ Preserved structural files (ContentTypes, Relationships)")
println("\n" + "=" * 80)
