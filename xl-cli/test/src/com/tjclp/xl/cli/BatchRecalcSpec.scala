package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Workbook, Sheet}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * GH-352: batch must end with one global recalculation so batch putf carries cached values (`<v>`)
 * exactly like single-op putf, formula errors surface in the summary without aborting the write,
 * presentation-only batches skip the recalculation, and the standalone `recalc` command freshens
 * any workbook.
 */
@SuppressWarnings(
  Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps")
)
class BatchRecalcSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  val config: WriterConfig = WriterConfig.default

  private def tempXlsx(): Path = Files.createTempFile("recalc-test", ".xlsx")

  private def writeOps(json: String): Path =
    val path = Files.createTempFile("batch-ops", ".json")
    Files.write(path, json.getBytes(StandardCharsets.UTF_8))
    path

  private def readBack(path: Path): Workbook =
    ExcelIO.instance[IO].read(path).unsafeRunSync()

  /** Raw bytes of a zip entry (e.g. `xl/worksheets/sheet2.xml`) inside an xlsx package. */
  private def zipEntryBytes(path: Path, entryName: String): Array[Byte] =
    val zip = new java.util.zip.ZipFile(path.toFile)
    try
      Option(zip.getEntry(entryName)) match
        case Some(entry) => zip.getInputStream(entry).readAllBytes()
        case None => fail(s"missing $entryName in $path")
    finally zip.close()

  /**
   * Marker comment that the writer never emits: it survives ONLY when the surgical writer copies
   * the worksheet part verbatim, so its presence distinguishes byte-for-byte preservation from a
   * (possibly byte-identical) regeneration.
   */
  private val preservationMarker = "<!--BatchRecalcSpec preservation marker-->"

  /** Rewrite one zip entry in place, leaving every other entry untouched. */
  private def rewriteZipEntry(path: Path, entryName: String)(
    transform: Array[Byte] => Array[Byte]
  ): Unit =
    import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}
    import scala.jdk.CollectionConverters.*
    val tmp = Files.createTempFile("zip-edit", ".xlsx")
    val zip = new ZipFile(path.toFile)
    try
      val out = new ZipOutputStream(Files.newOutputStream(tmp))
      try
        zip.entries().asScala.foreach { entry =>
          val bytes = zip.getInputStream(entry).readAllBytes()
          out.putNextEntry(new ZipEntry(entry.getName))
          out.write(if entry.getName == entryName then transform(bytes) else bytes)
          out.closeEntry()
        }
      finally out.close()
    finally zip.close()
    Files.move(tmp, path, java.nio.file.StandardCopyOption.REPLACE_EXISTING)

  /** Inject the preservation marker just before `</worksheet>` in a worksheet part. */
  private def injectPreservationMarker(path: Path, entryName: String): Unit =
    rewriteZipEntry(path, entryName) { bytes =>
      val xml = new String(bytes, StandardCharsets.UTF_8)
      assert(xml.contains("</worksheet>"), s"unexpected worksheet XML in $entryName")
      xml
        .replace("</worksheet>", s"$preservationMarker</worksheet>")
        .getBytes(StandardCharsets.UTF_8)
    }

  /** Cached value of the formula cell at (col0, row0) on the first sheet, failing otherwise. */
  private def cachedFormulaValue(wb: Workbook, col0: Int, row0: Int): Option[CellValue] =
    wb.sheets.head.cells.get(ARef.from0(col0, row0)).map(_.value) match
      case Some(CellValue.Formula(_, cached)) => cached
      case other => fail(s"Expected Formula cell, got $other")

  private def assertCachedNumber(cached: Option[CellValue], expected: Double): Unit =
    cached match
      case Some(CellValue.Number(n)) => assertEquals(n.toDouble, expected)
      case other => fail(s"Expected cached Number($expected), got $other")

  test("batch put+putf caches the formula value (GH-352 repro)") {
    val wb = Workbook(Sheet("Data"))
    val ops = writeOps("""[
      {"op":"put","ref":"A1","value":5},
      {"op":"putf","ref":"C1","value":"=A1*2"}
    ]""")
    val out = tempXlsx()

    val result = WriteCommands
      .batch(wb, Some(wb.sheets.head), ops.toString, out, config)
      .unsafeRunSync()

    assert(result.contains("Applied 2 operations"))
    assert(result.contains("Recalculated 1 formula"))

    assertCachedNumber(cachedFormulaValue(readBack(out), 2, 0), 10.0)
    Files.deleteIfExists(out)
    Files.deleteIfExists(ops)
  }

  test("batch putf and single-op putf converge on the same cached value") {
    val wb = Workbook(Sheet("Data").put(ARef.from0(0, 0), CellValue.Number(BigDecimal(5))))

    val singleOut = tempXlsx()
    WriteCommands
      .putFormula(wb, Some(wb.sheets.head), "C1", List("=A1*2"), singleOut, config)
      .unsafeRunSync()

    val batchOut = tempXlsx()
    val ops = writeOps("""[{"op":"putf","ref":"C1","value":"=A1*2"}]""")
    WriteCommands
      .batch(wb, Some(wb.sheets.head), ops.toString, batchOut, config)
      .unsafeRunSync()

    val singleCached = cachedFormulaValue(readBack(singleOut), 2, 0)
    val batchCached = cachedFormulaValue(readBack(batchOut), 2, 0)
    assertCachedNumber(singleCached, 10.0)
    assertEquals(batchCached, singleCached)
    Files.deleteIfExists(singleOut)
    Files.deleteIfExists(batchOut)
    Files.deleteIfExists(ops)
  }

  test("batch with a formula error still writes and reports it in the summary") {
    val wb = Workbook(Sheet("Data"))
    val ops = writeOps("""[
      {"op":"put","ref":"A1","value":5},
      {"op":"putf","ref":"B1","value":"=B1+1"},
      {"op":"putf","ref":"C1","value":"=A1*3"}
    ]""")
    val out = tempXlsx()

    val result = WriteCommands
      .batch(wb, Some(wb.sheets.head), ops.toString, out, config)
      .unsafeRunSync()

    // Write proceeded and the error is surfaced (count + ref), not thrown
    assert(result.contains("Saved:"))
    assert(result.contains("1 error"))
    assert(result.contains("B1"))
    assert(result.contains("Circular reference"))

    val imported = readBack(out)
    // The healthy formula still cached; the circular one stays uncached
    assertCachedNumber(cachedFormulaValue(imported, 2, 0), 15.0)
    assertEquals(cachedFormulaValue(imported, 1, 0), None)
    Files.deleteIfExists(out)
    Files.deleteIfExists(ops)
  }

  test("style-only batch does not recalculate (uncached formulas stay uncached)") {
    // Disk round-trip so the no-recalc path is pinned against the SourceContext /
    // ModificationTracker seam, not just the in-memory workbook.
    val srcFile = tempXlsx()
    ExcelIO
      .instance[IO]
      .write(
        Workbook(
          Sheet("Data")
            .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(5)))
            .put(ARef.from0(2, 0), CellValue.Formula("A1*2", None))
        ),
        srcFile
      )
      .unsafeRunSync()
    val wb = readBack(srcFile)
    val ops = writeOps("""[{"op":"style","range":"A1","bold":true}]""")
    val out = tempXlsx()

    val result = WriteCommands
      .batch(wb, Some(wb.sheets.head), ops.toString, out, config)
      .unsafeRunSync()

    assert(!result.contains("Recalculated"))
    assertEquals(cachedFormulaValue(readBack(out), 2, 0), None)
    Files.deleteIfExists(srcFile)
    Files.deleteIfExists(out)
    Files.deleteIfExists(ops)
  }

  test("styles-only clear does not recalculate (values untouched, presentation-only path)") {
    // BatchParser.applyClear only clears contents when `all || (!styles && !comments)`;
    // a styles-only clear must classify as non-mutating and skip the recalculation.
    val wb = Workbook(
      Sheet("Data")
        .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(5)))
        .put(ARef.from0(2, 0), CellValue.Formula("A1*2", None))
    )
    val ops = writeOps("""[{"op":"clear","range":"A1","styles":true}]""")
    val out = tempXlsx()

    val result = WriteCommands
      .batch(wb, Some(wb.sheets.head), ops.toString, out, config)
      .unsafeRunSync()

    assert(!result.contains("Recalculated"))
    val imported = readBack(out)
    // The cleared-styles cell keeps its value; the formula stays uncached
    assertEquals(
      imported.sheets.head.cells.get(ARef.from0(0, 0)).map(_.value),
      Some(CellValue.Number(BigDecimal(5)))
    )
    assertEquals(cachedFormulaValue(imported, 2, 0), None)
    Files.deleteIfExists(out)
    Files.deleteIfExists(ops)
  }

  test("contents clear still recalculates (default clear mutates cell values)") {
    val wb = Workbook(
      Sheet("Data")
        .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(5)))
        .put(ARef.from0(2, 0), CellValue.Formula("A1*2", None))
    )
    val ops = writeOps("""[{"op":"clear","range":"A1"}]""")
    val out = tempXlsx()

    val result = WriteCommands
      .batch(wb, Some(wb.sheets.head), ops.toString, out, config)
      .unsafeRunSync()

    assert(result.contains("Recalculated"))
    Files.deleteIfExists(out)
    Files.deleteIfExists(ops)
  }

  test("mutating batch preserves untouched formula-bearing sheets byte-for-byte") {
    // Regression pin for the GH-352 over-marking hazard: a batch edit on Inputs recalculates
    // the whole workbook, but Report's formula recomputes to the value it already cached, so
    // Report must NOT be marked modified — the surgical writer then preserves its worksheet
    // XML verbatim (which is also what keeps unparsed parts like pivot tables alive).
    val srcFile = tempXlsx()
    ExcelIO
      .instance[IO]
      .write(
        Workbook(
          Sheet("Inputs").put(ARef.from0(0, 0), CellValue.Number(BigDecimal(5))),
          Sheet("Report")
            .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(3)))
            .put(
              ARef.from0(1, 0),
              CellValue.Formula("A1*2", Some(CellValue.Number(BigDecimal(6))))
            )
        ),
        srcFile
      )
      .unsafeRunSync()
    // Stand-in for content our writer can't regenerate (pivot tables, slicers, ...): it
    // survives only if Report is never marked modified and rides the verbatim copy path.
    injectPreservationMarker(srcFile, "xl/worksheets/sheet2.xml")
    val wb = readBack(srcFile)

    // Edit Inputs!A1; no Report formula depends on it, so Report's caches are already correct
    val ops = writeOps("""[{"op":"put","ref":"Inputs!A1","value":7}]""")
    val out = tempXlsx()
    val result = WriteCommands
      .batch(wb, wb.sheets.headOption, ops.toString, out, config)
      .unsafeRunSync()
    assert(result.contains("Recalculated"))

    // Report (sheet2.xml) rides byte-for-byte preservation; the Inputs edit still lands
    assertEquals(
      zipEntryBytes(out, "xl/worksheets/sheet2.xml").toSeq,
      zipEntryBytes(srcFile, "xl/worksheets/sheet2.xml").toSeq
    )
    assert(
      new String(zipEntryBytes(out, "xl/worksheets/sheet2.xml"), StandardCharsets.UTF_8)
        .contains(preservationMarker),
      "Report worksheet was regenerated: preservation marker lost"
    )
    val imported = readBack(out)
    assertEquals(
      imported.sheets.head.cells.get(ARef.from0(0, 0)).map(_.value),
      Some(CellValue.Number(BigDecimal(7)))
    )
    val report = imported.sheets.find(_.name.value == "Report").getOrElse(fail("no Report"))
    report.cells.get(ARef.from0(1, 0)).map(_.value) match
      case Some(CellValue.Formula(_, Some(CellValue.Number(n)))) => assertEquals(n.toDouble, 6.0)
      case other => fail(s"Expected cached Report formula, got $other")
    Files.deleteIfExists(srcFile)
    Files.deleteIfExists(out)
    Files.deleteIfExists(ops)
  }

  test("recalc command caches a previously-uncached workbook (disk round-trip)") {
    // Round-trip through disk: workbooks read from a file carry a SourceContext whose
    // ModificationTracker gates the surgical writer — recalculated caches must survive it
    // (GH-352: recalculateImpl now reinstalls changed sheets via Workbook.put).
    val staleFile = tempXlsx()
    ExcelIO
      .instance[IO]
      .write(
        Workbook(
          Sheet("Data")
            .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(5)))
            .put(ARef.from0(2, 0), CellValue.Formula("A1*2", None))
        ),
        staleFile
      )
      .unsafeRunSync()
    val wb = readBack(staleFile)
    assertEquals(cachedFormulaValue(wb, 2, 0), None) // starts uncached on disk

    val out = tempXlsx()
    val result = WriteCommands.recalc(wb, out, config).unsafeRunSync()

    assert(result.contains("Recalculated 1 formula"))
    assertCachedNumber(cachedFormulaValue(readBack(out), 2, 0), 10.0)
    Files.deleteIfExists(staleFile)
    Files.deleteIfExists(out)
  }

  test("recalc on an already-fresh workbook preserves worksheet XML byte-for-byte") {
    // Every recomputed value equals its existing cache, so no sheet may be marked modified
    // and the surgical writer must copy the worksheet part verbatim.
    val srcFile = tempXlsx()
    ExcelIO
      .instance[IO]
      .write(
        Workbook(
          Sheet("Data")
            .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(5)))
            .put(
              ARef.from0(2, 0),
              CellValue.Formula("A1*2", Some(CellValue.Number(BigDecimal(10))))
            )
        ),
        srcFile
      )
      .unsafeRunSync()
    injectPreservationMarker(srcFile, "xl/worksheets/sheet1.xml")
    val wb = readBack(srcFile)

    val out = tempXlsx()
    val result = WriteCommands.recalc(wb, out, config).unsafeRunSync()

    assert(result.contains("Recalculated 1 formula"))
    assertEquals(
      zipEntryBytes(out, "xl/worksheets/sheet1.xml").toSeq,
      zipEntryBytes(srcFile, "xl/worksheets/sheet1.xml").toSeq
    )
    assert(
      new String(zipEntryBytes(out, "xl/worksheets/sheet1.xml"), StandardCharsets.UTF_8)
        .contains(preservationMarker),
      "worksheet was regenerated despite fresh caches: preservation marker lost"
    )
    Files.deleteIfExists(srcFile)
    Files.deleteIfExists(out)
  }

  test("batch refreshes cross-sheet dependent caches read from disk") {
    // A batch edit on Data must refresh the cached value of an Other!A1 formula that
    // depends on it — including through the surgical writer's preservation logic.
    val srcFile = tempXlsx()
    ExcelIO
      .instance[IO]
      .write(
        Workbook(
          Sheet("Data").put(ARef.from0(0, 0), CellValue.Number(BigDecimal(5))),
          Sheet("Other").put(ARef.from0(0, 0), CellValue.Formula("Data!A1*10", None))
        ),
        srcFile
      )
      .unsafeRunSync()
    val wb = readBack(srcFile)

    val ops = writeOps("""[{"op":"put","ref":"Data!A1","value":7}]""")
    val out = tempXlsx()
    WriteCommands.batch(wb, wb.sheets.headOption, ops.toString, out, config).unsafeRunSync()

    val imported = readBack(out)
    val otherSheet = imported.sheets.find(_.name.value == "Other").getOrElse(fail("no Other"))
    otherSheet.cells.get(ARef.from0(0, 0)).map(_.value) match
      case Some(CellValue.Formula(_, Some(CellValue.Number(n)))) => assertEquals(n.toDouble, 70.0)
      case other => fail(s"Expected cached cross-sheet formula, got $other")
    Files.deleteIfExists(srcFile)
    Files.deleteIfExists(out)
    Files.deleteIfExists(ops)
  }

  test("recalc command reports formula errors without failing (exit stays clean)") {
    val wb = Workbook(
      Sheet("Data")
        .put(ARef.from0(0, 0), CellValue.Formula("A1+1", None)) // self-circular
        .put(ARef.from0(2, 0), CellValue.Formula("41+1", None))
    )
    val out = tempXlsx()

    // Must not raise: formula errors are data conditions, not tool failures
    val result = WriteCommands.recalc(wb, out, config).unsafeRunSync()

    assert(result.contains("1 error"))
    assert(result.contains("Circular reference"))
    assert(result.contains("Saved:"))

    val imported = readBack(out)
    assertCachedNumber(cachedFormulaValue(imported, 2, 0), 42.0)
    assertEquals(cachedFormulaValue(imported, 0, 0), None)
    Files.deleteIfExists(out)
  }
