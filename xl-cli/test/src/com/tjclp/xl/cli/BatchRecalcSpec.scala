package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{CellRange, Workbook, Sheet, given}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.sheets.dataTableSyntax.*

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
      case Some(CellValue.Formula(_, cached, _)) => cached
      case other => fail(s"Expected Formula cell, got $other")

  private def assertCachedNumber(cached: Option[CellValue], expected: Double): Unit =
    cached match
      case Some(CellValue.Number(n)) => assertEquals(n.toDouble, expected)
      case other => fail(s"Expected cached Number($expected), got $other")

  test("GH-344: recalc summary counts error VALUES and the error cell round-trips the file") {
    val wb = Workbook(
      Sheet("Data")
        .put(ARef.from0(0, 0), CellValue.Number(BigDecimal(5)))
        .put(ARef.from0(2, 0), CellValue.Formula("1/0", None))
    )
    val out = tempXlsx()
    val result = WriteCommands.recalc(wb, out, config).unsafeRunSync()

    // Error values are data conditions: the run is clean, the summary reports the count
    assert(result.contains("Recalculated 1 formula (1 error value)"), s"summary was: $result")

    // The cached #DIV/0! survives the write/read round-trip (t="e" cell)
    cachedFormulaValue(readBack(out), 2, 0) match
      case Some(CellValue.Error(err)) => assertEquals(err.toExcel, "#DIV/0!")
      case other => fail(s"expected cached #DIV/0!, got $other")
    Files.deleteIfExists(out)
  }

  test("GH-344: recalc summary reports host failures and error values side by side") {
    val wb = Workbook(
      Sheet("Data")
        .put(ARef.from0(0, 0), CellValue.Formula("NOSUCHFN(1)", None))
        .put(ARef.from0(1, 0), CellValue.Formula("1/0", None))
    )
    val out = tempXlsx()
    val result = WriteCommands.recalc(wb, out, config).unsafeRunSync()
    assert(result.contains("(1 error value)"), s"summary was: $result")
    assert(result.contains("1 error ("), s"summary was: $result")
    Files.deleteIfExists(out)
  }

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
      case Some(CellValue.Formula(_, Some(CellValue.Number(n)), _)) => assertEquals(n.toDouble, 6.0)
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
      case Some(CellValue.Formula(_, Some(CellValue.Number(n)), _)) =>
        assertEquals(n.toDouble, 70.0)
      case other => fail(s"Expected cached cross-sheet formula, got $other")
    Files.deleteIfExists(srcFile)
    Files.deleteIfExists(out)
    Files.deleteIfExists(ops)
  }

  // ===== GH-442: `recalc --tables` seeds data-table interiors =====

  /**
   * A 2-D what-if table over D5:F6 with NO seeds: corner formula C4 = B1*B2 over base-case inputs
   * B1=5/B2=7, row axis D4:F4 substituting into B1, column axis C5:C6 substituting into B2. Seeded
   * interiors come out as the axis product (D5 = 1*10 ... F6 = 3*20) and the base inputs are left
   * alone; unseeded the record carries no cache and the plain interiors are never materialized.
   */
  private def unseededDataTableWorkbook(): Workbook =
    val interior = CellRange.parse("D5:F6").fold(err => fail(err), identity)
    val authored = Sheet("Data")
      .put(ref"B1" -> 5, ref"B2" -> 7)
      .put(ref"C4", CellValue.Formula("B1*B2"))
      .put(ref"D4" -> 1, ref"E4" -> 2, ref"F4" -> 3)
      .put(ref"C5" -> 10, ref"C6" -> 20)
      .dataTable(interior, ref"B1", ref"B2")
      .fold(err => fail(s"authoring failed: ${err.message}"), identity)
    Workbook(authored)

  /** Interior value as an Int, reading a record cell through its cache. */
  private def interiorInt(wb: Workbook, a1: ARef): Option[Int] =
    wb.sheets.head.cells.get(a1).map(_.value).flatMap {
      case CellValue.Number(n) => Some(n.toInt)
      case CellValue.Formula(_, Some(CellValue.Number(n)), _) => Some(n.toInt)
      case _ => None
    }

  test("GH-442: recalc --tables refreshes every data-table interior cache") {
    val out = tempXlsx()
    val summary =
      WriteCommands.recalc(unseededDataTableWorkbook(), out, config, false, true).unsafeRunSync()
    val seeded = readBack(out)
    assertEquals(
      interiorInt(seeded, ref"D5"),
      Some(10)
    ) // record cell keeps its kind, gains a cache
    assertEquals(interiorInt(seeded, ref"E5"), Some(20))
    assertEquals(interiorInt(seeded, ref"F5"), Some(30))
    assertEquals(interiorInt(seeded, ref"D6"), Some(20))
    assertEquals(interiorInt(seeded, ref"E6"), Some(40))
    assertEquals(interiorInt(seeded, ref"F6"), Some(60))
    // The what-if substitution is scoped to the seed: the base-case inputs survive untouched.
    assertEquals(interiorInt(seeded, ref"B1"), Some(5))
    assertEquals(interiorInt(seeded, ref"C4"), Some(35))
    assert(summary.contains("data table"), s"summary must report the seeding: $summary")
    Files.deleteIfExists(out)
  }

  test("GH-442: default recalc leaves data-table caches pinned (GH-353/GH-430 semantics)") {
    val out = tempXlsx()
    val summary = WriteCommands.recalc(unseededDataTableWorkbook(), out, config).unsafeRunSync()
    val pinned = readBack(out)
    // The record's cache stays absent and the plain interiors are never materialized.
    assertEquals(interiorInt(pinned, ref"D5"), None)
    assertEquals(interiorInt(pinned, ref"E5"), None)
    assertEquals(interiorInt(pinned, ref"F6"), None)
    // The ordinary corner formula still recalculates on the default path.
    assertEquals(interiorInt(pinned, ref"C4"), Some(35))
    assert(!summary.contains("data table"), s"default recalc must not mention seeding: $summary")
    Files.deleteIfExists(out)
  }

  test("GH-442: recalc --tables parses to Recalc(tables = true); bare recalc stays false") {
    val cmd = com.monovore.decline.Command("xl", "test")(Main.recalcCmd)
    cmd.parse(Seq("recalc", "--tables"), Map.empty) match
      case Right(CliCommand.Recalc(tables)) => assert(tables, "--tables must set tables = true")
      case other => fail(s"expected Recalc(true), got $other")
    cmd.parse(Seq("recalc"), Map.empty) match
      case Right(CliCommand.Recalc(tables)) =>
        assert(!tables, "bare recalc must keep the pinned-cache default")
      case other => fail(s"expected Recalc(false), got $other")
  }

  // ===== GH-453/GH-454: recalc honors the file's declared calcPr; --tables on circular books =====

  /**
   * Canonical circular fixture (GH-453): C1=100, B1=C1+B2, B2=$A$1*(C1+B1)/2 over rate input A1,
   * corner F9=IFERROR(B1/2,0), 1-var column table F10:F12 with left axis E10:E12 =
   * 0.002/0.004/0.006. Analytic per-axis interiors 50.1001/50.2004/50.3009 (strictly increasing);
   * flat interiors are the GH-453 bug.
   */
  private def circularTableWorkbook(calcPr: Option[com.tjclp.xl.workbooks.CalcPr]): Workbook =
    val kind: com.tjclp.xl.cells.FormulaKind.DataTable = com.tjclp.xl.cells.FormulaKind.DataTable(
      ref = CellRange.parse("F10:F12").fold(err => fail(err), identity),
      dt2D = false,
      dtr = false,
      r1 = Some(ref"A1"),
      r2 = None
    )
    val sheet = Sheet("Data")
      .put(ref"A1" -> 0, ref"C1" -> 100)
      .put(ref"B1", CellValue.Formula("C1+B2"))
      .put(ref"B2", CellValue.Formula("$A$1*(C1+B1)/2"))
      .put(ref"F9", CellValue.Formula("IFERROR(B1/2,0)"))
      .put(ref"E10" -> 0.002, ref"E11" -> 0.004, ref"E12" -> 0.006)
      .put(ref"F10", CellValue.dataTable(kind, None))
    calcPr.fold(Workbook(sheet))(cp => Workbook(sheet).withCalcPr(cp))

  private val iterateTight = com.tjclp.xl.workbooks.CalcPr(
    iterativeCalculation = true,
    Some(200),
    Some(BigDecimal("0.0000001"))
  )

  /** Interior value as BigDecimal, reading plain values or record caches. */
  private def interiorDec(wb: Workbook, a1: ARef): Option[BigDecimal] =
    wb.sheets.head.cells.get(a1).map(_.value).flatMap {
      case CellValue.Number(n) => Some(n)
      case CellValue.Formula(_, Some(CellValue.Number(n)), _) => Some(n)
      case _ => None
    }

  test("GH-453: recalc --tables on an iterate-declared circular book seeds varied interiors") {
    val out = tempXlsx()
    val summary = WriteCommands
      .recalc(circularTableWorkbook(Some(iterateTight)), out, config, false, true)
      .unsafeRunSync()
    // GH-454 verification instrument: the declared settings were honored and converged.
    assert(summary.contains("converged in"), s"summary: $summary")
    assert(!summary.contains("Circular reference"), s"summary: $summary")
    assert(!summary.contains("WARNING"), s"summary: $summary")
    val seeded = readBack(out)
    val values =
      Vector(ref"F10", ref"F11", ref"F12").map(r =>
        interiorDec(seeded, r).getOrElse(fail(s"${r.toA1} must seed"))
      )
    assertEquals(
      values.distinct.size,
      3,
      s"interiors must vary per axis — flat interiors are the GH-453 bug: $values"
    )
    values.zip(Vector(50.1001001, 50.2004008, 50.3009027)).foreach { (v, expected) =>
      assert((v.toDouble - expected).abs < 1e-4, s"interior $v, expected ~$expected")
    }
    Files.deleteIfExists(out)
  }

  test("GH-453: recalc without iterate declared keeps the circular-error posture") {
    val out = tempXlsx()
    val summary =
      WriteCommands.recalc(circularTableWorkbook(None), out, config).unsafeRunSync()
    assert(summary.contains("Circular reference"), s"summary: $summary")
    assert(!summary.contains("converged in"), s"summary: $summary")
    Files.deleteIfExists(out)
  }

  test("GH-453: recalc --tables without iterate warns and leaves the circular table unseeded") {
    val out = tempXlsx()
    val summary = WriteCommands
      .recalc(circularTableWorkbook(None), out, config, false, true)
      .unsafeRunSync()
    assert(summary.contains("depends on a circular reference"), s"summary: $summary")
    assert(summary.contains("left unseeded"), s"summary: $summary")
    assert(summary.contains("Saved:"), s"exit must stay clean: $summary")
    val written = readBack(out)
    assertEquals(interiorDec(written, ref"F10"), None, "record cache must stay absent")
    assertEquals(interiorDec(written, ref"F11"), None)
    assertEquals(interiorDec(written, ref"F12"), None)
    Files.deleteIfExists(out)
  }

  test("GH-453: recalc --tables renders the Skipped warning for unseedable interiors") {
    // The axis value feeding F10 references a missing sheet: the seeder leaves that interior
    // unseeded and reports one Skipped — the summary must surface it instead of reading clean.
    val base = circularTableWorkbook(Some(iterateTight))
    val broken = base.sheets.headOption.getOrElse(fail("missing sheet"))
    val wb = base.put(broken.put(ref"E10", CellValue.Formula("Missing!A1")))
    val out = tempXlsx()
    val summary = WriteCommands.recalc(wb, out, config, false, true).unsafeRunSync()
    assert(summary.contains("1 interior cell(s) left unseeded"), s"summary: $summary")
    assert(summary.contains("Saved:"), s"exit must stay clean: $summary")
    Files.deleteIfExists(out)
  }

  test("GH-454: exhausted iterative declaration renders the non-convergence WARNING") {
    val wb = Workbook(
      Sheet("Data")
        .put(ref"B1", CellValue.Formula("1-B2"))
        .put(ref"B2", CellValue.Formula("1-B1"))
    ).withCalcPr(com.tjclp.xl.workbooks.CalcPr(iterativeCalculation = true, Some(5), None))
    val out = tempXlsx()
    val summary = WriteCommands.recalc(wb, out, config).unsafeRunSync()
    assert(
      summary.contains(
        "WARNING: iterative calculation exhausted 5 round(s) without converging"
      ),
      s"summary: $summary"
    )
    Files.deleteIfExists(out)
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
