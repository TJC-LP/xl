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

  test("GH-493: an unresolvable cone reports the CONE, never 'interior cell(s) left unseeded'") {
    // Round-2 review: B1 is axis-dependent but its cross-sheet leg is unresolvable, so the cone
    // stays on its stale cache and the grid seeds FLAT (11 everywhere). Every interior IS seeded —
    // rendering that as "1 interior cell(s) left unseeded" is a false statement about the run.
    val kind: com.tjclp.xl.cells.FormulaKind.DataTable = com.tjclp.xl.cells.FormulaKind.DataTable(
      ref = CellRange.parse("D5:D7").fold(err => fail(err), identity),
      dt2D = false,
      dtr = false,
      r1 = Some(ref"A1"),
      r2 = None
    )
    val sheet = Sheet("Data")
      .put(ref"A1" -> 0)
      .put(ref"B1", CellValue.Formula("A1*10+Missing!A1", Some(CellValue.Number(BigDecimal(10)))))
      .put(ref"D4", CellValue.Formula("B1+1"))
      .put(ref"C5" -> 1, ref"C6" -> 2, ref"C7" -> 3)
      .put(ref"D5", CellValue.dataTable(kind, None))
    val out = tempXlsx()
    val summary =
      WriteCommands.recalc(Workbook(sheet), out, config, false, true).unsafeRunSync()
    assert(
      summary.contains("precedent cell(s) in the what-if cone could not be re-derived"),
      s"summary: $summary"
    )
    assert(
      !summary.contains("left unseeded"),
      s"the interiors WERE seeded — the warning must not claim otherwise: $summary"
    )
    assert(summary.contains("Data!B1"), s"the warning must name the cone cell: $summary")
    val written = readBack(out)
    Vector(ref"D5", ref"D6", ref"D7").foreach { r =>
      assert(interiorDec(written, r).isDefined, s"${r.toA1} was seeded: $summary")
    }
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

  // ===== GH-468: a write must not stomp caches outside its dirty dependency cone =====

  /**
   * Stand-in for the GH-468 field repro: caches authored out-of-band (an external engine, a
   * LibreOffice arbiter) that xl's own evaluator disagrees with. C1 caches 999 for `=A1*2` and F1
   * caches 888 for `=D1*2`; any whole-book recalculation rewrites both to 10/14 — the re-poisoning
   * signature. Two independent chains so a dirty-cone scope is observable: an edit to A1 may
   * refresh C1 and must still leave F1 alone.
   */
  private def splicedCacheWorkbook(): Workbook =
    Workbook(
      Sheet("Data")
        .put(ref"A1" -> 5, ref"D1" -> 7)
        .put(ref"C1", CellValue.Formula("A1*2", Some(CellValue.Number(BigDecimal(999)))))
        .put(ref"F1", CellValue.Formula("D1*2", Some(CellValue.Number(BigDecimal(888)))))
    )

  private def assertCached(wb: Workbook, at: ARef, expected: Double): Unit =
    wb.sheets.head.cells.get(at).map(_.value) match
      case Some(CellValue.Formula(_, Some(CellValue.Number(n)), _)) =>
        assertEquals(n.toDouble, expected, s"${at.toA1} cache")
      case other => fail(s"expected cached number at ${at.toA1}, got $other")

  private def assertCachedOn(wb: Workbook, sheet: String, at: ARef, expected: Double): Unit =
    wb.sheets.find(_.name.value == sheet).flatMap(_.cells.get(at)).map(_.value) match
      case Some(CellValue.Formula(_, Some(CellValue.Number(n)), _)) =>
        assertEquals(n.toDouble, expected, s"$sheet!${at.toA1} cache")
      case other => fail(s"expected cached number at $sheet!${at.toA1}, got $other")

  test("GH-468: an unrelated batch put leaves externally-authored caches untouched") {
    val wb = splicedCacheWorkbook()
    val ops = writeOps("""[{"op":"put","ref":"Z99","value":1}]""")
    val out = tempXlsx()

    WriteCommands.batch(wb, wb.sheets.headOption, ops.toString, out, config).unsafeRunSync()

    val written = readBack(out)
    assertCached(written, ref"C1", 999.0) // NOT the evaluator's 10
    assertCached(written, ref"F1", 888.0) // NOT the evaluator's 14
    assertEquals(
      written.sheets.head.cells.get(ref"Z99").map(_.value),
      Some(CellValue.Number(BigDecimal(1))),
      "the edit itself must still land"
    )
    Files.deleteIfExists(out)
    Files.deleteIfExists(ops)
  }

  test("GH-468: a batch edit refreshes its dirty cone and nothing else") {
    val wb = splicedCacheWorkbook()
    val ops = writeOps("""[{"op":"put","ref":"A1","value":6}]""")
    val out = tempXlsx()

    val summary =
      WriteCommands.batch(wb, wb.sheets.headOption, ops.toString, out, config).unsafeRunSync()

    val written = readBack(out)
    assertCached(written, ref"C1", 12.0) // A1's dependent: inside the cone, refreshed
    assertCached(written, ref"F1", 888.0) // untouched chain: outside the cone, preserved
    assert(summary.contains("Recalculated 1 formula"), s"summary: $summary")
    Files.deleteIfExists(out)
    Files.deleteIfExists(ops)
  }

  test("GH-468: insert-rows keeps the caches of a sheet that does not reference the edited one") {
    // Narrow claim, deliberately: `Other` is SELF-CONTAINED (Other!B1 = Other!A1*2), the only
    // shape whose caches survive a structural edit today. StructuralEditor drops the cache of
    // every formula that transitively READS the edited sheet (GH-455's `stale` predicate),
    // cross-sheet readers included — pinned by the next test, tracked as GH-503. What GH-468
    // fixes here is the trailing global recalculate(), which rewrote even THIS 888 cache to 14.
    val wb = Workbook(
      Sheet("Data")
        .put(ref"D10" -> 7)
        .put(ref"F10", CellValue.Formula("D10*2", None)),
      Sheet("Other")
        .put(ref"A1" -> 7)
        .put(ref"B1", CellValue.Formula("A1*2", Some(CellValue.Number(BigDecimal(888)))))
    )
    val out = tempXlsx()

    WriteCommands.insertRows(wb, wb.sheets.headOption, 5, 1, out, config).unsafeRunSync()

    val written = readBack(out)
    assertCached(written, ref"F11", 14.0) // shifted by the edit: recalculated
    val other = written.sheets.find(_.name.value == "Other").getOrElse(fail("no Other"))
    other.cells.get(ref"B1").map(_.value) match
      case Some(CellValue.Formula(_, Some(CellValue.Number(n)), _)) =>
        assertEquals(n.toDouble, 888.0, "an untouched sheet's cache must survive a structural edit")
      case other => fail(s"expected a cached formula at Other!B1, got $other")
    Files.deleteIfExists(out)
  }

  /**
   * Data lives in rows 1 and 6 only, and every formula caches an out-of-band 777. A structural edit
   * at row 20 / column AA therefore shifts NOTHING — no cell moves, no formula text is rewritten —
   * so every one of these caches is still the truth about the file afterwards. `Other!B1` reads the
   * edited sheet across the sheet boundary; `Data!C30` is the one cell that genuinely moves under
   * an insert at row 20 while keeping its text byte-identical (its reference is absolute and above
   * the band).
   */
  private def structuralPoisonWorkbook(): Workbook =
    Workbook(
      Sheet("Data")
        .put(ref"A1" -> 5)
        .put(ref"C1", CellValue.Formula("A1*2", Some(CellValue.Number(BigDecimal(777)))))
        .put(ref"S1", CellValue.Formula("SUM(A1:A10)", Some(CellValue.Number(BigDecimal(777)))))
        .put(ref"C30", CellValue.Formula("$A$1*3", Some(CellValue.Number(BigDecimal(777))))),
      Sheet("Other")
        .put(ref"B1", CellValue.Formula("Data!A1*2", Some(CellValue.Number(BigDecimal(777)))))
    )

  test("insert-rows over-invalidates a cross-sheet reader's cache (pins current behavior)") {
    // GH-503: StructuralEditor.rewriteFormulas' `stale` predicate is "transitively reads the
    // edited sheet", not "reads something the edit moved", so an insert at row 20 of a sheet
    // whose data stops at row 6 still drops these caches and the cone re-derives them. The
    // expected values below are WRONG on purpose — flip every one of them back to 777.0 when
    // GH-503 tightens the predicate.
    val wb = structuralPoisonWorkbook()
    val out = tempXlsx()

    WriteCommands.insertRows(wb, wb.sheets.headOption, 20, 1, out, config).unsafeRunSync()

    val written = readBack(out)
    assertCachedOn(written, "Data", ref"C1", 10.0) // GH-503: should be 777.0
    assertCachedOn(written, "Data", ref"S1", 5.0) // GH-503: should be 777.0
    assertCachedOn(written, "Other", ref"B1", 10.0) // GH-503: should be 777.0 (CROSS-SHEET)
    Files.deleteIfExists(out)
  }

  // ===== GH-481: batch honors the file's declared calcPr like recalc does =====

  /** C1=100, B1=C1+B2, B2=$A$1*(C1+B1)/2 over rate input A1 — a declared circular book. */
  private def circularRateWorkbook(calcPr: Option[com.tjclp.xl.workbooks.CalcPr]): Workbook =
    val sheet = Sheet("Data")
      .put(ref"A1" -> 0, ref"C1" -> 100)
      .put(ref"B1", CellValue.Formula("C1+B2"))
      .put(ref"B2", CellValue.Formula("$A$1*(C1+B1)/2"))
    calcPr.fold(Workbook(sheet))(cp => Workbook(sheet).withCalcPr(cp))

  test("GH-481: a batch edit to an iterate-declared circular book fixpoints instead of erroring") {
    val wb = circularRateWorkbook(Some(iterateTight))
    val ops = writeOps("""[{"op":"put","ref":"A1","value":0.01}]""")
    val out = tempXlsx()

    val summary =
      WriteCommands.batch(wb, wb.sheets.headOption, ops.toString, out, config).unsafeRunSync()

    assert(summary.contains("converged in"), s"summary: $summary")
    assert(!summary.contains("Circular reference"), s"summary: $summary")
    // B1 = 100 + 0.005*(100 + B1)  =>  B1 = 100.5 / 0.995 = 101.00502...
    readBack(out).sheets.head.cells.get(ref"B1").map(_.value) match
      case Some(CellValue.Formula(_, Some(CellValue.Number(n)), _)) =>
        assert((n.toDouble - 101.005).abs < 1e-3, s"B1 fixpoint was $n")
      case other => fail(s"expected a cached fixpoint at B1, got $other")
    Files.deleteIfExists(out)
    Files.deleteIfExists(ops)
  }

  // ===== GH-468: --no-recalc / --preserve-caches =====

  private val preserveCaches: WritePolicy = WritePolicy(noRecalc = true)
  private val strictPolicy: WritePolicy = WritePolicy(strict = true)

  test("GH-468: batch --no-recalc applies the edit and recalculates nothing") {
    val wb = splicedCacheWorkbook()
    val ops = writeOps("""[{"op":"put","ref":"A1","value":6}]""")
    val out = tempXlsx()

    val summary = WriteCommands
      .batch(wb, wb.sheets.headOption, ops.toString, out, config, false, preserveCaches)
      .unsafeRunSync()

    assert(summary.contains("Recalculation skipped (--no-recalc)"), s"summary: $summary")
    assert(!summary.contains("Recalculated"), s"summary: $summary")
    val written = readBack(out)
    assertCached(written, ref"C1", 999.0) // even the dirty cone is left alone
    assertCached(written, ref"F1", 888.0)
    assertEquals(
      written.sheets.head.cells.get(ref"A1").map(_.value),
      Some(CellValue.Number(BigDecimal(6))),
      "the edit itself must still land"
    )
    Files.deleteIfExists(out)
    Files.deleteIfExists(ops)
  }

  test("GH-468: put --no-recalc leaves the dependent's cache alone") {
    val wb = splicedCacheWorkbook()
    val out = tempXlsx()

    val summary = WriteCommands
      .put(wb, wb.sheets.headOption, "A1", List("6"), out, config, policy = preserveCaches)
      .unsafeRunSync()

    assert(summary.contains("Recalculation skipped (--no-recalc)"), s"summary: $summary")
    assertCached(readBack(out), ref"C1", 999.0)
    Files.deleteIfExists(out)
  }

  test("GH-468: put without the flag still refreshes its dependents (default unchanged)") {
    val wb = splicedCacheWorkbook()
    val out = tempXlsx()

    val summary =
      WriteCommands.put(wb, wb.sheets.headOption, "A1", List("6"), out, config).unsafeRunSync()

    assert(!summary.contains("--no-recalc"), s"summary: $summary")
    assertCached(readBack(out), ref"C1", 12.0)
    Files.deleteIfExists(out)
  }

  test("GH-468: --no-recalc on insert-rows preserves the caches the structural edit dropped") {
    // The escape hatch prints "every existing cached value preserved". StructuralEditor strips
    // the cache of every formula reading the edited sheet BEFORE the CLI sees the workbook, so
    // without carrying the pre-edit caches forward the flag would write formula cells with no
    // <v> at all — worse than the recalculation it replaces. Row 20 shifts nothing here.
    val wb = structuralPoisonWorkbook()
    val out = tempXlsx()

    val summary = WriteCommands
      .insertRows(wb, wb.sheets.headOption, 20, 1, out, config, false, preserveCaches)
      .unsafeRunSync()

    assert(summary.contains("Recalculation skipped (--no-recalc)"), s"summary: $summary")
    val written = readBack(out)
    assertCachedOn(written, "Data", ref"C1", 777.0)
    assertCachedOn(written, "Data", ref"S1", 777.0)
    assertCachedOn(written, "Other", ref"B1", 777.0)
    // C30 MOVED to C31. Text identity is not enough to re-assert a cache on a cell that relocated
    // (the very next tests show =ROW()/=COLUMN() lying under exactly this shape), so the moved
    // cell is left uncached and the summary counts it.
    assertEquals(formulaOn(written, "Data", ref"C31").cachedValue, None)
    assert(summary.contains("3 cached value(s) preserved"), s"summary: $summary")
    assert(summary.contains("1 formula(s)"), s"summary: $summary")
    Files.deleteIfExists(out)
  }

  test("GH-468: --no-recalc on delete-cols and delete-rows preserves caches too") {
    val wb = structuralPoisonWorkbook()
    val rowsOut = tempXlsx()
    val colsOut = tempXlsx()

    WriteCommands
      .deleteRows(wb, wb.sheets.headOption, 20, 1, rowsOut, config, false, preserveCaches)
      .unsafeRunSync()
    WriteCommands
      .deleteColumns(wb, wb.sheets.headOption, "AA", 1, colsOut, config, false, preserveCaches)
      .unsafeRunSync()

    val rows = readBack(rowsOut)
    assertCachedOn(rows, "Data", ref"C1", 777.0)
    assertCachedOn(rows, "Other", ref"B1", 777.0)
    // C30 shifted up to C29 by the delete: a relocated cell never keeps its pre-edit cache.
    assertEquals(formulaOn(rows, "Data", ref"C29").cachedValue, None)
    val cols = readBack(colsOut)
    assertCachedOn(cols, "Data", ref"C1", 777.0)
    assertCachedOn(cols, "Other", ref"B1", 777.0)
    Files.deleteIfExists(rowsOut)
    Files.deleteIfExists(colsOut)
  }

  test("GH-468: --no-recalc never restores a cache onto a formula the edit REWROTE") {
    // The safety condition: a shifted formula's cached value belongs to the old text. F10's
    // `D10*2` becomes `D11*2` at F11 — restoring 999 there would invent a value no engine ever
    // computed, so the cell must stay uncached (Excel's own "dirty" state).
    val wb = Workbook(
      Sheet("Data")
        .put(ref"D10" -> 7)
        .put(ref"F10", CellValue.Formula("D10*2", Some(CellValue.Number(BigDecimal(999)))))
    )
    val out = tempXlsx()

    WriteCommands
      .insertRows(wb, wb.sheets.headOption, 5, 1, out, config, false, preserveCaches)
      .unsafeRunSync()

    val written = readBack(out)
    written.sheets.head.cells.get(ref"F11").map(_.value) match
      case Some(CellValue.Formula("D11*2", None, _)) => ()
      case other => fail(s"a rewritten formula must stay uncached, got $other")
    Files.deleteIfExists(out)
  }

  /** The formula cell at `at` on `sheet`, failing when the cell is absent or not a formula. */
  private def formulaOn(wb: Workbook, sheet: String, at: ARef): CellValue.Formula =
    wb.sheets.find(_.name.value == sheet).flatMap(_.cells.get(at)).map(_.value) match
      case Some(f: CellValue.Formula) => f
      case other => fail(s"expected a formula at $sheet!${at.toA1}, got $other")

  test("GH-468: --no-recalc never carries a POSITION-DEPENDENT cache across a row shift") {
    // `=ROW()` at C30 caches 30. An insert above it moves the cell to C31 without touching its
    // text, so the byte-identity check alone would happily re-cache 30 where the truth is 31.
    val wb = Workbook(
      Sheet("Data")
        .put(ref"C30", CellValue.Formula("ROW()", Some(CellValue.Number(BigDecimal(30)))))
    )
    val out = tempXlsx()

    WriteCommands
      .insertRows(wb, wb.sheets.headOption, 20, 1, out, config, false, preserveCaches)
      .unsafeRunSync()

    val moved = formulaOn(readBack(out), "Data", ref"C31")
    assertEquals(moved.expression, "ROW()")
    assertEquals(moved.cachedValue, None, s"a relocated =ROW() must not keep its old cache: $moved")
    Files.deleteIfExists(out)
  }

  test("GH-468: --no-recalc never carries a POSITION-DEPENDENT cache across a column shift") {
    // `=COLUMN()` at D30 caches 4; two inserted columns land it at F30, where the truth is 6.
    val wb = Workbook(
      Sheet("Data")
        .put(ref"D30", CellValue.Formula("COLUMN()", Some(CellValue.Number(BigDecimal(4)))))
    )
    val out = tempXlsx()

    WriteCommands
      .insertColumns(wb, wb.sheets.headOption, "A", 2, out, config, false, preserveCaches)
      .unsafeRunSync()

    val moved = formulaOn(readBack(out), "Data", ref"F30")
    assertEquals(moved.expression, "COLUMN()")
    assertEquals(moved.cachedValue, None, s"a relocated =COLUMN() must not keep its cache: $moved")
    Files.deleteIfExists(out)
  }

  test("GH-468: --no-recalc never restores a DYNAMIC-reference cache (its target may have moved)") {
    // F1 = INDIRECT("A30")*2 caches 10 off A30 = 5. The insert moves that content to A31 without
    // rewriting the string literal, so the formula now reads a blank: 10 is provably wrong.
    val wb = Workbook(
      Sheet("Data")
        .put(ref"A30" -> 5)
        .put(
          ref"F1",
          CellValue.Formula("INDIRECT(\"A30\")*2", Some(CellValue.Number(BigDecimal(10))))
        )
    )
    val out = tempXlsx()

    WriteCommands
      .insertRows(wb, wb.sheets.headOption, 20, 1, out, config, false, preserveCaches)
      .unsafeRunSync()

    val dynamic = formulaOn(readBack(out), "Data", ref"F1")
    assertEquals(
      dynamic.cachedValue,
      None,
      s"a dynamic reference must never keep a pre-edit cache: $dynamic"
    )
    Files.deleteIfExists(out)
  }

  /**
   * The dominant real shape: a contiguous block the edit cuts through. A1:A10 data, B1:B10 = An*2,
   * D1 = SUM(A1:A10) — every cache poisoned to 777 so any recalculation is visible.
   */
  private def contiguousBlockWorkbook(): Workbook =
    val poisoned = CellValue.Number(BigDecimal(777))
    val block = (1 to 10).foldLeft(Sheet("Data")) { (s, i) =>
      s.put(ARef.from0(0, i - 1), CellValue.Number(BigDecimal(i)))
        .put(ARef.from0(1, i - 1), CellValue.Formula(s"A$i*2", Some(poisoned)))
    }
    Workbook(block.put(ref"D1", CellValue.Formula("SUM(A1:A10)", Some(poisoned))))

  test("GH-468: --no-recalc through a contiguous block reports what it could NOT preserve") {
    // The row-20 fixture rewrites nothing; this one rewrites plenty. Every formula whose text the
    // edit changed is left uncached, and the summary must SAY SO rather than claim total
    // preservation.
    val wb = contiguousBlockWorkbook()
    val out = tempXlsx()

    val summary = WriteCommands
      .insertRows(wb, wb.sheets.headOption, 5, 1, out, config, false, preserveCaches)
      .unsafeRunSync()

    val written = readBack(out)
    // B1:B4 sit above the insert and still read A1:A4 — text unchanged, caches carried forward.
    (1 to 4).foreach(i => assertCachedOn(written, "Data", ARef.from0(1, i - 1), 777.0))
    // B5:B10 shifted to B6:B11 and their references were rewritten: no cache may be invented.
    (6 to 11).foreach { i =>
      val cell = formulaOn(written, "Data", ARef.from0(1, i - 1))
      assertEquals(cell.cachedValue, None, s"B$i must stay uncached, got $cell")
    }
    // D1 = SUM(A1:A10) became SUM(A1:A11): the old total answered a different question.
    assertEquals(formulaOn(written, "Data", ref"D1").cachedValue, None)
    assert(summary.contains("Recalculation skipped (--no-recalc)"), s"summary: $summary")
    assert(summary.contains("4 cached value(s) preserved"), s"summary: $summary")
    assert(summary.contains("7 formula(s)"), s"summary: $summary")
    assert(
      !summary.contains("every existing cached value preserved"),
      s"the structural arm must not claim total preservation: $summary"
    )
    Files.deleteIfExists(out)
  }

  test("GH-468: the unconditional preservation claim still holds for the NON-structural verbs") {
    val wb = splicedCacheWorkbook()
    val out = tempXlsx()

    val summary = WriteCommands
      .put(wb, wb.sheets.headOption, "A1", List("6"), out, config, policy = preserveCaches)
      .unsafeRunSync()

    assert(summary.contains("every existing cached value preserved"), s"summary: $summary")
    Files.deleteIfExists(out)
  }

  // ===== GH-496: --strict =====

  private def strictFailure(io: cats.effect.IO[String]): String =
    io.attempt.unsafeRunSync() match
      case Left(f: StrictFailure) => f.summary
      case Left(other) => fail(s"expected StrictFailure, got $other")
      case Right(summary) => fail(s"expected StrictFailure, got a clean summary: $summary")

  test("GH-496: --strict promotes recalc's formula errors to a failure (file still written)") {
    val wb = Workbook(Sheet("Data").put(ref"A1", CellValue.Formula("A1+1", None)))
    val lenientOut = tempXlsx()
    val strictOut = tempXlsx()

    // Default posture is unchanged: reported, never raised.
    val advisory = WriteCommands.recalc(wb, lenientOut, config).unsafeRunSync()
    assert(advisory.contains("Circular reference"), s"summary: $advisory")

    val failed = strictFailure(WriteCommands.recalc(wb, strictOut, config, policy = strictPolicy))
    assert(failed.contains("Circular reference"), s"the summary must survive verbatim: $failed")
    assert(failed.contains("STRICT FAILURE (--strict)"), s"summary: $failed")
    assert(failed.contains("1 formula evaluation error(s)"), s"summary: $failed")
    assert(java.nio.file.Files.size(strictOut) > 0L, "the output file is written either way")
    Files.deleteIfExists(lenientOut)
    Files.deleteIfExists(strictOut)
  }

  test("GH-496: --strict promotes the GH-454 non-convergence WARNING") {
    val wb = Workbook(
      Sheet("Data")
        .put(ref"B1", CellValue.Formula("1-B2"))
        .put(ref"B2", CellValue.Formula("1-B1"))
    ).withCalcPr(com.tjclp.xl.workbooks.CalcPr(iterativeCalculation = true, Some(5), None))
    val out = tempXlsx()

    val failed = strictFailure(WriteCommands.recalc(wb, out, config, policy = strictPolicy))
    assert(failed.contains("exhausted 5 round(s) without converging"), s"summary: $failed")
    Files.deleteIfExists(out)
  }

  test("GH-496: --strict promotes the GH-453 seed warnings of recalc --tables") {
    val out = tempXlsx()
    val failed = strictFailure(
      WriteCommands.recalc(circularTableWorkbook(None), out, config, false, true, strictPolicy)
    )
    assert(failed.contains("data table"), s"summary: $failed")
    assert(failed.contains("1 data-table seeding warning(s)"), s"summary: $failed")
    Files.deleteIfExists(out)
  }

  test("GH-496: --strict on a clean write is a no-op") {
    val wb = Workbook(
      Sheet("Data").put(ref"A1" -> 5).put(ref"C1", CellValue.Formula("A1*2", None))
    )
    val out = tempXlsx()
    val summary = WriteCommands.recalc(wb, out, config, policy = strictPolicy).unsafeRunSync()
    assert(summary.contains("Recalculated 1 formula"), s"summary: $summary")
    assert(!summary.contains("STRICT"), s"summary: $summary")
    Files.deleteIfExists(out)
  }

  test("GH-496: --strict fails a batch whose dirty cone contains a formula error") {
    val wb = Workbook(Sheet("Data"))
    val ops = writeOps("""[{"op":"putf","ref":"B1","value":"=B1+1"}]""")
    val out = tempXlsx()

    val failed = strictFailure(
      WriteCommands.batch(wb, wb.sheets.headOption, ops.toString, out, config, false, strictPolicy)
    )
    assert(failed.contains("Circular reference"), s"summary: $failed")
    assert(failed.contains("STRICT FAILURE (--strict)"), s"summary: $failed")
    Files.deleteIfExists(out)
    Files.deleteIfExists(ops)
  }

  test("GH-481: without the declaration the same batch keeps the circular-error posture") {
    val wb = circularRateWorkbook(None)
    val ops = writeOps("""[{"op":"put","ref":"A1","value":0.01}]""")
    val out = tempXlsx()

    val summary =
      WriteCommands.batch(wb, wb.sheets.headOption, ops.toString, out, config).unsafeRunSync()

    assert(summary.contains("Circular reference"), s"summary: $summary")
    assert(!summary.contains("converged in"), s"summary: $summary")
    assert(summary.contains("Saved:"), s"exit must stay clean: $summary")
    Files.deleteIfExists(out)
    Files.deleteIfExists(ops)
  }
