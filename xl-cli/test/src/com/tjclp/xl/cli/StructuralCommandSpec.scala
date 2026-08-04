package com.tjclp.xl.cli

import java.nio.file.Files

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * Integration tests for the `insert-rows` / `delete-rows` / `insert-cols` / `delete-cols` CLI
 * commands. These round-trip through a real `.xlsx` so they exercise the full path INCLUDING the
 * SourceContext interaction — i.e. they regress the bug where a structural edit on a freshly-read
 * (clean) workbook was silently dropped by the writer's verbatim-copy fast-path.
 */
@SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
class StructuralCommandSpec extends CatsEffectSuite:

  private val excel = ExcelIO.instance[IO]
  private val config = WriterConfig.default

  private def tmp(tag: String) =
    val p = Files.createTempFile(s"struct-$tag-", ".xlsx")
    p.toFile.deleteOnExit()
    p

  test("insert-rows: read -> shift cells & rewrite formulas -> write (no stale verbatim copy)") {
    val wb = Workbook(
      Vector(
        Sheet("S")
          .put(ref"A1", CellValue.Number(10))
          .put(ref"A3", CellValue.Number(30))
          .put(ref"B1", CellValue.Formula("=A1+A3", None))
      )
    )
    val in = tmp("in")
    val out = tmp("out")
    for
      _ <- excel.write(wb, in)
      read <- excel.read(in) // sourceContext present & clean
      _ <- WriteCommands.insertRows(read, read.sheets.headOption, 2, 1, out, config)
      result <- excel.read(out)
    yield
      val s = result.sheets.head
      // GH-427: equals-free canonical form; the command's recalc re-bakes the cache (10+30).
      assertEquals(s(ref"B1").value, CellValue.Formula("A1+A4", Some(CellValue.Number(40))))
      assertEquals(s(ref"A4").value, CellValue.Number(30)) // cell shifted down
      assertEquals(s(ref"A1").value, CellValue.Number(10)) // unchanged
  }

  test("insert-rows recomputes cached <v> and emits a single equals-free <f>") {
    // Deliberately WRONG cache (999): only the command's recalc can produce 4, so the file
    // assertion below distinguishes fresh-correct from stale-preserved.
    val stale = Some(CellValue.Number(BigDecimal(999)))
    val wb = Workbook(
      Vector(
        Sheet("S")
          .put(ref"A1", CellValue.Number(2))
          .put(ref"B1", CellValue.Formula("A1*2", stale))
      )
    )
    val in = tmp("in427")
    val out = tmp("out427")
    for
      _ <- excel.write(wb, in)
      read <- excel.read(in)
      _ <- WriteCommands.insertRows(read, read.sheets.headOption, 2, 1, out, config)
      result <- excel.read(out)
      raw <- IO.blocking {
        val zip = new java.util.zip.ZipFile(out.toFile)
        try
          val entry = zip.getEntry("xl/worksheets/sheet1.xml")
          new String(zip.getInputStream(entry).readAllBytes(), "UTF-8")
        finally zip.close()
      }
    yield
      // Model: equals-free text; the structural edit invalidated the stale cache and the
      // command's global recalc re-baked the true value.
      assertEquals(
        result.sheets.head(ref"B1").value,
        CellValue.Formula("A1*2", Some(CellValue.Number(4)))
      )
      // File: <f> carries no '=' and <v> is the recomputed value, never the stale one.
      assert(raw.contains("<f>A1*2</f>"), s"expected equals-free <f> in: $raw")
      assert(!raw.contains("<f>="), s"leading '=' must not land inside <f>: $raw")
      assert(raw.contains("<v>4</v>"), s"recomputed <v> must be present: $raw")
      assert(!raw.contains("<v>999</v>"), s"stale <v> must not survive: $raw")
  }

  test("delete-rows re-bakes a shrinking SUM to its new true value") {
    val wb = Workbook(
      Vector(
        Sheet("S")
          .put(ref"A1", CellValue.Number(1))
          .put(ref"A2", CellValue.Number(2))
          .put(ref"A3", CellValue.Number(3))
          .put(ref"C1", CellValue.Formula("SUM(A1:A3)", Some(CellValue.Number(6))))
      )
    )
    val in = tmp("inShrink")
    val out = tmp("outShrink")
    for
      _ <- excel.write(wb, in)
      read <- excel.read(in)
      _ <- WriteCommands.deleteRows(read, read.sheets.headOption, 2, 1, out, config)
      result <- excel.read(out)
    yield
      // A2 deleted: SUM(A1:A2) now spans A1=1 and A2=3 (was A3). The old cache (6) would be
      // silently wrong; the command's recalc writes the true 4.
      assertEquals(
        result.sheets.head(ref"C1").value,
        CellValue.Formula("SUM(A1:A2)", Some(CellValue.Number(4)))
      )
  }

  test("delete-rows: a reference into the deleted row becomes #REF!") {
    val wb = Workbook(
      Vector(
        Sheet("S")
          .put(ref"A4", CellValue.Number(40))
          .put(ref"B1", CellValue.Formula("=A4", None))
      )
    )
    val in = tmp("in2")
    val out = tmp("out2")
    for
      _ <- excel.write(wb, in)
      read <- excel.read(in)
      _ <- WriteCommands.deleteRows(read, read.sheets.headOption, 4, 1, out, config)
      result <- excel.read(out)
    yield assertEquals(result.sheets.head(ref"B1").value, CellValue.Error(CellError.Ref))
  }

  test("insert-cols: column references at/after the insertion point shift right") {
    val wb = Workbook(
      Vector(
        Sheet("S")
          .put(ref"A1", CellValue.Number(1))
          .put(ref"A2", CellValue.Formula("=C1", None))
      )
    )
    val in = tmp("in3")
    val out = tmp("out3")
    for
      _ <- excel.write(wb, in)
      read <- excel.read(in)
      _ <- WriteCommands.insertColumns(read, read.sheets.headOption, "B", 1, out, config)
      result <- excel.read(out)
    yield assertEquals(
      result.sheets.head(ref"A2").value,
      // C1 -> D1 (GH-427: equals-free canonical form); the command's recalc caches the
      // empty-reference result (0, Excel-consistent).
      CellValue.Formula("D1", Some(CellValue.Number(0)))
    )
  }

  test("GH-429: insert-rows moves DV + print area + table + autoFilter (the field repro)") {
    import com.tjclp.xl.sheets.{AutoFilterState, PageSetup}
    import com.tjclp.xl.tables.TableSpec
    val table = TableSpec
      .fromColumnNames("T1", "T1", ref"A1:C5", Vector("Region", "Product", "Units"))
      .fold(e => fail(s"table: $e"), identity)
    val sheet = Sheet("Alpha")
      .put(ref"A1", CellValue.Text("Region"))
      .put(ref"B1", CellValue.Text("Product"))
      .put(ref"C1", CellValue.Text("Units"))
      // the Excel field shape: list DV with prompt/error flags stamped
      .withDataValidation(
        ref"H5:H10",
        DataValidation.listOf("yes", "no").withPrompt("Pick", "yes or no")
      )
      .withTable(table)
      .copy(
        pageSetup = Some(PageSetup(printArea = Some(ref"A1:D10"))),
        autoFilter = Some(AutoFilterState.Ranged(ref"A1:C5"))
      )
    val in = tmp("in429")
    val out = tmp("out429")
    def zipEntry(p: java.nio.file.Path, name: String): String =
      val zip = new java.util.zip.ZipFile(p.toFile)
      try
        val entry = Option(zip.getEntry(name)).getOrElse(fail(s"missing $name"))
        new String(zip.getInputStream(entry).readAllBytes(), "UTF-8")
      finally zip.close()
    for
      _ <- excel.write(Workbook(sheet), in)
      read <- excel.read(in)
      _ <- WriteCommands.insertRows(read, read.sheets.headOption, 2, 2, out, config)
      result <- excel.read(out)
    yield
      val s = result.sheets.head
      // model: all four range-bearing features moved with the data
      assertEquals(s.typedDataValidations.flatMap(_.ranges.map(_.toA1)), Vector("H7:H12"))
      assertEquals(s.pageSetup.flatMap(_.printArea).map(_.toA1), Some("A1:D12"))
      assertEquals(s.tables.get("T1").map(_.range.toA1), Some("A1:C7"))
      assertEquals(s.autoFilter, Some(AutoFilterState.Ranged(ref"A1:C7": CellRange)))
      // file: zip-level proof for the adversarial reader
      val sheetXml = zipEntry(out, "xl/worksheets/sheet1.xml")
      assert(sheetXml.contains("sqref=\"H7:H12\""), sheetXml)
      assert(sheetXml.contains("<autoFilter ref=\"A1:C7\""), sheetXml)
      assert(zipEntry(out, "xl/tables/table1.xml").contains("ref=\"A1:C7\""))
      assert(zipEntry(out, "xl/workbook.xml").contains("Alpha!$A$1:$D$12"))
  }

  test("delete-cols: range form C:E removes the whole span (GH-129)") {
    val wb = Workbook(
      Vector(
        Sheet("S")
          .put(ref"A1", CellValue.Number(1))
          .put(ref"B1", CellValue.Number(2))
          .put(ref"C1", CellValue.Number(3))
          .put(ref"D1", CellValue.Number(4))
          .put(ref"E1", CellValue.Number(5))
          .put(ref"F1", CellValue.Number(6))
      )
    )
    val in = tmp("in4")
    val out = tmp("out4")
    for
      _ <- excel.write(wb, in)
      read <- excel.read(in)
      // count arg is ignored when a range is given; C:E deletes 3 columns
      _ <- WriteCommands.deleteColumns(read, read.sheets.headOption, "C:E", 1, out, config)
      result <- excel.read(out)
    yield
      val s = result.sheets.head
      assertEquals(s(ref"A1").value, CellValue.Number(1)) // unchanged
      assertEquals(s(ref"B1").value, CellValue.Number(2)) // unchanged
      assertEquals(s(ref"C1").value, CellValue.Number(6)) // F shifted left into C
      assert(!s.contains(ref"D1")) // only 3 columns remain
  }

  test("delete-cols: reversed range E:C normalizes to the same span") {
    val wb = Workbook(
      Vector(
        Sheet("S")
          .put(ref"A1", CellValue.Number(1))
          .put(ref"B1", CellValue.Number(2))
          .put(ref"C1", CellValue.Number(3))
          .put(ref"D1", CellValue.Number(4))
          .put(ref"E1", CellValue.Number(5))
          .put(ref"F1", CellValue.Number(6))
      )
    )
    val in = tmp("in5")
    val out = tmp("out5")
    for
      _ <- excel.write(wb, in)
      read <- excel.read(in)
      _ <- WriteCommands.deleteColumns(read, read.sheets.headOption, "E:C", 1, out, config)
      result <- excel.read(out)
    yield
      val s = result.sheets.head
      assertEquals(s(ref"C1").value, CellValue.Number(6)) // identical to C:E
      assert(!s.contains(ref"D1"))
  }

  test("delete-cols: malformed ranges are rejected (total Left, never crash)") {
    val wb = Workbook(Vector(Sheet("S").put(ref"A1", CellValue.Number(1))))
    val in = tmp("in6")
    val out = tmp("out6")
    for
      _ <- excel.write(wb, in)
      read <- excel.read(in)
      r1 <- WriteCommands
        .deleteColumns(read, read.sheets.headOption, "C:E:F", 1, out, config)
        .attempt
      r2 <- WriteCommands.deleteColumns(read, read.sheets.headOption, "C:", 1, out, config).attempt
    yield
      assert(r1.isLeft) // three-part range
      assert(r2.isLeft) // trailing colon (empty endpoint)
  }

  // ===== GH-472: refuse inserts that would shift populated cells past the sheet edge =====

  test("GH-472: insert-rows refusing an over-the-edge shift leaves the file byte-unchanged") {
    // The exact field repro: data at rows 1048570-1048572, insert 20 rows at 5. 0.19.0 exited 0
    // and wrote cells past row 1048576 (plus a <dimension> past the cap) — Excel refuses the file.
    val wb = Workbook(
      Vector(
        Sheet("S")
          .put(ref"A1048570", CellValue.Number(1))
          .put(ref"B1048570", CellValue.Number(2))
          .put(ref"A1048572", CellValue.Number(3))
          .put(ref"B1048572", CellValue.Number(4))
      )
    )
    val in = tmp("in-gh472")
    for
      _ <- excel.write(wb, in)
      before <- IO(Files.readAllBytes(in).toVector)
      read <- excel.read(in)
      // in-place edit (out == in): a refusal must not touch the file
      r <- WriteCommands.insertRows(read, read.sheets.headOption, 5, 20, in, config).attempt
      after <- IO(Files.readAllBytes(in).toVector)
    yield
      r match
        case Left(e) =>
          assert(e.getMessage.contains("1048572"), e.getMessage) // offending populated row
          assert(e.getMessage.contains("1048576"), e.getMessage) // the cap
        case Right(msg) => fail(s"expected refusal, got success: $msg")
      assertEquals(after, before) // file byte-identical
  }

  test("GH-472: insert-rows just under the cap still succeeds (lands exactly on 1048576)") {
    val wb = Workbook(Vector(Sheet("S").put(ref"A1048556", CellValue.Number(42))))
    val in = tmp("in-gh472b")
    val out = tmp("out-gh472b")
    for
      _ <- excel.write(wb, in)
      read <- excel.read(in)
      _ <- WriteCommands.insertRows(read, read.sheets.headOption, 5, 20, out, config)
      result <- excel.read(out)
    yield
      val s = result.sheets.head
      assertEquals(s(ref"A1048576").value, CellValue.Number(42))
      assert(!s.contains(ref"A1048556"))
  }

  test("GH-472: insert-cols refuses a shift past column XFD and leaves the file unchanged") {
    val wb = Workbook(Vector(Sheet("S").put(ref"XFC1", CellValue.Number(1))))
    val in = tmp("in-gh472c")
    for
      _ <- excel.write(wb, in)
      before <- IO(Files.readAllBytes(in).toVector)
      read <- excel.read(in)
      r <- WriteCommands.insertColumns(read, read.sheets.headOption, "C", 20, in, config).attempt
      after <- IO(Files.readAllBytes(in).toVector)
    yield
      r match
        case Left(e) =>
          assert(e.getMessage.contains("XFC"), e.getMessage)
          assert(e.getMessage.contains("XFD"), e.getMessage)
        case Right(msg) => fail(s"expected refusal, got success: $msg")
      assertEquals(after, before)
  }

  test("GH-472: insert-cols just under the cap still succeeds (lands exactly on XFD)") {
    val wb = Workbook(Vector(Sheet("S").put(ref"XFC1", CellValue.Number(7))))
    val in = tmp("in-gh472d")
    val out = tmp("out-gh472d")
    for
      _ <- excel.write(wb, in)
      read <- excel.read(in)
      _ <- WriteCommands.insertColumns(read, read.sheets.headOption, "C", 1, out, config)
      result <- excel.read(out)
    yield assertEquals(result.sheets.head(ref"XFD1").value, CellValue.Number(7))
  }
