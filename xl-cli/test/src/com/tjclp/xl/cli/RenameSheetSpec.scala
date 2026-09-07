package com.tjclp.xl.cli

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.{SheetCommands, WriteCommands}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * GH-559: `rename-sheet` (the verb and the batch op) must rewrite every `Sheet!ref` in other
 * sheets' formulas, not just the tab name — through a REAL file, because the surgical writer copies
 * unmodified sheets byte-for-byte from the source zip and an in-memory-only rewrite would still
 * write `Sheet1!A1` (the #559 symptom: the file lints clean and Excel shows `#REF!`).
 */
@SuppressWarnings(Array("org.wartremover.warts.IterableOps"))
class RenameSheetSpec extends CatsEffectSuite:

  private val excel = ExcelIO.instance[IO]
  private val config = WriterConfig.default

  private def num(i: Int): CellValue = CellValue.Number(BigDecimal(i))
  private def f(text: String, cached: Option[CellValue]): CellValue =
    CellValue.Formula(text, cached)

  private def tmp(tag: String): Path =
    val p = Files.createTempFile(s"rename-$tag-", ".xlsx")
    p.toFile.deleteOnExit()
    p

  private def opsFile(json: String): Path =
    val p = Files.createTempFile("rename-ops-", ".json")
    p.toFile.deleteOnExit()
    Files.write(p, json.getBytes(StandardCharsets.UTF_8))
    p

  private def fixture: Workbook =
    Workbook(
      Sheet("Sheet1").put(ref"A1", num(5)),
      Sheet("Sheet2")
        .put(ref"A1", f("Sheet1!A1*2", Some(num(10))))
        .put(ref"B1", f("SUM(Sheet1!A1:A1)", Some(num(5))))
        .put(ref"C1", f("A1*2", Some(num(20)))),
      Sheet("Notes").put(ref"A1", CellValue.Text("no formulas here"))
    ).withDefinedName("Total", "Sheet1!$A$1")

  private def zipEntryText(path: Path, entryName: String): IO[String] = IO.blocking {
    val zip = new java.util.zip.ZipFile(path.toFile)
    try
      Option(zip.getEntry(entryName)) match
        case Some(entry) =>
          new String(zip.getInputStream(entry).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"missing $entryName in $path")
    finally zip.close()
  }

  private def sheetNamed(wb: Workbook, name: String): Sheet =
    wb.sheets.find(_.name.value == name).getOrElse(fail(s"missing sheet $name in ${wb.sheetNames}"))

  private def assertRewritten(result: Workbook, sheet2Xml: String): Unit =
    assertEquals(result.sheets.map(_.name.value), Vector("Data", "Sheet2", "Notes"))
    val sheet2 = sheetNamed(result, "Sheet2")
    assertEquals(sheet2(ref"A1").value, f("Data!A1*2", Some(num(10))))
    assertEquals(sheet2(ref"B1").value, f("SUM(Data!A1:A1)", Some(num(5))))
    assertEquals(sheet2(ref"C1").value, f("A1*2", Some(num(20))))
    assertEquals(result.metadata.definedNames.map(_.formula), Vector("Data!$A$1"))
    assert(sheet2Xml.contains("Data!A1*2"), sheet2Xml)
    assert(!sheet2Xml.contains("Sheet1!"), sheet2Xml)

  test(
    "rename-sheet verb: read -> rename -> write rewrites dependents in the FILE; summary counts"
  ) {
    val in = tmp("verb-in")
    val out = tmp("verb-out")
    for
      _ <- excel.write(fixture, in)
      read <- excel.read(in)
      message <- SheetCommands.renameSheet(read, "Sheet1", "Data", out, config)
      result <- excel.read(out)
      sheet2Xml <- zipEntryText(out, "xl/worksheets/sheet2.xml")
    yield
      assert(message.startsWith("Renamed: Sheet1 → Data; 2 formula(s) rewritten\n"), message)
      assertRewritten(result, sheet2Xml)
  }

  test("rename-sheet verb: no count suffix when nothing referenced the sheet") {
    val in = tmp("verb-notes-in")
    val out = tmp("verb-notes-out")
    for
      _ <- excel.write(fixture, in)
      read <- excel.read(in)
      message <- SheetCommands.renameSheet(read, "Notes", "Memo", out, config)
      result <- excel.read(out)
    yield
      assert(message.startsWith("Renamed: Notes → Memo\n"), message)
      assertEquals(result.sheets.map(_.name.value), Vector("Sheet1", "Sheet2", "Memo"))
      assertEquals(sheetNamed(result, "Sheet2")(ref"A1").value, f("Sheet1!A1*2", Some(num(10))))
  }

  test("rename-sheet verb: a new name that needs quoting is quoted") {
    val in = tmp("verb-quote-in")
    val out = tmp("verb-quote-out")
    for
      _ <- excel.write(fixture, in)
      read <- excel.read(in)
      _ <- SheetCommands.renameSheet(read, "Sheet1", "Q1 Data", out, config)
      result <- excel.read(out)
    yield
      assertEquals(sheetNamed(result, "Sheet2")(ref"A1").value, f("'Q1 Data'!A1*2", Some(num(10))))
      assertEquals(result.metadata.definedNames.map(_.formula), Vector("'Q1 Data'!$A$1"))
  }

  test("rename-sheet verb: refusal on an unparsable dependent leaves no output behind") {
    val in = tmp("verb-refuse-in")
    val out = tmp("verb-refuse-out")
    Files.deleteIfExists(out)
    val broken = fixture.put(sheetNamed(fixture, "Sheet2").put(ref"D1", f("Sheet1!A1+", None)))
    for
      _ <- excel.write(broken, in)
      read <- excel.read(in)
      attempt <- SheetCommands.renameSheet(read, "Sheet1", "Data", out, config).attempt
    yield
      attempt match
        case Left(err) => assert(err.getMessage.contains("Sheet2!D1"), err.getMessage)
        case Right(msg) => fail(s"expected a refusal, got: $msg")
      assert(!Files.exists(out), "nothing may be written when the rename is refused")
  }

  test("batch rename-sheet: rewrites dependents in the FILE without recalculating") {
    val in = tmp("batch-in")
    val out = tmp("batch-out")
    val ops = opsFile("""[{"op":"rename-sheet","from":"Sheet1","to":"Data"}]""")
    for
      _ <- excel.write(fixture, in)
      read <- excel.read(in)
      summary <- WriteCommands.batch(read, read.sheets.headOption, ops.toString, out, config)
      result <- excel.read(out)
      sheet2Xml <- zipEntryText(out, "xl/worksheets/sheet2.xml")
    yield
      assert(summary.contains("RENAME-SHEET Sheet1 -> Data"), summary)
      // a rename changes no value: the presentation-only classification (no recalculation) holds
      assert(!summary.contains("Recalculated"), summary)
      assertRewritten(result, sheet2Xml)
  }

  test("batch rename-sheet: the pre-rename dependents' caches are the ones on disk") {
    // Wrong caches prove preservation: a recalculation would replace 999 with 10.
    val in = tmp("batch-cache-in")
    val out = tmp("batch-cache-out")
    val ops = opsFile("""[{"op":"rename-sheet","from":"Sheet1","to":"Data"}]""")
    val stale =
      fixture.put(sheetNamed(fixture, "Sheet2").put(ref"A1", f("Sheet1!A1*2", Some(num(999)))))
    for
      _ <- excel.write(stale, in)
      read <- excel.read(in)
      _ <- WriteCommands.batch(read, read.sheets.headOption, ops.toString, out, config)
      result <- excel.read(out)
    yield assertEquals(sheetNamed(result, "Sheet2")(ref"A1").value, f("Data!A1*2", Some(num(999))))
  }
