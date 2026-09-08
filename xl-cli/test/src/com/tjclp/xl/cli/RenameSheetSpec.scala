package com.tjclp.xl.cli

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cf.CfRule
import com.tjclp.xl.cli.commands.{SheetCommands, WriteCommands}
import com.tjclp.xl.cli.contract.{CliException, CliHarness, EnvelopeSchema, Location, TestFixtures}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.sheets.DataValidation

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

  test("rename-sheet verb: a new name another sheet has in a different case is DUPLICATE_SHEET") {
    // With tabs S and T, `rename-sheet T s` used to write two tabs Excel treats as one name and
    // rewrite `=T!A1*2` to `=s!A1*2` — which Excel resolves against the ORIGINAL S.
    val in = tmp("verb-dup-in")
    val out = tmp("verb-dup-out")
    Files.deleteIfExists(out)
    val wb = Workbook(
      Sheet("S").put(ref"A1", num(1)),
      Sheet("T").put(ref"A1", num(5)),
      Sheet("U").put(ref"A1", f("T!A1*2", Some(num(10))))
    )
    for
      _ <- excel.write(wb, in)
      read <- excel.read(in)
      attempt <- SheetCommands.renameSheet(read, "T", "s", out, config).attempt
      untouched <- excel.read(in)
    yield
      attempt match
        case Left(err: CliException) =>
          assertEquals(err.error.code, "DUPLICATE_SHEET")
          assertEquals(err.error.exitCode.code, 3)
          assert(err.getMessage.contains("Sheet 's' already exists"), err.getMessage)
        case Left(other) => fail(s"expected a CliException, got: $other")
        case Right(msg) => fail(s"expected a refusal, got: $msg")
      assert(!Files.exists(out), "nothing may be written when the rename is refused")
      assertEquals(untouched.sheets.map(_.name.value), Vector("S", "T", "U"))
      assertEquals(sheetNamed(untouched, "U")(ref"A1").value, f("T!A1*2", Some(num(10))))
  }

  test("add-sheet and copy-sheet refuse a name already used in any case as DUPLICATE_SHEET") {
    val in = tmp("verb-add-dup-in")
    val out = tmp("verb-add-dup-out")
    Files.deleteIfExists(out)
    for
      _ <- excel.write(fixture, in)
      read <- excel.read(in)
      added <- SheetCommands.addSheet(read, "sheet1", None, None, out, config).attempt
      copied <- SheetCommands.copySheet(read, "Sheet1", "NOTES", out, config).attempt
    yield
      for attempt <- Vector(added, copied) do
        attempt match
          case Left(err: CliException) => assertEquals(err.error.code, "DUPLICATE_SHEET")
          case other => fail(s"expected DUPLICATE_SHEET, got: $other")
      assert(!Files.exists(out), "nothing may be written when the add or copy is refused")
  }

  test("rename-sheet verb: an unrewritable dependent is FORMULA_ERROR at its location (GH-608)") {
    // `Summary!I23 = IF(ZZZNOTAFUNC(1)=1,'M&A'!I12,0)`: mentions the sheet, cannot be parsed. The
    // refusal is a data condition, not a defect: typed code, exit 3, the cell, a human diagnostic.
    val in = tmp("verb-unrewritable-in")
    val out = tmp("verb-unrewritable-out")
    Files.deleteIfExists(out)
    for
      _ <- excel.write(TestFixtures.qualifiedBook(), in)
      read <- excel.read(in)
      attempt <- SheetCommands.renameSheet(read, "M&A", "MandA", out, config).attempt
    yield
      attempt match
        case Left(err: CliException) =>
          assertEquals(err.error.code, "FORMULA_ERROR")
          assertEquals(err.error.exitCode.code, 3)
          assertEquals(
            err.error.location,
            Some(Location(None, Some("Summary"), Some("I23"), None))
          )
          assert(err.error.message.contains("Summary!I23"), err.error.message)
          assert(err.error.message.contains("Unknown function 'ZZZNOTAFUNC'"), err.error.message)
          assert(!err.error.message.contains("UnknownFunction("), err.error.message)
          assert(!err.error.message.contains("List("), err.error.message)
          assertEquals(
            err.error.hint,
            Some("fix or replace the formula at Summary!I23 before renaming")
          )
        case Left(other) => fail(s"expected a CliException, got: $other")
        case Right(msg) => fail(s"expected a refusal, got: $msg")
      assert(!Files.exists(out), "nothing may be written when the rename is refused")
  }

  test("rename-sheet refusal envelopes: the verb is FORMULA_ERROR, the batch op BATCH_OP_FAILED") {
    val in = tmp("json-unrewritable-in")
    val out = tmp("json-unrewritable-out")
    Files.deleteIfExists(out)
    val ops = opsFile("""[{"op":"rename-sheet","from":"M&A","to":"MandA"}]""")
    for
      _ <- excel.write(TestFixtures.qualifiedBook(), in)
      read <- excel.read(in)
      direct <- WriteCommands.batch(read, read.sheets.headOption, ops.toString, out, config).attempt
      verb <- CliHarness.run(
        "-f",
        in.toString,
        "-o",
        out.toString,
        "--json",
        "rename-sheet",
        "M&A",
        "MandA"
      )
      batch <- CliHarness.run(
        "-f",
        in.toString,
        "-o",
        out.toString,
        "--json",
        "batch",
        ops.toString
      )
    yield
      assertEquals(verb.exit, 3, verb.stderr)
      val envelope = ujson.read(verb.stdout)
      EnvelopeSchema.assertValid(envelope)
      val error = envelope("error")
      assertEquals(error("code"), ujson.Str("FORMULA_ERROR"))
      assertEquals(error("location")("sheet"), ujson.Str("Summary"))
      assertEquals(error("location")("ref"), ujson.Str("I23"))
      assert(error("message").str.contains("Unknown function 'ZZZNOTAFUNC'"), error("message").str)
      assert(!error("message").str.contains("UnknownFunction("), error("message").str)
      assertEquals(
        error("hint"),
        ujson.Str("fix or replace the formula at Summary!I23 before renaming")
      )
      assert(!Files.exists(out), "nothing may be written when the rename is refused")
      // the batch op: BATCH_OP_FAILED names the op index; the verb's cause, hint and location ride
      // along (the op raises the same typed refusal, not a plain exception)
      assertEquals(batch.exit, 3, batch.stderr)
      val batchError = ujson.read(batch.stdout)("error")
      assertEquals(batchError("code"), ujson.Str("BATCH_OP_FAILED"))
      assert(
        batchError("message").str.startsWith("Object 1 (rename-sheet): "),
        batchError("message").str
      )
      assert(batchError("message").str.contains("Summary!I23"), batchError("message").str)
      assert(!batchError("message").str.contains("UnknownFunction("), batchError("message").str)
      assertEquals(
        batchError("hint"),
        ujson.Str("fix or replace the formula at Summary!I23 before renaming")
      )
      assertEquals(batchError("location")("sheet"), ujson.Str("Summary"))
      assertEquals(batchError("location")("ref"), ujson.Str("I23"))
      assertEquals(batchError("location")("opIndex"), ujson.Num(1))
      direct match
        case Left(err: CliException) =>
          assertEquals(err.error.code, "BATCH_OP_FAILED")
          assertEquals(err.error.cause.map(_.code), Some("FORMULA_ERROR"))
          assertEquals(
            err.error.hint,
            Some("fix or replace the formula at Summary!I23 before renaming")
          )
          assertEquals(
            err.error.location,
            Some(Location(None, Some("Summary"), Some("I23"), Some(1)))
          )
        case other => fail(s"expected a BATCH_OP_FAILED CliException, got: $other")
      assert(!Files.exists(out), "nothing may be written when the batch is refused")
  }

  test(
    "rename-sheet refusal on a conditional format or data validation locates the sheet, no ref"
  ) {
    // GH-608: `location.sheet` names the sheet whose rule mentions the renamed sheet; `location.ref`
    // is null (a rule is not a cell) — the branch an agent takes on the envelope depends on it.
    def book(summary: Sheet): Workbook = Workbook(Sheet("M&A").put(ref"I12", num(7)), summary)
    val cfBook = book(
      Sheet("Summary")
        .conditionalFormat(ref"A1:A3", CfRule.Expression("ZZZNOTAFUNC('M&A'!I12)", None, 1))
    )
    val dvBook = book(
      Sheet("Summary")
        .withDataValidation(ref"B1:B3", DataValidation.custom("ZZZNOTAFUNC('M&A'!I12)"))
    )
    def refusal(tag: String, wb: Workbook): IO[(ujson.Value, java.nio.file.Path)] =
      val in = tmp(s"$tag-in")
      val out = tmp(s"$tag-out")
      Files.deleteIfExists(out)
      for
        _ <- excel.write(wb, in)
        run <- CliHarness.run(
          "-f",
          in.toString,
          "-o",
          out.toString,
          "--json",
          "rename-sheet",
          "M&A",
          "MandA"
        )
      yield
        assertEquals(run.exit, 3, run.stderr)
        val envelope = ujson.read(run.stdout)
        EnvelopeSchema.assertValid(envelope)
        (envelope("error"), out)
    for
      (cf, cfOut) <- refusal("cf", cfBook)
      (dv, dvOut) <- refusal("dv", dvBook)
    yield
      for (error, kind) <- Vector((cf, "conditional format"), (dv, "data validation")) do
        assertEquals(error("code"), ujson.Str("FORMULA_ERROR"), kind)
        assertEquals(error("location")("sheet"), ujson.Str("Summary"), kind)
        assertEquals(error("location")("ref"), ujson.Null, kind)
        assert(error("message").str.contains(s"Summary!$kind: "), error("message").str)
        assert(
          error("message").str.contains("Unknown function 'ZZZNOTAFUNC'"),
          error("message").str
        )
        assert(!error("message").str.contains("UnknownFunction("), error("message").str)
        assertEquals(
          error("hint"),
          ujson.Str(
            s"fix or remove the $kind on sheet Summary that names the renamed sheet before renaming"
          ),
          kind
        )
      assert(!Files.exists(cfOut) && !Files.exists(dvOut), "nothing may be written on a refusal")
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
